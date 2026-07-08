import sys
import os
import asyncio
import threading
import unicodedata
import subprocess
import argparse
import random
import json
try:                       # tkinter 仅本工具自带 GUI 用；打包成 PySide6 exe 时可能没带全，
    import tkinter as tk   # 缺失也不能影响 --cli 合成与被 import 读配置。
    from tkinter import filedialog, ttk, messagebox
except Exception:
    tk = ttk = filedialog = messagebox = None
import edge_tts

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ------------------------------------------------------------------------
# 1. Language & Voice Registry
# ------------------------------------------------------------------------
def load_lang_config():
    json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "all_voices.json")
    
    # Preferred order to maintain compatibility with original IDs (0-27)
    preferred = [
        ("en-US", "English"), ("fr-FR", "French"), ("es-ES", "Spanish"),
        ("de-DE", "German"), ("it-IT", "Italian"), ("ar-SA", "Arabic"),
        ("fil-PH", "Filipino"), ("hi-IN", "Hindi"), ("ms-MY", "Malay"),
        ("pl-PL", "Polish"), ("th-TH", "Thai"), ("tr-TR", "Turkish"),
        ("zh-HK", "Yueyu-Cantonese"), ("ru-RU", "Russian"), ("ja-JP", "Japanese"),
        ("pt-PT", "Portuguese"), ("id-ID", "Indonesian"), ("fa-IR", "Persian"),
        ("sv-SE", "Swedish"), ("nl-NL", "Dutch"), ("da-DK", "Danish"),
        ("nb-NO", "Norwegian"), ("vi-VN", "Vietnamese"), ("uz-UZ", "Uzbek"),
        ("kk-KZ", "Kazakh"), ("he-IL", "Hebrew"), ("ko-KR", "Korean"),
        ("sl-SI", "Slovenian")
    ]

    try:
        if not os.path.exists(json_path):
            # Fallback to a very minimal config if file is missing
            return {f"{name} ({i})": (str(i), []) for i, (loc, name) in enumerate(preferred)}

        with open(json_path, "r", encoding="utf-8") as f:
            voices_data = json.load(f)
        
        locales_map = {}
        for v in voices_data:
            loc = v["Locale"]
            if loc not in locales_map:
                locales_map[loc] = []
            locales_map[loc].append((v["ShortName"], v["Gender"]))
            
        config = {}
        used_locales = set()
        
        # 1. Map preferred locales to IDs 0-27
        for i, (loc, name) in enumerate(preferred):
            if loc in locales_map:
                config[f"{name} ({i})"] = (str(i), locales_map[loc])
                used_locales.add(loc)
            else:
                config[f"{name} ({i})"] = (str(i), [])
        
        # 2. Add other locales found in the JSON
        next_id = len(preferred)
        for loc in sorted(locales_map.keys()):
            if loc not in used_locales:
                # Use locale as name if not in preferred
                config[f"{loc} ({next_id})"] = (str(next_id), locales_map[loc])
                next_id += 1
                
        return config
    except Exception as e:
        print(f"[ERR] Failed to load voices from {json_path}: {e}")
        return {f"{name} ({i})": (str(i), []) for i, (loc, name) in enumerate(preferred)}

LANG_CONFIG = load_lang_config()

# Character-level tokenization for CJK languages
# We dynamically find the IDs for zh-CN, zh-HK, ja-JP, ko-KR
def get_char_level_langs(config):
    target_locales = {"zh-CN", "zh-HK", "ja-JP", "ko-KR"}
    ids = set()
    for key, (lang_id, _) in config.items():
        # Check if the key corresponds to a target locale
        # Key format is "Name (ID)"
        # This is a bit brittle, but we can check the ID in the config
        # Actually, it's better to store the locale in the config too, 
        # but for now let's just use a fixed mapping or search.
        pass
    # For now, let's keep it simple and just use the IDs we know or find them
    found_ids = set()
    for key, (val_id, voices) in config.items():
        if voices and any(v[0].startswith(("zh-", "ja-", "ko-")) for v in voices):
            found_ids.add(val_id)
    return found_ids

CHAR_LEVEL_LANGS = get_char_level_langs(LANG_CONFIG)


# ------------------------------------------------------------------------
# 2. Engine helpers
# ------------------------------------------------------------------------
def get_ffmpeg_path():
    base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    exe_name = "ffmpeg.exe" if os.name == "nt" else "ffmpeg"
    for candidate in [
        os.path.join(base_dir, exe_name),
        os.path.join(getattr(sys, "_MEIPASS", ""), exe_name),
        os.path.join(os.getcwd(), exe_name),
    ]:
        if candidate and os.path.exists(candidate):
            return candidate
    raise FileNotFoundError(
        f"Critical Binary Missing: cannot find {exe_name}.\nChecked: {base_dir}"
    )


def parse_line(raw: str):
    """
    Split a corpus line into (text, pause_ms).
    Syntax: "sentence text||500"  →  ("sentence text", 500)
    Lines without '||' get 0 ms tail silence.
    """
    if "||" in raw:
        parts = raw.rsplit("||", 1)
        try:
            return parts[0].strip(), max(0, int(parts[1].strip()))
        except ValueError:
            pass
    return raw.strip(), 0


def pick_rate(min_pct: int, max_pct: int) -> str:
    """Pick a random speed rate within [min_pct, max_pct] percent."""
    val = random.randint(min(min_pct, max_pct), max(min_pct, max_pct))
    return f"+{val}%" if val >= 0 else f"{val}%"


def assign_varied_indices(total: int, ratio_pct: int) -> set:
    """Return a set of indices that should be speed-varied."""
    n = max(0, min(total, round(total * ratio_pct / 100)))
    return set(random.sample(range(total), n)) if n > 0 else set()


def clean_text_for_mlf(text, lang_id):
    text = unicodedata.normalize("NFKC", text.lower())
    chars = []
    for c in text:
        if c in {"'", "’", "‘", "`", "´"}:
            chars.append("'")
        elif not unicodedata.category(c).startswith("P"):
            chars.append(c)
    text = "".join(chars)
    if lang_id in CHAR_LEVEL_LANGS:
        return list(text.replace(" ", ""))
    return text.split()


def run_ffmpeg_sync(bin_path, input_path, output_path, pause_ms=0):
    cmd = [bin_path, "-y", "-i", input_path, "-ar", "16000", "-ac", "1"]
    if pause_ms > 0:
        cmd += ["-af", "apad=pad_dur={:.3f}".format(pause_ms / 1000)]
    cmd += ["-f", "wav", output_path]

    # Use creationflags to hide the console window on Windows
    kwargs = {}
    if sys.platform == "win32":
        # 0x08000000 is CREATE_NO_WINDOW
        kwargs["creationflags"] = 0x08000000

    result = subprocess.run(
        cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True, encoding="utf-8", errors="ignore", **kwargs
    )
    if result.returncode != 0:
        detail = result.stderr.strip() or "Silent crash"
        raise RuntimeError(f"Exit {result.returncode} | {detail}")
    return True


async def process_line(sem, index, text, pause_ms, out_dir, voice, lang_id,
                       is_varied, min_pct, max_pct):
    base     = f"{index + 1:05d}"
    wav_path = os.path.join(out_dir, f"{base}.wav")
    tmp_mp3  = os.path.join(out_dir, f"{base}.tmp.mp3")

    tokens      = clean_text_for_mlf(text, lang_id)
    mlf_content = [f'"*/{base}.lab"\n'] + [f"{t}\n" for t in tokens if t] + [".\n"]
    mapping     = f"{base}.wav\t{text}\n"

    if os.path.exists(wav_path) and os.path.getsize(wav_path) > 1024:
        print(f"[CACHE] {base}.wav")
        return mapping, mlf_content

    rate    = pick_rate(min_pct, max_pct) if is_varied else "+0%"
    success = False

    async with sem:
        for attempt in range(3):
            try:
                comm = edge_tts.Communicate(text, voice, rate=rate)
                await asyncio.wait_for(comm.save(tmp_mp3), timeout=30)

                if not os.path.exists(tmp_mp3) or os.path.getsize(tmp_mp3) == 0:
                    raise FileNotFoundError("Downloaded MP3 is 0 bytes.")

                ffmpeg_bin = get_ffmpeg_path()
                if attempt == 0 and index == 0:
                    print(f"[DEBUG] FFmpeg: {ffmpeg_bin}")

                await asyncio.to_thread(
                    run_ffmpeg_sync, ffmpeg_bin, tmp_mp3, wav_path, pause_ms
                )

                if os.path.exists(wav_path) and os.path.getsize(wav_path) > 1024:
                    success = True
                    if os.path.exists(tmp_mp3):
                        os.remove(tmp_mp3)
                    tag = f"rate={rate}" if is_varied else "rate=normal"
                    print(f"[OK] {base}.wav | {tag} | pause={pause_ms}ms | {text[:35]}")
                    break
                else:
                    raise FileNotFoundError("WAV missing or empty after ffmpeg.")

            except Exception as e:
                print(f"[ERR {base} attempt {attempt+1}/3] {type(e).__name__}: {e}")
                await asyncio.sleep(2)

    if success:
        return mapping, mlf_content
    print(f"[FATAL] {base}.wav failed — dropped.")
    return None, None


async def run_engine(in_file, out_dir, lang_id, voice,
                     ratio_pct, min_pct, max_pct, concurrency, progress_cb=None):
    print(f"\n--- Pipeline | lang={lang_id} | voice={voice} "
          f"| varied={ratio_pct}% speed=[{min_pct}%~{max_pct}%] | conc={concurrency} ---")
    os.makedirs(out_dir, exist_ok=True)

    with open(in_file, "r", encoding="utf-8") as f:
        raw_lines = [l.rstrip("\n") for l in f if l.strip()]

    parsed = [parse_line(l) for l in raw_lines]
    total  = len(parsed)
    varied = assign_varied_indices(total, ratio_pct)
    print(f"[INFO] {total} utterances, {len(varied)} will be speed-varied")

    sem = asyncio.Semaphore(concurrency)

    async def _task(i, text, pause_ms):
        result = await process_line(
            sem, i, text, pause_ms, out_dir, voice, lang_id,
            i in varied, min_pct, max_pct
        )
        if progress_cb:
            progress_cb()
        return result

    tasks   = [_task(i, text, pause_ms) for i, (text, pause_ms) in enumerate(parsed)]
    results = await asyncio.gather(*tasks)

    valid = [r for r in results if r[0] is not None]
    if not valid:
        raise RuntimeError("0 audio files generated. Check ffmpeg or network.")

    with open(os.path.join(out_dir, "mapping.txt"), "w", encoding="utf-8") as fm, \
         open(os.path.join(out_dir, "labels.mlf"),  "w", encoding="utf-8") as fl:
        fl.write("#!MLF!#\n")
        for m, l in valid:
            fm.write(m)
            fl.writelines(l)

    print(f"--- Done. Yield: {len(valid)}/{len(parsed)} ---")
    return len(parsed)


# ------------------------------------------------------------------------
# 3. GUI
# ------------------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("ASR Acoustic Data Generator v2026")
        self.root.geometry("700x460")
        self.root.resizable(False, False)

        self.in_p        = tk.StringVar()
        self.out_p       = tk.StringVar()
        self.lang_name   = tk.StringVar(value="English (0)")
        self.voice_var   = tk.StringVar()
        self.spd_ratio_var = tk.StringVar(value="30")
        self.spd_min_var   = tk.StringVar(value="-20")
        self.spd_max_var   = tk.StringVar(value="20")
        self.concurr_var = tk.StringVar(value="10")

        self._total     = 0
        self._completed = 0

        self._build_ui()
        self._refresh_voices()

    def _build_ui(self):
        pad = {"padx": 8, "pady": 5}
        f = ttk.Frame(self.root, padding=20)
        f.pack(fill="both", expand=True)

        # Input / Output
        ttk.Label(f, text="Input Corpus (.txt):").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(f, textvariable=self.in_p, width=48).grid(row=0, column=1, columnspan=2, **pad)
        ttk.Button(f, text="Browse…", command=self._browse_in).grid(row=0, column=3, **pad)

        ttk.Label(f, text="Output Folder:").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(f, textvariable=self.out_p, width=48).grid(row=1, column=1, columnspan=2, **pad)
        ttk.Button(f, text="Browse…", command=self._browse_out).grid(row=1, column=3, **pad)

        ttk.Separator(f, orient="horizontal").grid(
            row=2, column=0, columnspan=4, sticky="ew", pady=8)

        # Language
        ttk.Label(f, text="Language:").grid(row=3, column=0, sticky="w", **pad)
        self.lang_cb = ttk.Combobox(
            f, textvariable=self.lang_name,
            values=list(LANG_CONFIG.keys()), width=28, state="readonly"
        )
        self.lang_cb.grid(row=3, column=1, sticky="w", **pad)
        self.lang_cb.bind("<<ComboboxSelected>>", lambda e: self._refresh_voices())

        # Voice
        ttk.Label(f, text="Voice:").grid(row=4, column=0, sticky="w", **pad)
        self.voice_cb = ttk.Combobox(
            f, textvariable=self.voice_var, width=44, state="readonly")
        self.voice_cb.grid(row=4, column=1, columnspan=2, sticky="w", **pad)

        ttk.Separator(f, orient="horizontal").grid(
            row=5, column=0, columnspan=4, sticky="ew", pady=8)

        # Speed variation
        ttk.Label(f, text="变速比例 (%):").grid(row=6, column=0, sticky="w", **pad)
        ttk.Entry(f, textvariable=self.spd_ratio_var, width=6).grid(
            row=6, column=1, sticky="w", **pad)

        ttk.Label(f, text="变速区间  Min:").grid(row=6, column=1, sticky="e", **pad)
        ttk.Entry(f, textvariable=self.spd_min_var, width=6).grid(
            row=6, column=2, sticky="w", **pad)

        ttk.Label(f, text="Max:").grid(row=6, column=2, sticky="e", **pad)
        ttk.Entry(f, textvariable=self.spd_max_var, width=6).grid(
            row=6, column=3, sticky="w", **pad)

        # Concurrency
        ttk.Label(f, text="Concurrency:").grid(row=7, column=0, sticky="w", **pad)
        ttk.Entry(f, textvariable=self.concurr_var, width=6).grid(
            row=7, column=1, sticky="w", **pad)

        ttk.Label(
            f,
            text='Tip: append "||500" to a line for 500 ms tail silence, e.g. "Hello world||500"',
            foreground="gray"
        ).grid(row=8, column=0, columnspan=4, sticky="w", padx=8, pady=(2, 6))

        ttk.Separator(f, orient="horizontal").grid(
            row=9, column=0, columnspan=4, sticky="ew", pady=4)

        # Start button
        self.btn = ttk.Button(
            self.root, text="START BATCH SYNTHESIS",
            command=self._start, width=38)
        self.btn.pack(pady=(8, 4))

        # Determinate progress bar
        self.pb_var = tk.DoubleVar(value=0.0)
        self.pb = ttk.Progressbar(
            self.root, variable=self.pb_var,
            mode="determinate", length=580, maximum=100)
        self.pb.pack(pady=(0, 2))

        # Progress label  "12 / 50"
        self.prog_lbl = ttk.Label(self.root, text="", foreground="#555")
        self.prog_lbl.pack()

        # Status label
        self.status_lbl = ttk.Label(self.root, text="Ready.", foreground="gray")
        self.status_lbl.pack(pady=(2, 0))

    def _browse_in(self):
        p = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if p:
            self.in_p.set(p)

    def _browse_out(self):
        p = filedialog.askdirectory()
        if p:
            self.out_p.set(p)

    def _refresh_voices(self):
        lang_config_entry = LANG_CONFIG.get(self.lang_name.get())
        if not lang_config_entry:
            return
            
        _, voice_pairs = lang_config_entry
        labels = [
            f"{'♀' if g == 'Female' else '♂'}  {n}"
            for n, g in voice_pairs
        ]
        
        if not labels:
            labels = ["No voices found for this locale"]
            
        self.voice_cb["values"] = labels
        self.voice_cb.current(0)
        self.voice_var.set(labels[0])
        self._voice_map = {
            f"{'♀' if g == 'Female' else '♂'}  {n}": n
            for n, g in voice_pairs
        }

    def _start(self):
        if not self.in_p.get() or not self.out_p.get():
            messagebox.showwarning("Missing Config",
                                   "Set both Input Corpus and Output Folder.")
            return
        # Count lines to initialise progress bar
        try:
            with open(self.in_p.get(), encoding="utf-8") as f:
                self._total = sum(1 for l in f if l.strip())
        except Exception:
            self._total = 0

        self._completed = 0
        self.pb_var.set(0)
        self.prog_lbl.config(text=f"0 / {self._total}")
        self.btn.config(state="disabled")
        self.status_lbl.config(text="Running…", foreground="blue")
        threading.Thread(target=self._worker, daemon=True).start()

    def _on_progress(self):
        """Called from the engine thread after each file finishes."""
        self._completed += 1
        n, t = self._completed, self._total
        pct = (n / t * 100) if t > 0 else 0
        self.root.after(0, lambda: self._update_progress(n, t, pct))

    def _update_progress(self, n, t, pct):
        self.pb_var.set(pct)
        self.prog_lbl.config(text=f"{n} / {t}")

    def _worker(self):
        try:
            lang_id   = LANG_CONFIG[self.lang_name.get()][0]
            voice     = self._voice_map.get(self.voice_var.get(), "")
            ratio_pct = max(0, min(100, int(self.spd_ratio_var.get() or 0)))
            min_pct   = int(self.spd_min_var.get() or 0)
            max_pct   = int(self.spd_max_var.get() or 0)
            concurr   = max(1, int(self.concurr_var.get() or 30))

            asyncio.run(run_engine(
                self.in_p.get(), self.out_p.get(),
                lang_id, voice, ratio_pct, min_pct, max_pct, concurr,
                progress_cb=self._on_progress
            ))
            self.root.after(0, lambda: messagebox.showinfo(
                "Complete", "All assets generated successfully."))
            self.root.after(0, lambda: self.status_lbl.config(
                text="Done.", foreground="green"))
        except Exception as e:
            err = str(e)
            self.root.after(0, lambda: messagebox.showerror("Error", err))
            self.root.after(0, lambda: self.status_lbl.config(
                text=f"Error: {err[:80]}", foreground="red"))
        finally:
            self.root.after(0, self._reset_ui)

    def _reset_ui(self):
        self.btn.config(state="normal")


# ------------------------------------------------------------------------
# 4. Entry Point (CLI / GUI router)
# ------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="ASR Data Generator — CLI/GUI")
    parser.add_argument("-i",  "--input",       help="Input corpus .txt")
    parser.add_argument("-d",  "--outdir",       help="Output directory")
    parser.add_argument("-l",  "--lang",         help="Language ID (0-27)")
    parser.add_argument("-v",  "--voice",        help="Voice short name")
    parser.add_argument("--spd-ratio", type=int, default=0,  help="Percentage of utterances to speed-vary (0-100)")
    parser.add_argument("--spd-min",   type=int, default=-20, help="Min speed %% for varied utterances")
    parser.add_argument("--spd-max",   type=int, default=20,  help="Max speed %% for varied utterances")
    parser.add_argument("-c",  "--concurrency",  type=int, default=30)
    parser.add_argument("--cli", action="store_true")
    args = parser.parse_args()

    if args.cli or (args.input and args.outdir and args.lang):
        if not (args.input and args.outdir and args.lang):
            print("[FATAL] CLI mode requires -i, -d, and -l.")
            sys.exit(1)
        entry = next((v for v in LANG_CONFIG.values() if v[0] == str(args.lang)), None)
        if not entry:
            print(f"[FATAL] Unknown language ID '{args.lang}'.")
            sys.exit(1)
        lang_id, voice_pairs = entry
        voice = args.voice or voice_pairs[0][0]
        try:
            asyncio.run(run_engine(
                args.input, args.outdir,
                lang_id, voice, args.spd_ratio, args.spd_min, args.spd_max, args.concurrency
            ))
        except KeyboardInterrupt:
            print("\n[WARN] Aborted by user.")
    else:
        if tk is None:      # 打包成无 tkinter 的 exe 时，无参启动没有图形界面
            print("[ERROR] 本环境未包含 tkinter 图形库，无法打开界面。"
                  "请用命令行 --cli 模式，或通过主程序调用。")
            sys.exit(1)
        root = tk.Tk()
        App(root)
        root.mainloop()


if __name__ == "__main__":
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    main()
