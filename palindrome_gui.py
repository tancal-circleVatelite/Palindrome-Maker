import queue
import threading
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import time

from palindrome_bunsetsu import (
    Bunsetsu,
    compute_naturality_score,
    generate_palindrome_candidates,
    load_bunsetsu_csv,
    find_seed,
    write_candidates_excel,
    safe_filename,
)


class PalindromeGeneratorApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("文節結合法による回文生成")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)

        self.csv_path = tk.StringVar(value="")
        self.output_dir = tk.StringVar(value=str(Path.cwd()))
        self.output_name = tk.StringVar(value="Result")
        self.target_count = tk.IntVar(value=4)
        self.status_text = tk.StringVar(value="CSV ファイルを選択してください。")
        self.latest_candidate = tk.StringVar(value="生成候補はここに表示されます。")

        self.phrase_entries: list[Bunsetsu] = []
        self.display_to_bunsetsu: dict[str, Bunsetsu] = {}
        self.generation_queue: queue.Queue = queue.Queue()
        self.is_generating = False
        self.target_seed_text = tk.StringVar(value="")
        self.target_seed_reading = tk.StringVar(value="")
        self.target_table = None
        self.generate_button = None
        self.cancel_button = None
        self.csv_select_button = None
        self.footer_font = self._resolve_footer_font()

        self._build_ui()

    @staticmethod
    def _resolve_footer_font():
        font_candidates = [
            Path(__file__).resolve().with_name("HelloHello-H.otf"),
            Path.cwd() / "HelloHello-H.otf",
        ]

        for font_path in font_candidates:
            if not font_path.exists():
                continue
            try:
                return tkfont.Font(file=str(font_path), size=12)
            except Exception:
                pass

        try:
            return tkfont.nametofont("TkDefaultFont").copy()
        except Exception:
            return tkfont.Font(size=12)

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=16)
        main.pack(fill="both", expand=True)

        # CSV selection
        csv_frame = ttk.LabelFrame(main, text="1. 文節集合 CSV", padding=12)
        csv_frame.pack(fill="x", pady=(0, 12))

        csv_row = ttk.Frame(csv_frame)
        csv_row.pack(fill="x")

        ttk.Entry(csv_row, textvariable=self.csv_path, state="readonly").pack(side="left", fill="x", expand=True)
        self.csv_select_button = ttk.Button(csv_row, text="参照", command=self._choose_csv)
        self.csv_select_button.pack(side="left", padx=(8, 0))

        target_frame = ttk.LabelFrame(main, text="2. 生成対象語句", padding=12)
        target_frame.pack(fill="x", pady=(0, 12))

        entry_row = ttk.Frame(target_frame)
        entry_row.pack(fill="x", pady=(0, 8))

        ttk.Label(entry_row, text="語句").pack(side="left")
        seed_text_entry = ttk.Entry(entry_row, textvariable=self.target_seed_text, width=30)
        seed_text_entry.pack(side="left", padx=(8, 12))
        seed_text_entry.bind("<Return>", lambda event: self._add_target_row())

        ttk.Label(entry_row, text="ひらがなの読み").pack(side="left")
        seed_reading_entry = ttk.Entry(entry_row, textvariable=self.target_seed_reading, width=30)
        seed_reading_entry.pack(side="left", padx=(8, 12))
        seed_reading_entry.bind("<Return>", lambda event: self._add_target_row())

        ttk.Button(entry_row, text="追加", command=self._add_target_row).pack(side="left")

        self.target_table = ttk.Treeview(target_frame, columns=("seed_text", "seed_reading"), show="headings", height=6)
        self.target_table.heading("seed_text", text="seed_text")
        self.target_table.heading("seed_reading", text="seed_reading")
        self.target_table.column("seed_text", width=220, anchor="center")
        self.target_table.column("seed_reading", width=220, anchor="center")
        self.target_table.pack(fill="x", pady=(0, 8))
        self.target_table.tag_configure("grid", background="#ffffff")
        self.target_table.configure(style="Treeview")

        style = ttk.Style()
        style.configure("Treeview", rowheight=26, borderwidth=1, relief="solid")
        style.map("Treeview", background=[("selected", "#cfe8ff")])

        self.target_table.bind("<Control-c>", self._copy_target_table)
        self.target_table.bind("<Control-v>", self._paste_target_table)
        self.target_table.bind("<Control-C>", self._copy_target_table)
        self.target_table.bind("<Control-V>", self._paste_target_table)

        table_button_row = ttk.Frame(target_frame)
        table_button_row.pack(fill="x")
        ttk.Button(table_button_row, text="選択行を削除", command=self._remove_selected_target_row).pack(side="left")

        # Output settings
        output_frame = ttk.LabelFrame(main, text="3. 出力設定", padding=12)
        output_frame.pack(fill="x", pady=(0, 12))

        output_dir_row = ttk.Frame(output_frame)
        output_dir_row.pack(fill="x", pady=(0, 8))

        ttk.Label(output_dir_row, text="出力先フォルダ:").pack(side="left")
        ttk.Entry(output_dir_row, textvariable=self.output_dir).pack(side="left", fill="x", expand=True, padx=(8, 8))
        ttk.Button(output_dir_row, text="参照", command=self._choose_output_dir).pack(side="left")

        target_count_row = ttk.Frame(output_frame)
        target_count_row.pack(fill="x", pady=(0, 8))

        ttk.Label(target_count_row, text="文節数:").pack(side="left")
        target_count_spinbox = ttk.Spinbox(
            target_count_row,
            from_=1,
            to=100,
            textvariable=self.target_count,
            width=8,
            justify="center",
        )
        target_count_spinbox.pack(side="left", padx=(8, 0))
        ttk.Label(target_count_row, text="文節").pack(side="left", padx=(8, 0))

        output_name_row = ttk.Frame(output_frame)
        output_name_row.pack(fill="x")

        ttk.Label(output_name_row, text="出力ファイル名:").pack(side="left")
        ttk.Entry(output_name_row, textvariable=self.output_name).pack(side="left", fill="x", expand=True, padx=(8, 0))
        ttk.Label(output_name_row, text=".xlsx").pack(side="left", padx=(8, 0))

        ttk.Label(output_frame, text="複数の語句を指定した場合は末尾に対象語句を追記したExcelファイルがそれぞれ出力されます", foreground="#555555").pack(anchor="w", pady=(8, 0))

        # Run button
        run_row = ttk.Frame(main)
        run_row.pack(fill="x", pady=(0, 8))
        self.generate_button = ttk.Button(run_row, text="回文を生成", command=self._generate)
        self.generate_button.pack(fill="x")

        cancel_row = ttk.Frame(main)
        cancel_row.pack(fill="x", pady=(0, 8))
        self.cancel_button = ttk.Button(cancel_row, text="生成中止", command=self._cancel_generation)
        self.cancel_button.pack(side="right")

        self.root.bind("<Return>", self._handle_return_key)

        status_row = ttk.Frame(main)
        status_row.pack(fill="x")
        ttk.Label(status_row, textvariable=self.status_text, foreground="#1f5f99").pack(anchor="w")

        candidate_frame = ttk.Frame(main)
        candidate_frame.pack(fill="x", pady=(10, 0))
        ttk.Label(candidate_frame, text="生成中候補:", font=("", 9, "bold")).pack(anchor="w")
        ttk.Label(candidate_frame, textvariable=self.latest_candidate, anchor="w", relief="sunken", padding=(6, 4), foreground="#1f3a5f").pack(fill="x")

        footer = tk.Label(
            self.root,
            text="サークルばてライト",
            font=self.footer_font,
            fg="#6a6a6a",
            bg=self.root.cget("bg"),
            anchor="e",
        )
        footer.pack(fill="x", side="bottom", padx=18, pady=(0, 10))

    def _choose_csv(self) -> None:
        path = filedialog.askopenfilename(
            title="CSVファイルを選択",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        self.csv_path.set(path)
        self._reload_bunsetsu_list()

    def _choose_output_dir(self) -> None:
        directory = filedialog.askdirectory(title="出力先フォルダを選択")
        if directory:
            self.output_dir.set(directory)

    def _reload_bunsetsu_list(self) -> None:
        csv_path = self.csv_path.get().strip()
        if not csv_path:
            self.phrase_entries = []
            self.display_to_bunsetsu = {}
            self.status_text.set("CSV ファイルを選択してください。")
            return

        try:
            entries = load_bunsetsu_csv(csv_path)
        except Exception as exc:  # pragma: no cover - GUI error path
            messagebox.showerror("CSV読み込みエラー", f"CSVの読み込みに失敗しました:\n{exc}")
            return

        self.phrase_entries = entries
        self.display_to_bunsetsu = {}

        unique_entries: list[Bunsetsu] = []
        seen_texts: set[str] = set()
        for item in entries:
            if item.text in seen_texts:
                continue
            seen_texts.add(item.text)
            unique_entries.append(item)

        for item in unique_entries:
            label = f"{item.text} ({item.reading})"
            self.display_to_bunsetsu[label] = item

        if unique_entries:
            self.status_text.set(f"{len(unique_entries)} 件の語句を読み込みました。CSV件数: {len(entries)}")
        else:
            self.status_text.set("CSV に語句が見つかりませんでした。")

    def _set_candidate_display(self, text: str) -> None:
        self.latest_candidate.set(text)

    def _process_queue(self) -> None:
        while True:
            try:
                item = self.generation_queue.get_nowait()
            except queue.Empty:
                break

            kind = item.get("kind")
            if kind == "candidate":
                st = item["state"]
                score = compute_naturality_score(st.parts)
                msg = f"最新候補: {st.text()} / {st.reading()} / 自然度={score}"
                self._set_candidate_display(msg)
            elif kind == "progress":
                done = item["done"]
                total = item["total"]
                if total:
                    percent = int(done * 100 / total)
                    self.status_text.set(f"回文生成を実行中です... {percent}% ({done}/{total})")

        if self.is_generating:
            self.root.after(100, self._process_queue)

    def _queue_candidate(self, st) -> None:
        self.generation_queue.put({"kind": "candidate", "state": st})

    def _queue_progress(self, done: int, total: int) -> None:
        self.generation_queue.put({"kind": "progress", "done": done, "total": total})

    def _add_target_row(self) -> None:
        text = self.target_seed_text.get().strip()
        reading = self.target_seed_reading.get().strip()
        if not text and not reading:
            messagebox.showwarning("入力エラー", "語句と読みを入力してください。")
            return

        if not text:
            text = reading
        if not reading:
            reading = text

        self.target_table.insert("", tk.END, values=(text, reading))
        self.target_seed_text.set("")
        self.target_seed_reading.set("")

    def _remove_selected_target_row(self) -> None:
        if self.target_table is None:
            return
        selected = self.target_table.selection()
        for item in selected:
            self.target_table.delete(item)

    def _copy_target_table(self, event=None) -> None:
        if self.target_table is None:
            return
        selected = self.target_table.selection()
        if not selected:
            return
        rows = []
        for item in selected:
            values = self.target_table.item(item, "values")
            rows.append("\t".join(str(v) for v in values))
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(rows))
        return "break"

    def _paste_target_table(self, event=None) -> None:
        if self.target_table is None:
            return
        try:
            text = self.root.clipboard_get()
        except tk.TclError:
            return
        if not text:
            return

        for line in text.splitlines():
            cells = [cell.strip() for cell in line.split("\t")]
            if len(cells) == 1:
                cells = [cells[0], ""]
            if len(cells) >= 2:
                self.target_table.insert("", tk.END, values=(cells[0], cells[1]))

        return "break"

    def _parse_target_phrases(self) -> list[tuple[str, str]]:
        if self.target_table is None:
            return []

        rows: list[tuple[str, str]] = []
        for item in self.target_table.get_children():
            values = self.target_table.item(item, "values")
            if not values:
                continue
            seed_text = str(values[0]).strip()
            seed_reading = str(values[1]).strip() if len(values) > 1 else ""
            if not seed_text and not seed_reading:
                continue
            if not seed_text:
                seed_text = seed_reading
            if not seed_reading:
                seed_reading = seed_text
            rows.append((seed_text, seed_reading))
        return rows

    def _normalize_phrase_entry(self, seed_text: str, seed_reading: str, db: list[Bunsetsu]) -> Bunsetsu:
        exact_match = next((item for item in db if item.text == seed_text and item.reading == seed_reading), None)
        if exact_match is not None:
            return exact_match

        if not seed_reading:
            seed_reading = seed_text
        appended = Bunsetsu(text=seed_text, reading=seed_reading, btype="custom")
        db.append(appended)
        return appended

    def _handle_return_key(self, event=None) -> str:
        if self.is_generating:
            return "break"
        if self.target_seed_text.get().strip() or self.target_seed_reading.get().strip():
            self._add_target_row()
            return "break"
        return "break"

    def _generation_complete(self, results: list[str]) -> None:
        self.is_generating = False
        self._set_generation_lock(False)
        self.status_text.set(f"{len(results)} 件の出力ファイルを作成しました。")
        self.root.title("文節結合法による回文生成")
        if results:
            summary = "\n".join(results)
            messagebox.showinfo("完了", f"回文候補を出力しました。\n\n{summary}")

    def _generation_failed(self, exc: Exception) -> None:
        self.is_generating = False
        self._set_generation_lock(False)
        self.status_text.set("回文生成に失敗しました。")
        self.root.title("文節結合法による回文生成")
        messagebox.showerror("生成エラー", f"回文生成中にエラーが発生しました:\n{exc}")

    def _cancel_generation(self) -> None:
        if not self.is_generating:
            return
        self.status_text.set("回文生成を中止しました。")
        self.latest_candidate.set("生成は中止されました。")
        self.root.title("文節結合法による回文生成")
        self.is_generating = False
        self._set_generation_lock(False)

    def _set_generation_lock(self, is_locked: bool) -> None:
        widgets = [
            self.csv_select_button,
            self.generate_button,
            self.cancel_button,
            self.target_table,
        ]
        for widget in widgets:
            if widget is None:
                continue
            if widget is self.cancel_button:
                widget.state(["!disabled"])
                continue
            if is_locked:
                widget.state(["disabled"])
            else:
                widget.state(["!disabled"])

    def _generate(self) -> None:
        if self.is_generating:
            return

        csv_path = self.csv_path.get().strip()
        if not csv_path:
            messagebox.showwarning("入力エラー", "文節集合CSVを選択してください")
            return

        target_phrases = self._parse_target_phrases()
        if not target_phrases:
            messagebox.showwarning("入力エラー", "対象語句を 1 件以上入力してください。")
            return

        try:
            target_count = int(self.target_count.get())
        except (TypeError, ValueError):
            messagebox.showwarning("入力エラー", "文節数は 1 以上の整数を入力してください。")
            return

        if target_count < 1:
            messagebox.showwarning("入力エラー", "文節数は 1 以上を指定してください。")
            return

        output_dir = Path(self.output_dir.get()).expanduser().resolve()
        output_dir.mkdir(parents=True, exist_ok=True)

        base_name = self.output_name.get().strip() or "palindrome_result"
        if not base_name:
            base_name = "palindrome_result"

        self.is_generating = True
        self._set_generation_lock(True)
        self.cancel_button.state(["!disabled"])
        self.root.title("文節結合法による回文生成 - 実行中")
        self.status_text.set("回文生成を実行中です... しばらくお待ちください。")
        self._set_candidate_display("生成中... 候補が見つかるたびに更新されます。")
        self.generation_queue = queue.Queue()

        def worker() -> None:
            try:
                db = load_bunsetsu_csv(csv_path)
                results: list[str] = []
                for seed_text, seed_reading in target_phrases:
                    if not self.is_generating:
                        break

                    seed = self._normalize_phrase_entry(seed_text, seed_reading, db)
                    start = time.perf_counter()
                    cands, explored_nodes = generate_palindrome_candidates(
                        seed=seed,
                        db=db,
                        target_bunsetsu=target_count,
                        allow_duplicate_parts=True,
                        progress_callback=self._queue_progress,
                        result_callback=self._queue_candidate,
                        cancel_callback=lambda: not self.is_generating,
                    )
                    elapsed_seconds = time.perf_counter() - start

                    if not self.is_generating:
                        break

                    file_name = f"{base_name}_{safe_filename(seed.text)}.xlsx" if len(target_phrases) > 1 else f"{base_name}.xlsx"
                    out_path = output_dir / file_name
                    write_candidates_excel(
                        out_path,
                        seed,
                        target_count,
                        cands,
                        elapsed_seconds,
                        explored_nodes,
                        {
                            "csv": csv_path,
                            "seed_text": seed.text,
                            "seed_reading": seed.reading,
                            "target_bunsetsu": target_count,
                            "output_dir": str(output_dir),
                            "output_file": str(out_path),
                            "selected_targets": ",".join(f"{t}/{r}" for t, r in target_phrases),
                        },
                    )
                    results.append(str(out_path))

                if self.is_generating:
                    self.root.after(0, self._generation_complete, results)
                else:
                    self.root.after(0, self._generation_cancelled)
            except Exception as exc:
                self.root.after(0, self._generation_failed, exc)

        threading.Thread(target=worker, daemon=True).start()
        self.root.after(100, self._process_queue)

    def _generation_cancelled(self) -> None:
        self.is_generating = False
        self._set_generation_lock(False)
        self.status_text.set("回文生成を中止しました。")
        self.latest_candidate.set("生成は中止されました。")
        self.root.title("文節結合法による回文生成")


def main() -> None:
    root = tk.Tk()
    app = PalindromeGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
