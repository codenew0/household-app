"""CSVのインポート・エクスポート画面。"""
import csv
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from ui.base_dialog import BaseDialog
from ui.month_picker import MonthPicker

class CsvDialog(BaseDialog):
    """指定期間のCSV入出力を行う。"""

    def __init__(self, parent, parent_app):
        self.parent_app = parent_app
        super().__init__(parent, "CSV インポート・エクスポート", 540, 330)
        self._create_widgets()
        self.show_ready()

    def _create_widgets(self):
        now = datetime.date.today()
        frame = tk.Frame(self, bg="#f0f0f0")
        frame.pack(fill=tk.BOTH, expand=True, padx=24, pady=20)

        tk.Label(frame, text="対象期間", font=("Arial", 14, "bold"),
                 bg="#f0f0f0").pack(anchor="w", pady=(0, 12))

        period = tk.Frame(frame, bg="#f0f0f0")
        period.pack(fill=tk.X)
        tk.Label(period, text="開始年月", bg="#f0f0f0").pack(side=tk.LEFT)
        self.start_period = ttk.Entry(period, width=11, state='normal', cursor='hand2')
        self.start_period.insert(0, f'{now.year}-01')
        self.start_period.configure(state='readonly')
        self.start_period.pack(side=tk.LEFT, padx=(6, 14))
        tk.Label(period, text="終了年月", bg="#f0f0f0").pack(side=tk.LEFT)
        self.end_period = ttk.Entry(period, width=11, state='normal', cursor='hand2')
        self.end_period.insert(0, f'{now.year}-{now.month:02d}')
        self.end_period.configure(state='readonly')
        self.end_period.pack(side=tk.LEFT, padx=6)
        self.start_picker = MonthPicker(period, now.year, 1,
                                        lambda year, month: self._pick_period(self.start_period, year, month),
                                        anchor=self.start_period, theme='light')
        self.end_picker = MonthPicker(period, now.year, now.month,
                                      lambda year, month: self._pick_period(self.end_period, year, month),
                                      anchor=self.end_period, theme='light')
        self.start_period.bind('<Button-1>', lambda event: self._open_picker(self.start_period, self.start_picker))
        self.end_period.bind('<Button-1>', lambda event: self._open_picker(self.end_period, self.end_picker))

        mode_frame = tk.LabelFrame(frame, text="インポート方式", bg="#f0f0f0",
                                   font=("Arial", 11, "bold"), padx=10, pady=8)
        mode_frame.pack(fill=tk.X, pady=18)
        self.import_mode = tk.StringVar(value="append")
        tk.Radiobutton(mode_frame, text="追加（既存データを残す）", variable=self.import_mode,
                       value="append", bg="#f0f0f0").pack(anchor="w")
        tk.Radiobutton(mode_frame, text="上書き（指定期間をCSV内容で置き換え）",
                       variable=self.import_mode, value="replace", bg="#f0f0f0").pack(anchor="w")

        buttons = tk.Frame(frame, bg="#f0f0f0")
        buttons.pack(fill=tk.X, pady=(4, 0))
        tk.Button(buttons, text="CSVをエクスポート", command=self._export,
                  bg="#2196f3", fg="white", font=("Arial", 11, "bold"), padx=12,
                  pady=7).pack(side=tk.LEFT)
        tk.Button(buttons, text="CSVをインポート", command=self._import,
                  bg="#4caf50", fg="white", font=("Arial", 11, "bold"), padx=12,
                  pady=7).pack(side=tk.LEFT, padx=10)
        tk.Button(buttons, text="閉じる", command=self.destroy, bg="#777", fg="white",
                  font=("Arial", 11), padx=12, pady=7).pack(side=tk.RIGHT)

    def _open_picker(self, field, picker):
        year, month = map(int, field.get().split('-'))
        picker.set(year, month)
        picker.open()
        return 'break'

    def _pick_period(self, field, year, month):
        field.configure(state='normal')
        field.delete(0, 'end')
        field.insert(0, f'{year}-{month:02d}')
        field.configure(state='readonly')

    def _period(self):
        try:
            start = tuple(map(int, self.start_period.get().split('-')))
            end = tuple(map(int, self.end_period.get().split('-')))
        except ValueError:
            raise ValueError("年月は数値で指定してください。")
        if not 1 <= start[1] <= 12 or not 1 <= end[1] <= 12:
            raise ValueError("月は1～12で指定してください。")
        if start > end:
            raise ValueError("開始年月は終了年月以前にしてください。")
        return start, end

    def _export(self):
        try:
            start, end = self._period()
        except ValueError as error:
            messagebox.showwarning("入力エラー", str(error), parent=self)
            return
        path = filedialog.asksaveasfilename(
            parent=self, title="CSVの保存先", defaultextension=".csv",
            filetypes=[("CSVファイル", "*.csv")],
            initialfile=f"household_{start[0]}{start[1]:02d}_{end[0]}{end[1]:02d}.csv")
        if not path:
            return
        try:
            count = self.parent_app.data_manager.export_csv(
                path, start, end, self.parent_app.get_all_columns())
            messagebox.showinfo("完了", f"{count}件をエクスポートしました。", parent=self)
        except (OSError, UnicodeError, csv.Error) as error:
            messagebox.showerror("エクスポートエラー", str(error), parent=self)

    def _import(self):
        try:
            start, end = self._period()
        except ValueError as error:
            messagebox.showwarning("入力エラー", str(error), parent=self)
            return
        path = filedialog.askopenfilename(
            parent=self, title="CSVを選択", filetypes=[("CSVファイル", "*.csv")])
        if not path:
            return
        mode = self.import_mode.get()
        if mode == "replace" and not messagebox.askyesno(
                "上書き確認", "指定期間の既存データを削除し、CSV内容で置き換えますか？",
                parent=self):
            return
        try:
            manager = self.parent_app.data_manager
            imported = manager.import_csv(path, start, end, mode, len(self.parent_app.get_all_columns()), preview_only=True)
            duplicates = manager.duplicate_candidates(imported, include_existing=mode == 'append')
            if duplicates:
                review = DuplicateReviewDialog(self, duplicates)
                self.wait_window(review)
                self.grab_set()
                if not review.accepted:
                    return
            count = self.parent_app.data_manager.import_csv(
                path, start, end, mode, len(self.parent_app.get_all_columns()))
            # CSV変更はメイン画面の既存Undo履歴では安全に巻き戻せない。
            self.parent_app.undo_stack.clear()
            self.parent_app._show_month(self.parent_app.current_month)
            messagebox.showinfo("完了", f"{count}件をインポートしました。", parent=self)
        except (OSError, UnicodeError, csv.Error, ValueError) as error:
            messagebox.showerror("インポートエラー", str(error), parent=self)


class DuplicateReviewDialog(BaseDialog):
    def __init__(self, parent, duplicates):
        super().__init__(parent, 'CSVの重複候補を確認', 780, 500)
        self.accepted = False
        ttk.Label(self, text=f'{len(duplicates)}件の重複候補があります（日付・支払先・利用前金額が一致）。\n同額の別取引の場合もあるため、自動削除はしません。').pack(padx=12, pady=12)
        tree = ttk.Treeview(self, columns=('date', 'partner', 'amount', 'memo'), show='headings')
        for key, title in zip(tree['columns'], ('日付キー', '支払先', '金額', 'メモ')):
            tree.heading(key, text=title)
        tree.pack(fill='both', expand=True, padx=12)
        for key, row in duplicates:
            tree.insert('', 'end', values=(key, row[0], row[1], row[2]))
        bar = ttk.Frame(self); bar.pack(fill='x', padx=12, pady=12)
        ttk.Button(bar, text='キャンセルして見直す', command=self.destroy).pack(side='right')
        ttk.Button(bar, text='候補を含めて取り込む', command=self.accept).pack(side='right', padx=8)
        self.show_ready()

    def accept(self):
        self.accepted = True
        self.destroy()
