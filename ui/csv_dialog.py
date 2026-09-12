"""CSVのインポート・エクスポート画面。"""
import csv
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from ui.base_dialog import BaseDialog


CSV_COLUMNS = ["年", "月", "日", "列番号", "項目", "支払先", "金額", "メモ"]


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
        self.start_year = self._spin(period, now.year, 2000, 2100)
        tk.Label(period, text="年", bg="#f0f0f0").pack(side=tk.LEFT)
        self.start_month = self._spin(period, 1, 1, 12, width=4)
        tk.Label(period, text="月  ～  ", bg="#f0f0f0").pack(side=tk.LEFT)
        self.end_year = self._spin(period, now.year, 2000, 2100)
        tk.Label(period, text="年", bg="#f0f0f0").pack(side=tk.LEFT)
        self.end_month = self._spin(period, now.month, 1, 12, width=4)
        tk.Label(period, text="月", bg="#f0f0f0").pack(side=tk.LEFT)

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

    @staticmethod
    def _spin(parent, value, minimum, maximum, width=6):
        widget = ttk.Spinbox(parent, from_=minimum, to=maximum, width=width)
        widget.set(value)
        widget.pack(side=tk.LEFT, padx=(0, 3))
        return widget

    def _period(self):
        try:
            start = (int(self.start_year.get()), int(self.start_month.get()))
            end = (int(self.end_year.get()), int(self.end_month.get()))
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
            count = self.parent_app.data_manager.import_csv(
                path, start, end, mode, len(self.parent_app.get_all_columns()))
            self.parent_app._show_month(self.parent_app.current_month)
            messagebox.showinfo("完了", f"{count}件をインポートしました。", parent=self)
        except (OSError, UnicodeError, csv.Error, ValueError) as error:
            messagebox.showerror("インポートエラー", str(error), parent=self)
