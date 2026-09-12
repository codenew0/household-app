# ui/chart_dialog.py
"""
項目別の月間推移グラフを表示するダイアログ
"""
import tkinter as tk
from models.transactions import cash_amount, describe, normalize, is_pending, validate
from tkinter import ttk
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import date
from ui.base_dialog import BaseDialog
from config import DialogConfig, parse_amount
from utils.font_utils import setup_japanese_font


class ChartDialog(BaseDialog):
    """
    項目別の月間推移グラフを表示するダイアログ。
    """
    
    def __init__(self, parent, parent_app):
        """
        グラフダイアログを初期化する。
        """
        self.parent_app = parent_app
        self.current_column_index = -1  # -1: 総支出、-2: 総収入、1以上: 各項目
        
        # 表示対象の年（デフォルトは現在の年）
        self.target_year = parent_app.current_year
        
        self.figure = None
        self.canvas = None
        
        super().__init__(parent, "項目別月間推移グラフ",
                         DialogConfig.CHART_WIDTH,
                         DialogConfig.CHART_HEIGHT)
        
        self._create_widgets()
        self._update_chart()
        self.show_ready()
    
    def _get_available_years(self):
        """
        データが存在する年と現在の年を含むリストを取得する。
        降順（新しい年が上）で返す。
        """
        years = {self.parent_app.current_year}
        for dict_key in self.parent_app.data_manager.data.keys():
            try:
                # key形式: "YYYY-MM-DD-COL"
                y = int(dict_key.split("-")[0])
                years.add(y)
            except (ValueError, IndexError):
                pass
        return sorted(list(years), reverse=True)

    def _create_widgets(self):
        """グラフダイアログのUI要素を作成する"""
        # グリッドレイアウトの設定
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # タイトルセクション
        title_frame = tk.Frame(self, bg='#f0f0f0')
        title_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        
        # タイトル
        title_label = tk.Label(title_frame, text="項目別月間推移グラフ",
                               font=('Arial', 16, 'bold'), bg='#f0f0f0')
        title_label.pack(side=tk.LEFT)

        # 年選択エリア（タイトルの右側）
        year_frame = tk.Frame(title_frame, bg='#f0f0f0')
        year_frame.pack(side=tk.RIGHT)
        
        tk.Label(year_frame, text="対象年:", font=('Arial', 11), bg='#f0f0f0').pack(side=tk.LEFT, padx=5)
        
        self.year_var = tk.StringVar(value=str(self.target_year))
        self.year_combo = ttk.Combobox(year_frame, textvariable=self.year_var, 
                                       values=self._get_available_years(), 
                                       width=6, state="readonly", font=('Arial', 10))
        self.year_combo.pack(side=tk.LEFT)
        self.year_combo.bind("<<ComboboxSelected>>", self._on_year_change)
        
        # タブセクション
        tab_frame = tk.Frame(self, bg='#f0f0f0')
        tab_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 5))
        
        # 総支出・総収入ボタンコンテナ
        summary_button_frame = tk.Frame(tab_frame, bg='#f0f0f0')
        summary_button_frame.pack(side=tk.LEFT, padx=(0, 15))
        
        # 総支出ボタン
        self.total_expense_btn = tk.Button(summary_button_frame, text="総支出",
                                           font=('Arial', 11, 'bold'),
                                           bg='#f44336', fg='white',
                                           relief='raised', bd=2,
                                           activebackground='#d32f2f',
                                           command=lambda: self._select_tab(-1))
        self.total_expense_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 総収入ボタン
        self.total_income_btn = tk.Button(summary_button_frame, text="総収入",
                                          font=('Arial', 11, 'bold'),
                                          bg='#4caf50', fg='white',
                                          relief='raised', bd=2,
                                          activebackground='#45a049',
                                          command=lambda: self._select_tab(-2))
        self.total_income_btn.pack(side=tk.LEFT)
        
        # 区切り線
        separator = tk.Frame(tab_frame, bg='#cccccc', width=2)
        separator.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        
        # 項目別タブボタン(スクロール可能)
        all_columns = self.parent_app.get_all_columns()
        self.tab_buttons = []
        
        tab_canvas = tk.Canvas(tab_frame, height=35, bg='#f0f0f0', highlightthickness=0)
        tab_canvas.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        tab_inner_frame = tk.Frame(tab_canvas, bg='#f0f0f0')
        tab_canvas.create_window((0, 0), window=tab_inner_frame, anchor="nw")
        
        # 各項目のタブボタンを作成(日付列を除く)
        for i, col_name in enumerate(all_columns[1:], start=1):
            btn = tk.Button(tab_inner_frame, text=col_name,
                            font=('Arial', 10),
                            bg='#e0e0e0', fg='black',
                            relief='raised', bd=2,
                            command=lambda idx=i: self._select_tab(idx))
            btn.pack(side=tk.LEFT, padx=2, pady=2, fill=tk.Y)
            self.tab_buttons.append(btn)
        
        # 初期選択状態を反映
        self._update_button_colors()
        
        # タブフレームのスクロール領域を設定
        tab_inner_frame.update_idletasks()
        tab_canvas.configure(scrollregion=tab_canvas.bbox("all"))
        
        # グラフ表示セクション
        chart_frame = tk.Frame(self, bg='#f0f0f0')
        chart_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        chart_frame.grid_rowconfigure(0, weight=1)
        chart_frame.grid_columnconfigure(0, weight=1)
        
        # matplotlib設定
        plt.style.use('default')
        setup_japanese_font()
        
        # グラフの作成
        self.figure = plt.Figure(figsize=(10, 5), dpi=80)
        self.canvas = FigureCanvasTkAgg(self.figure, chart_frame)
        self.canvas.get_tk_widget().grid(row=0, column=0, sticky="nsew")
        
        # 閉じるボタン
        close_btn = tk.Button(self, text="閉じる", font=('Arial', 12),
                              bg='#f44336', fg='white', relief='raised', bd=2,
                              activebackground='#d32f2f', command=self.destroy)
        close_btn.grid(row=3, column=0, pady=10, ipady=5)
        
        # キーボードショートカット
        self.bind('<Escape>', lambda e: self.destroy())
    
    def _on_year_change(self, event):
        """年が変更された時の処理"""
        try:
            self.target_year = int(self.year_var.get())
            self._update_chart()
        except ValueError:
            pass

    def _select_tab(self, index):
        """タブを選択してグラフを更新する"""
        self.current_column_index = index
        self._update_button_colors()
        self._update_chart()
    
    def _update_button_colors(self):
        """選択されているタブのボタンの色を更新する"""
        self.total_expense_btn.config(bg='#f44336', fg='white')
        self.total_income_btn.config(bg='#4caf50', fg='white')
        
        for btn in self.tab_buttons:
            btn.config(bg='#e0e0e0', fg='black')
        
        if self.current_column_index == -1:  # 総支出
            self.total_expense_btn.config(bg='#d32f2f', fg='white')
        elif self.current_column_index == -2:  # 総収入
            self.total_income_btn.config(bg='#45a049', fg='white')
        elif 0 < self.current_column_index <= len(self.tab_buttons):
            self.tab_buttons[self.current_column_index - 1].config(bg='#2196f3', fg='white')
    
    def _update_chart(self):
        """現在選択されているタブと年に応じてグラフを更新する"""
        if not self.figure:
            return
        
        # 既存のグラフをクリア
        self.figure.clear()
        
        # 選択されたタブに応じてデータを取得
        if self.current_column_index == -1:  # 総支出
            monthly_data = self._collect_total_expense_data()
            title = f"{self.target_year}年 月間総支出の推移"
            ylabel = "支出額 (円)"
            color = '#f44336'
        elif self.current_column_index == -2:  # 総収入
            monthly_data = self._collect_total_income_data()
            title = f"{self.target_year}年 月間総収入の推移"
            ylabel = "収入額 (円)"
            color = '#4caf50'
        else:  # 各項目
            monthly_data = self._collect_category_data()
            all_columns = self.parent_app.get_all_columns()
            column_name = all_columns[self.current_column_index] if self.current_column_index < len(
                all_columns) else "不明"
            title = f'{self.target_year}年 {column_name} の月間推移'
            ylabel = "金額 (円)"
            color = '#2196f3'
        
        # 対象年の1月〜12月を0埋め（データがない月も0円として表示するため）
        for m in range(1, 13):
            d = date(self.target_year, m, 1)
            if d not in monthly_data:
                monthly_data[d] = 0

        # 全てのデータが0の場合は「データがありません」を表示
        if not monthly_data or all(v == 0 for v in monthly_data.values()):
            ax = self.figure.add_subplot(111)
            ax.text(0.5, 0.5, 'データがありません',
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax.transAxes, fontsize=16)
            ax.set_title(title)
            self.canvas.draw()
            return
        
        # グラフの描画
        ax = self.figure.add_subplot(111)
        
        # データを日付順にソート
        sorted_data = sorted(monthly_data.items())
        dates = [item[0] for item in sorted_data]
        amounts = [item[1] for item in sorted_data]
        
        # 折れ線グラフを描画
        ax.plot(dates, amounts, marker='o', linewidth=2, markersize=8, color=color)
        
        # 各データポイントに金額ラベルを追加
        for date_val, amount in zip(dates, amounts):
            # 金額が0の場合はラベルを表示しない等の調整も可能だが、現状は全て表示
            ax.annotate(f'¥{amount:,}',
                        xy=(date_val, amount),
                        xytext=(0, 10),
                        textcoords='offset points',
                        ha='center',
                        va='bottom',
                        fontsize=9,
                        bbox=dict(boxstyle='round,pad=0.3',
                                  facecolor='white',
                                  edgecolor=color,
                                  alpha=0.8))
        
        # グラフの装飾
        ax.set_title(title, fontsize=14, fontweight='bold')
        ax.set_xlabel('月', fontsize=12)
        ax.set_ylabel(ylabel, fontsize=12)
        
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{int(x):,}'))
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%m月')) # 年はタイトルにあるので月のみ
        ax.xaxis.set_major_locator(mdates.MonthLocator())
        
        ax.grid(True, alpha=0.3)
        
        # Y軸の範囲を調整
        y_min, y_max = ax.get_ylim()
        if y_min == y_max:
             y_range = 1000
             ax.set_ylim(0, y_range)
        else:
             y_range = y_max - y_min
             ax.set_ylim(y_min - y_range * 0.05, y_max + y_range * 0.15)
        
        self.figure.tight_layout()
        self.canvas.draw()
    
    def _collect_total_expense_data(self):
        """対象年の月間総支出データを収集する"""
        monthly_totals = {}
        
        for dict_key, data_list in self.parent_app.data_manager.data.items():
            try:
                parts = dict_key.split("-")
                if len(parts) == 4:
                    year = int(parts[0])
                    month = int(parts[1])
                    day = int(parts[2])
                    
                    # 年フィルタリング
                    if year != self.target_year:
                        continue
                    
                    if day == 0:
                        continue
                    
                    month_key = date(year, month, 1)
                    if month_key not in monthly_totals:
                        monthly_totals[month_key] = 0
                    
                    for row in data_list:
                        if len(row) > 1:
                            amount = cash_amount(row)
                            monthly_totals[month_key] += amount
            except (ValueError, IndexError):
                continue
        
        return monthly_totals
    
    def _collect_total_income_data(self):
        """対象年の月間総収入データを収集する"""
        monthly_totals = {}
        
        for dict_key, data_list in self.parent_app.data_manager.data.items():
            try:
                parts = dict_key.split("-")
                if len(parts) == 4:
                    year = int(parts[0])
                    month = int(parts[1])
                    day = int(parts[2])
                    col_index = int(parts[3])
                    
                    # 年フィルタリング
                    if year != self.target_year:
                        continue
                    
                    # まとめ行の収入列のみ対象
                    if day == 0 and col_index == 3:
                        month_key = date(year, month, 1)
                        total_income = sum(cash_amount(row) for row in data_list if len(row) > 1)
                        if total_income > 0:
                            monthly_totals[month_key] = total_income
            except (ValueError, IndexError):
                continue
        
        return monthly_totals
    
    def _collect_category_data(self):
        """対象年の特定項目の月間データを収集する"""
        monthly_totals = {}
        
        for dict_key, data_list in self.parent_app.data_manager.data.items():
            try:
                parts = dict_key.split("-")
                if len(parts) == 4:
                    year = int(parts[0])
                    month = int(parts[1])
                    day = int(parts[2])
                    col_index = int(parts[3])
                    
                    # 年フィルタリング
                    if year != self.target_year:
                        continue
                    
                    if day == 0:
                        continue
                    
                    if col_index == self.current_column_index:
                        month_key = date(year, month, 1)
                        if month_key not in monthly_totals:
                            monthly_totals[month_key] = 0
                        
                        for row in data_list:
                            if len(row) > 1:
                                amount = cash_amount(row)
                                monthly_totals[month_key] += amount
            except (ValueError, IndexError):
                continue
        
        return monthly_totals
