# ui/main_window.py
"""
メインウィンドウの実装
家計管理アプリケーションのメインUI
"""
import tkinter as tk
import json
from tkinter import ttk, messagebox
import tkinter.font as tkfont
from config import (
    WindowConfig, ColorTheme, TreeviewConfig, DefaultColumns,
    FontConfig, get_current_year, get_current_month, parse_amount
)
from models.data_manager import DataManager
from ui.tooltip import TreeviewTooltip
from ui.transaction_dialog import TransactionDialog
from ui.monthly_data_dialog import MonthlyDataDialog
from ui.search_dialog import SearchDialog
from ui.chart_dialog import ChartDialog
from utils.date_utils import get_days_in_month
import datetime
import re


class MainWindow:
    """
    家計管理アプリケーションのメインウィンドウクラス。
    
    年間の家計データを月別に管理し、項目ごとの支出・収入を
    記録・集計・分析する機能を提供する。
    """
    
    def __init__(self, root):
        """
        メインウィンドウを初期化する。
        
        Args:
            root: Tkinterのルートウィンドウ
        """
        self.root = root
        self.data_manager = DataManager()
        self.tree = None
        self.tooltip = None
        self.current_year = get_current_year()
        self.current_month = get_current_month()
        self.colors = self._get_color_theme()

        # コピペ用：選択された列のIDを保持
        self.selected_column_id = None
        
        # 範囲選択用
        self.selection_start_row = None  # 範囲選択の開始行
        self.selection_start_col = None  # 範囲選択の開始列
        
        # Ctrl選択用：個別に選択されたセル [(row_id, col_id), ...]
        self.ctrl_selected_cells = []
        
        # 元に戻す機能用
        self.undo_stack = []  # 操作履歴 [(action_type, data), ...]
        self.max_undo_count = 50  # 最大保持数
        
        # 月選択ボタンのリスト
        self.month_buttons = []
        self.current_month_button = None
        self.year_label = None
        
        # 初期化
        self._setup_window()
        self._load_data()
        self._create_ui()
        self._show_month(self.current_month)
        
        # ウィンドウクローズ時の処理を設定
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # グローバルキーボードショートカット
        self.root.bind('<Control-f>', lambda e: SearchDialog(self.root, self))

        # コピー＆ペーストのショートカット
        self.root.bind('<Control-c>', self._copy_cells)
        self.root.bind('<Control-x>', self._cut_cells)
        self.root.bind('<Control-v>', self._paste_cells)
        self.root.bind('<Delete>', self._delete_cells)
        self.root.bind('<Control-z>', self._undo)
    
    def _get_color_theme(self):
        """カラーテーマを取得"""
        return {
            'bg_primary': ColorTheme.BG_PRIMARY,
            'bg_secondary': ColorTheme.BG_SECONDARY,
            'bg_tertiary': ColorTheme.BG_TERTIARY,
            'accent': ColorTheme.ACCENT,
            'accent_green': ColorTheme.ACCENT_GREEN,
            'accent_red': ColorTheme.ACCENT_RED,
            'text_primary': ColorTheme.TEXT_PRIMARY,
            'text_secondary': ColorTheme.TEXT_SECONDARY,
            'border': ColorTheme.BORDER,
            'hover': ColorTheme.HOVER
        }
    
    def _setup_window(self):
        """メインウィンドウの基本設定を行う"""
        self.root.title("💰 家計管理 2025")
        
        # ウィンドウサイズと位置の設定
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        x = (screen_width - WindowConfig.WIDTH) // 2
        y = (screen_height - WindowConfig.HEIGHT) // 2
        
        self.root.geometry(f"{WindowConfig.WIDTH}x{WindowConfig.HEIGHT}+{x}+{y}")
        self.root.minsize(WindowConfig.MIN_WIDTH, WindowConfig.MIN_HEIGHT)
        self.root.resizable(*WindowConfig.RESIZABLE)
        
        self.root.configure(bg=self.colors['bg_primary'])
        self._setup_styles()
    
    def _setup_styles(self):
        """ttkウィジェットのカスタムスタイルを定義する"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # 通常のボタンスタイル
        style.configure('Modern.TButton',
                        background=self.colors['bg_secondary'],
                        foreground=self.colors['text_primary'],
                        borderwidth=0,
                        focuscolor='none',
                        font=FontConfig.BUTTON,
                        relief='flat')
        
        style.map('Modern.TButton',
                  background=[('active', self.colors['hover']),
                              ('pressed', self.colors['accent'])],
                  foreground=[('active', '#ffffff')])
        
        # アクセントボタンスタイル
        style.configure('Accent.TButton',
                        background=self.colors['accent'],
                        foreground='#ffffff',
                        borderwidth=0,
                        focuscolor='none',
                        font=('Segoe UI', 10, 'bold'),
                        relief='flat')
        
        style.map('Accent.TButton',
                  background=[('active', self.colors['hover']),
                              ('pressed', '#4dabf7')])
        
        # 月表示ボタンスタイル
        style.configure('Month.TButton',
                        background=self.colors['accent'],
                        foreground='#ffffff',
                        borderwidth=0,
                        focuscolor='none',
                        font=FontConfig.BUTTON_LARGE,
                        relief='flat')
        
        style.map('Month.TButton',
                  background=[('active', self.colors['hover']),
                              ('pressed', '#4dabf7')])
        
        # 選択された月ボタンスタイル
        style.configure('Selected.TButton',
                        background=self.colors['accent'],
                        foreground='#ffffff',
                        borderwidth=0,
                        focuscolor='none',
                        font=('Segoe UI', 9, 'bold'),
                        relief='flat')
        
        # ナビゲーションボタンスタイル
        style.configure('Nav.TButton',
                        background=self.colors['bg_tertiary'],
                        foreground=self.colors['text_primary'],
                        borderwidth=0,
                        focuscolor='none',
                        font=FontConfig.BUTTON_LARGE,
                        relief='flat')
        
        style.map('Nav.TButton',
                  background=[('active', self.colors['accent']),
                              ('pressed', self.colors['hover'])])
    
    def _load_data(self):
        """データと設定を読み込む"""
        self.data_manager.load_settings()
        self.data_manager.load_data()
    
    def _save_data(self):
        """データと設定を保存する"""
        self.data_manager.save_data()
        self.data_manager.save_settings()
    
    def _on_closing(self):
        """ウィンドウが閉じられる時の処理"""
        self._save_data()
        self.data_manager.save_backup()
        self.root.destroy()
    
    def _create_ui(self):
        """メインウィンドウのUI要素を作成する"""
        # メインコンテナ
        main_container = tk.Frame(self.root, bg=self.colors['bg_primary'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        # ヘッダーセクション
        header = tk.Frame(main_container, bg=self.colors['bg_secondary'])
        header.pack(fill=tk.X, pady=(0, 8))
        
        header_inner = tk.Frame(header, bg=self.colors['bg_secondary'])
        header_inner.pack(fill=tk.X, padx=15, pady=8)
        
        # 年選択コントロール(左側)
        self._create_year_controls(header_inner)
        
        # 月選択ボタン(1月～12月)
        self._create_month_buttons(header_inner)
        
        # 検索ボタン
        self._create_search_button(header_inner)
        
        # 図表ボタン
        self._create_chart_button(header_inner)
        
        # 現在月表示(右側、クリック可能)
        self._create_current_month_button(header_inner)
        
        # メインテーブルセクション
        tree_section = tk.Frame(main_container, bg=self.colors['bg_secondary'])
        tree_section.pack(fill=tk.BOTH, expand=True)
        
        self._create_treeview(tree_section)
        self._update_month_buttons()
    
    def _create_year_controls(self, parent):
        """年選択コントロールを作成"""
        year_container = tk.Frame(parent, bg=self.colors['bg_secondary'])
        year_container.pack(side=tk.LEFT)
        
        year_nav = tk.Frame(year_container, bg=self.colors['bg_secondary'])
        year_nav.pack()
        
        # 前年ボタン
        ttk.Button(year_nav, text="◀", width=3, style='Nav.TButton',
                command=self._prev_month).pack(side=tk.LEFT, padx=(0, 4))
        
        # 年表示（クリックでダイアログ表示）
        year_display = tk.Frame(year_nav, bg=self.colors['bg_tertiary'], 
                                cursor='hand2')
        year_display.pack(side=tk.LEFT, padx=4)
        
        self.year_label = tk.Label(year_display, text=str(self.current_year),
                                font=FontConfig.TITLE,
                                bg=self.colors['bg_tertiary'],
                                fg=self.colors['text_primary'],
                                padx=12, pady=4,
                                cursor='hand2')
        self.year_label.pack()
        
        # クリックでダイアログを表示
        self.year_label.bind('<Button-1>', self._open_year_input_dialog)
        year_display.bind('<Button-1>', self._open_year_input_dialog)
        
        # 翌年ボタン
        ttk.Button(year_nav, text="▶", width=3, style='Nav.TButton',
                command=self._next_month).pack(side=tk.LEFT, padx=(4, 0))
    
    def _create_month_buttons(self, parent):
        """月選択ボタンを作成"""
        month_container = tk.Frame(parent, bg=self.colors['bg_secondary'])
        month_container.pack(side=tk.LEFT, padx=(20, 0))
        
        self.month_buttons = []
        for m in range(1, 13):
            btn = ttk.Button(month_container, text=f"{m:02d}", width=4, style='Modern.TButton',
                             command=lambda mo=m: self.select_month(mo))
            btn.pack(side=tk.LEFT, padx=1)
            self.month_buttons.append(btn)
    
    def _create_search_button(self, parent):
        """検索ボタンを作成"""
        search_container = tk.Frame(parent, bg=self.colors['bg_secondary'])
        search_container.pack(side=tk.LEFT, padx=(20, 0))
        
        ttk.Button(search_container, text="🔍 検索 (Ctrl+F)", width=15, style='Accent.TButton',
                   command=lambda: SearchDialog(self.root, self)).pack()
    
    def _create_chart_button(self, parent):
        """図表ボタンを作成"""
        chart_container = tk.Frame(parent, bg=self.colors['bg_secondary'])
        chart_container.pack(side=tk.LEFT, padx=(10, 0))
        
        ttk.Button(chart_container, text="📊 図表", width=10, style='Accent.TButton',
                   command=lambda: ChartDialog(self.root, self)).pack()
    
    def _create_current_month_button(self, parent):
        """現在月表示ボタンを作成"""
        month_info = tk.Frame(parent, bg=self.colors['bg_secondary'])
        month_info.pack(side=tk.RIGHT)
        
        self.current_month_button = ttk.Button(month_info,
                                               text=f"📅 {self.current_month:02d}月",
                                               style='Month.TButton',
                                               command=self._open_monthly_data)
        self.current_month_button.pack()
    
    def _create_treeview(self, parent):
        """メインのTreeviewウィジェットを作成する"""
        if self.tree:
            return
        
        # 既存のウィジェットをクリア
        for widget in parent.winfo_children():
            widget.destroy()
        
        # 列の定義
        all_columns = self.get_all_columns()
        columns_with_button = all_columns + ["+"]
        
        # Treeviewを格納するフレーム
        tree_frame = tk.Frame(parent, bg='white', relief='solid', bd=1)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Treeviewの作成
        self.tree = ttk.Treeview(tree_frame, columns=columns_with_button, show="headings", height=25)
        self.tree.grid(row=0, column=0, sticky="nsew")

        # ヘッダーフォントの計測用オブジェクトを作成
        # FontConfig.HEADING の設定 ('Arial', 10, 'bold') を使用
        heading_font = tkfont.Font(root=self.root, font=FontConfig.HEADING)
        
        # 各列の設定
        self.default_column_widths = {}
        for i, col in enumerate(columns_with_button):
            self.tree.heading(col, text=col)
            
            # 幅の計算ロジックを変更
            if i == 0:  # 日付列
                width = TreeviewConfig.COL_WIDTH_DATE
                min_w = 50
                stretch_opt = True
            elif col == "+":  # 追加ボタン列
                width = TreeviewConfig.COL_WIDTH_BUTTON
                min_w = 40
                stretch_opt = False
            else:  # データ列
                # タイトルの文字幅を計測し、左右にパディング(+20px)を追加
                title_width = heading_font.measure(col) + 20
                # デフォルト幅(80px)とタイトル幅の大きい方を採用
                width = max(TreeviewConfig.COL_WIDTH_DATA, title_width)
                min_w = 60
                stretch_opt = True
            
            # 列設定を適用
            self.tree.column(col, anchor="center", width=width, minwidth=min_w, stretch=stretch_opt)
            self.default_column_widths[col] = width
        
        # スクロールバー(縦)
        v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=v_scrollbar.set)
        
        # スクロールバー(横)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        self.tree.configure(xscrollcommand=h_scrollbar.set)
        
        # Treeviewのスタイル設定
        self._configure_treeview_style()
        
        # 行のタグ設定
        self.tree.tag_configure(TreeviewConfig.TAG_TOTAL,
                                background=TreeviewConfig.BG_TOTAL,
                                font=FontConfig.HEADING)
        self.tree.tag_configure(TreeviewConfig.TAG_SUMMARY,
                                background=TreeviewConfig.BG_SUMMARY,
                                font=FontConfig.HEADING)
        self.tree.tag_configure(TreeviewConfig.TAG_NORMAL,
                                background=TreeviewConfig.BG_NORMAL)
        self.tree.tag_configure(TreeviewConfig.TAG_ODD,
                                background=TreeviewConfig.BG_ODD)
        self.tree.tag_configure(TreeviewConfig.TAG_SAT,
                                background=TreeviewConfig.BG_SAT)
        self.tree.tag_configure(TreeviewConfig.TAG_SUN,
                                background=TreeviewConfig.BG_SUN)
        
        # イベントバインド
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Button-1>", self._on_single_click)
        self.tree.bind("<Button-3>", self._on_right_click)
        self.tree.bind("<MouseWheel>", self._on_mousewheel)
        self.tree.bind("<Shift-MouseWheel>",
                       lambda e: self.tree.xview_scroll(int(-1 * (e.delta / 120)), "units"))
        self.tree.bind("<space>", self._on_space_key)
        
        # 右クリックメニュー(カスタム列用)
        self.column_context_menu = tk.Menu(self.root, tearoff=0)
        
        # ツールチップを初期化
        self.tooltip = TreeviewTooltip(self.tree, self)

    def _open_year_input_dialog(self, event=None):
        """年入力ダイアログを開く"""
        # ダイアログを作成
        dialog = tk.Toplevel(self.root)
        dialog.title("年を入力")
        dialog.resizable(False, False)
        
        # ダイアログのサイズと位置
        dialog_width = 300
        dialog_height = 150
        
        # 親ウィンドウの中央に配置
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog_width) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog_height) // 2
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        
        # モーダルダイアログに設定
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 背景色
        dialog.configure(bg='#f0f0f0')
        
        # タイトル
        title_label = tk.Label(dialog, text="移動する年を入力してください",
                            font=('Arial', 12, 'bold'),
                            bg='#f0f0f0')
        title_label.pack(pady=(20, 10))
        
        # 入力フレーム
        input_frame = tk.Frame(dialog, bg='#f0f0f0')
        input_frame.pack(pady=10)
        
        # ラベル
        tk.Label(input_frame, text="年:", font=('Arial', 11),
                bg='#f0f0f0').pack(side=tk.LEFT, padx=(0, 10))
        
        # テキストボックス
        year_entry = tk.Entry(input_frame, font=('Arial', 14),
                            width=10, justify='center')
        year_entry.pack(side=tk.LEFT)
        year_entry.insert(0, str(self.current_year))
        year_entry.select_range(0, tk.END)
        year_entry.focus_set()
        
        # ボタンフレーム
        button_frame = tk.Frame(dialog, bg='#f0f0f0')
        button_frame.pack(pady=(10, 20))
        
        def on_ok():
            """OKボタンの処理"""
            try:
                new_year = int(year_entry.get().strip())
                
                # 妥当な範囲かチェック
                if 1900 <= new_year <= 2100:
                    self.current_year = new_year
                    self.update_year_display()
                    self._update_month_buttons()
                    self._show_month(self.current_month)
                    dialog.destroy()
                else:
                    messagebox.showwarning("警告",
                                        "年は1900～2100の範囲で入力してください。",
                                        parent=dialog)
                    year_entry.select_range(0, tk.END)
                    year_entry.focus_set()
            except ValueError:
                messagebox.showwarning("警告",
                                    "正しい年数を入力してください。",
                                    parent=dialog)
                year_entry.select_range(0, tk.END)
                year_entry.focus_set()
        
        def on_cancel():
            """キャンセルボタンの処理"""
            dialog.destroy()
        
        # OKボタン
        ok_button = tk.Button(button_frame, text="OK", font=('Arial', 11),
                            bg='#2196f3', fg='white',
                            width=10, command=on_ok)
        ok_button.pack(side=tk.LEFT, padx=5)
        
        # キャンセルボタン
        cancel_button = tk.Button(button_frame, text="キャンセル",
                                font=('Arial', 11),
                                bg='#f44336', fg='white',
                                width=10, command=on_cancel)
        cancel_button.pack(side=tk.LEFT, padx=5)
        
        # Enterキーでも確定できるようにする
        year_entry.bind('<Return>', lambda e: on_ok())
        year_entry.bind('<KP_Enter>', lambda e: on_ok())  # テンキーのEnter
        
        # Escapeキーでキャンセル
        dialog.bind('<Escape>', lambda e: on_cancel())
    
    def _configure_treeview_style(self):
        """Treeviewのスタイルを設定"""
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure("Treeview",
                        fieldbackground="white",
                        background="white",
                        rowheight=TreeviewConfig.ROW_HEIGHT,
                        font=FontConfig.DEFAULT,
                        borderwidth=1,
                        relief="solid")
        
        style.configure("Treeview.Heading",
                        background="#e8e8e8",
                        font=FontConfig.HEADING,
                        relief="raised",
                        borderwidth=1)
        
        style.map("Treeview",
                  background=[('selected', '#0078d4')],
                  foreground=[('selected', 'yellow')])
    
    def _prev_month(self):
        """前月に移動する"""
        self.current_month -= 1
        if self.current_month < 1:
            self.current_month = 12
            self.current_year -= 1
        self.update_year_display()
        self._update_month_buttons()
        self._show_month(self.current_month)
    
    def _next_month(self):
        """翌月に移動する"""
        self.current_month += 1
        if self.current_month > 12:
            self.current_month = 1
            self.current_year += 1
        self.update_year_display()
        self._update_month_buttons()
        self._show_month(self.current_month)
    
    def update_year_display(self):
        """年表示を更新"""
        if self.year_label:
            self.year_label.config(text=str(self.current_year))
    
    def select_month(self, month):
        """指定された月を選択する"""
        self.current_month = month
        self.current_month_button.config(text=f"📅 {month:02d}月")
        self._update_month_buttons()
        self._show_month(month)

    def navigate_to_cell(self, day, col_index):
        """
        指定された日・列のセルに移動して選択状態にする。
        検索ダイアログや月間データダイアログから呼び出す用。

        Args:
            day: 移動先の日（0=まとめ行）
            col_index: 移動先の列インデックス
        """
        if not self.tree:
            return

        items = self.tree.get_children()
        if not items:
            return

        target_item = None

        if day == 0:
            target_item = items[-1]
        else:
            for item in items[:-2]:
                values = self.tree.item(item, 'values')
                if values and str(values[0]).strip().split('(')[0].strip() == str(day):
                    target_item = item
                    break

        if target_item:
            # <Button-1>イベントを一時的に無効化して選択が上書きされないようにする
            self.tree.unbind("<Button-1>")

            self.tree.selection_set(target_item)
            self.tree.see(target_item)
            self.tree.focus(target_item)

            col_id = f"#{col_index + 1}"
            self.selected_column_id = col_id
            self.selection_start_row = target_item
            self.selection_start_col = col_id
            self.ctrl_selected_cells = [(target_item, col_id)]

            # 次のイベントループでバインドを復元
            self.root.after(100, lambda: self.tree.bind("<Button-1>", self._on_single_click))
    
    def _update_month_buttons(self):
        """月選択ボタンのハイライトを更新する"""
        for i, btn in enumerate(self.month_buttons, start=1):
            if i == self.current_month:
                btn.configure(style='Selected.TButton')
            else:
                btn.configure(style='Modern.TButton')
    
    def _open_monthly_data(self):
        """月間データ詳細ダイアログを開く"""
        MonthlyDataDialog(self.root, self, self.current_year, self.current_month)
    
    def get_all_columns(self):
        """全ての列(デフォルト + カスタム)を取得"""
        return DefaultColumns.ITEMS + self.data_manager.custom_columns
    
    def get_days_in_month(self):
        """現在の月の日数を取得"""
        return get_days_in_month(self.current_year, self.current_month)
    
    def _show_month(self, month):
        """指定された月のデータを表示する"""
        if not self.tree:
            return
        
        # 既存の表示をクリア
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        all_columns = self.get_all_columns()
        days = self.get_days_in_month()
        
        # 各日のデータを表示
        for day in range(1, days + 1):
            row_values = self._calculate_day_totals(day)
            formatted_values = self._format_row_values(row_values)
            formatted_values.append("")  # +ボタン列
            
            # 土日・奇数偶数行で背景色を変える
            weekday = datetime.date(self.current_year, self.current_month, day).weekday()
            if weekday == 5:  # 5=土
                tag = TreeviewConfig.TAG_SAT
            elif weekday == 6:  # 6=日
                tag = TreeviewConfig.TAG_SUN
            else:
                tag = TreeviewConfig.TAG_ODD if day % 2 == 1 else TreeviewConfig.TAG_NORMAL

            self.tree.insert("", "end", values=formatted_values, tags=(tag,))
        
        # 合計行
        total_row = [" 合計 "] + ["  "] * (len(all_columns) - 1) + [""]
        self.tree.insert("", "end", values=total_row, tags=(TreeviewConfig.TAG_TOTAL,))
        
        # まとめ行(収入・支出の表示)
        income_val = self._get_income_total()
        inc_str = f" {income_val} " if income_val != 0 else "  "
        summary_row = [" まとめ ", "  ", " 収入 ", inc_str, " 支出 ", "  "] + \
                      ["  "] * (len(all_columns) - 6) + [""]
        self.tree.insert("", "end", values=summary_row, tags=(TreeviewConfig.TAG_SUMMARY,))
        
        # 合計とまとめ行の値を更新
        self._update_totals()
    
    def _calculate_day_totals(self, day):
        """特定の日の各項目の合計金額を計算する"""
        all_columns = self.get_all_columns()
        totals = [""] * len(all_columns)
        
        weekdays = ["月", "火", "水", "木", "金", "土", "日"]
        wd = datetime.date(self.current_year, self.current_month, day).weekday()
        totals[0] = f"{day}({weekdays[wd]})"  # 日付列
        
        # 各項目の合計を計算
        for col_index in range(1, len(all_columns)):
            dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_index}"
            data_list = self.data_manager.get_transaction_data(dict_key)
            if data_list:
                # 金額列(インデックス1)を合計
                total = sum(parse_amount(row[1]) for row in data_list if len(row) > 1)
                if total != 0:
                    totals[col_index] = str(total)
        
        return totals
    
    def _format_row_values(self, values):
        """行データを表示用にフォーマットする"""
        formatted = []
        for i, val in enumerate(values):
            if i == 0:  # 日付列
                formatted.append(f" {val} ")
            else:
                formatted.append(f" {val} " if val else "  ")
        return formatted
    
    def _get_income_total(self):
        """現在月の収入合計を取得する"""
        dict_key = f"{self.current_year}-{self.current_month}-0-3"
        data_list = self.data_manager.get_transaction_data(dict_key)
        if data_list:
            return sum(parse_amount(row[1]) for row in data_list if len(row) > 1)
        return 0
    
    def _update_totals(self):
        """合計行とまとめ行の値を更新する"""
        items = self.tree.get_children()
        if len(items) < 2:
            return
        
        total_row_id = items[-2]  # 合計行
        summary_row_id = items[-1]  # まとめ行
        all_columns = self.get_all_columns()
        cols = len(all_columns)
        
        # 各列の合計を計算
        sums = [0] * (cols - 1)
        for row_id in items[:-2]:  # 日付行のみ対象
            row_vals = self.tree.item(row_id, 'values')
            for i in range(1, cols):
                try:
                    val_str = str(row_vals[i]).strip() if i < len(row_vals) else ""
                    sums[i - 1] += int(val_str) if val_str else 0
                except (ValueError, TypeError, IndexError):
                    pass
        
        # 合計行を更新
        total_vals = list(self.tree.item(total_row_id, 'values'))
        for i in range(1, cols):
            total_vals[i] = f" {sums[i - 1]} " if sums[i - 1] != 0 else "  "
        
        while len(total_vals) <= cols:
            total_vals.append("")
        self.tree.item(total_row_id, values=total_vals)
        
        # 総支出を計算
        grand_total = sum(int(str(v).strip()) for v in total_vals[1:cols]
                          if v and str(v).strip() and str(v).strip().lstrip('-').isdigit())
        
        # まとめ行を更新
        summary_vals = list(self.tree.item(summary_row_id, 'values'))
        try:
            income_str = str(summary_vals[3]).strip() if len(summary_vals) > 3 else ""
            income_val = int(income_str) if income_str else 0
        except:
            income_val = 0
        
        # 収支差額と総支出を更新
        balance = income_val - grand_total
        summary_vals[1] = f" {balance} " if balance != 0 else "  "
        summary_vals[5] = f" {grand_total} " if grand_total != 0 else "  "
        
        while len(summary_vals) <= cols:
            summary_vals.append("")
        
        self.tree.item(summary_row_id, values=summary_vals)
    
    def _on_single_click(self, event):
        """シングルクリックイベントを処理する"""
        region = self.tree.identify_region(event.x, event.y)

        # クリックされた列IDを取得して保存（コピー＆ペースト用）
        col_id = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        
        if col_id:
            self.selected_column_id = col_id
        
        # Shift+クリックの場合は範囲選択
        if event.state & 0x1 and row_id and col_id:  # Shiftキー
            if self.selection_start_row and self.selection_start_col:
                # 範囲選択を実行
                self._select_range(self.selection_start_row, self.selection_start_col, row_id, col_id)
                # Ctrl選択リストをクリア
                self.ctrl_selected_cells = []
                return
        
        # Ctrl+クリックの場合は個別選択モード
        if event.state & 0x4 and row_id and col_id:  # Ctrlキー
            # このセルをCtrl選択リストに追加（重複チェック）
            cell_tuple = (row_id, col_id)
            if cell_tuple in self.ctrl_selected_cells:
                # 既に選択されている場合は削除（トグル）
                self.ctrl_selected_cells.remove(cell_tuple)
            else:
                # 新規追加
                self.ctrl_selected_cells.append(cell_tuple)
            return
        
        # 通常のクリック（Shift/Ctrl押下なし）の場合
        if row_id and col_id:
            # 範囲選択の開始点を記録
            self.selection_start_row = row_id
            self.selection_start_col = col_id
            # Ctrl選択リストをクリア
            self.ctrl_selected_cells = [(row_id, col_id)]

        if region == "heading":
            if col_id:
                col_index = int(col_id[1:]) - 1
                all_columns = self.get_all_columns()
                
                if col_index == len(all_columns):  # +ボタン列
                    self._add_column()
    
    def _select_range(self, start_row_id, start_col_id, end_row_id, end_col_id):
        """
        開始セルと終了セルの間の矩形範囲を選択する
        
        Args:
            start_row_id: 開始行ID
            start_col_id: 開始列ID（"#1", "#2"など）
            end_row_id: 終了行ID
            end_col_id: 終了列ID
        """
        items = self.tree.get_children()
        
        # 行のインデックスを取得
        try:
            start_row_idx = items.index(start_row_id)
            end_row_idx = items.index(end_row_id)
        except ValueError:
            return
        
        # 列のインデックスを取得
        start_col_idx = int(start_col_id[1:]) - 1
        end_col_idx = int(end_col_id[1:]) - 1
        
        # 開始と終了を正規化（小さい方が先）
        if start_row_idx > end_row_idx:
            start_row_idx, end_row_idx = end_row_idx, start_row_idx
        if start_col_idx > end_col_idx:
            start_col_idx, end_col_idx = end_col_idx, start_col_idx
        
        # 範囲内のすべての行を選択
        selected_rows = []
        for i in range(start_row_idx, end_row_idx + 1):
            if i < len(items):
                selected_rows.append(items[i])
        
        # Treeviewの選択を更新
        self.tree.selection_set(selected_rows)
    
    def _on_double_click(self, event):
        """ダブルクリックイベントを処理する"""
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        
        if not row_id or not col_id:
            return
        
        # ヘッダーのクリック処理
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            col_index = int(col_id[1:]) - 1
            all_columns = self.get_all_columns()
            
            if col_index == len(all_columns):  # +ボタン
                self._add_column()
            elif col_index >= len(DefaultColumns.ITEMS):  # カスタム列
                self._edit_column_name(col_index)
            return
        
        # 行の種類を判定
        items = self.tree.get_children()
        if len(items) < 2:
            return
        
        total_row_id = items[-2]
        summary_row_id = items[-1]
        
        # 合計行は編集不可
        if row_id == total_row_id:
            return
        
        # まとめ行は収入列のみ編集可能
        if row_id == summary_row_id and col_id != "#4":
            return
        
        # 日付列は編集不可
        if col_id == "#1":
            return
        
        # +ボタン列は編集不可
        col_index = int(col_id[1:]) - 1
        all_columns = self.get_all_columns()
        if col_index >= len(all_columns):
            return
        
        # 行データを取得
        row_vals = self.tree.item(row_id, 'values')
        if not row_vals:
            return
        
        # 日付と列名を特定
        if row_id == summary_row_id:
            day = 0
            col_index = 3  # 収入列
            col_name = "収入"
        else:
            try:
                m = re.search(r'\d+', str(row_vals[0]))
                if not m:
                    return
                day = int(m.group())
            except:
                return
            col_name = self.tree.heading(col_id, "text")
        
        # 取引詳細ダイアログを開く
        dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_index}"
        
        # ダイアログを開く前にデータを保存
        old_data = self.data_manager.get_transaction_data(dict_key)
        
        # ダイアログを開く
        dialog = TransactionDialog(self.root, self, dict_key, col_name)
        
        # ダイアログが閉じた後、データが変更されていれば元に戻すスタックに保存
        self.root.wait_window(dialog)
        new_data = self.data_manager.get_transaction_data(dict_key)
        
        # データが変更されていれば記録
        if old_data != new_data:
            self._save_undo_state('edit_detail', [(dict_key, old_data[:] if old_data else None)])
    
    def _on_right_click(self, event):
        """右クリックイベントを処理する"""
        region = self.tree.identify_region(event.x, event.y)

        # ヘッダー以外（セル）での右クリックの場合、コピペメニューを表示
        if region != "heading":
            # クリック位置の行と列を選択状態にする
            row_id = self.tree.identify_row(event.y)
            col_id = self.tree.identify_column(event.x)
            
            if row_id and col_id:
                # 選択状態を更新
                self.tree.selection_set(row_id)
                self.tree.focus(row_id)
                self.selected_column_id = col_id
                
                # コンテキストメニュー作成
                cell_menu = tk.Menu(self.root, tearoff=0)
                cell_menu.add_command(label="元に戻す (Ctrl+Z)", command=self._undo)
                cell_menu.add_separator()
                cell_menu.add_command(label="切り取り (Ctrl+X)", command=self._cut_cells)
                cell_menu.add_command(label="コピー (Ctrl+C)", command=self._copy_cells)
                cell_menu.add_command(label="貼り付け (Ctrl+V)", command=self._paste_cells)
                cell_menu.add_separator()
                cell_menu.add_command(label="削除 (Delete)", command=self._delete_cells)
                cell_menu.post(event.x_root, event.y_root)
            return
        
        col_id = self.tree.identify_column(event.x)
        if not col_id:
            return
        
        col_index = int(col_id[1:]) - 1
        all_columns = self.get_all_columns()
        
        # 右クリックメニューを再作成
        self.column_context_menu = tk.Menu(self.root, tearoff=0)
        
        # カスタム列の場合は編集・削除オプションを追加
        if len(all_columns) > col_index >= len(DefaultColumns.ITEMS) and col_index != 0:
            self.selected_column_index = col_index
            self.column_context_menu.add_command(label="列名を編集", command=self._edit_column_name)
            self.column_context_menu.add_separator()
            self.column_context_menu.add_command(label="列を削除", command=self._delete_column)
            self.column_context_menu.add_separator()
        
        # すべての列で列幅リセットを利用可能
        self.column_context_menu.add_command(label="全ての列幅をリセット",
                                             command=self._reset_all_column_widths)
        
        self.column_context_menu.post(event.x_root, event.y_root)
    
    def _on_mousewheel(self, event):
        """マウスホイールイベントを処理する"""
        if event.state & 0x4:  # Ctrlキーが押されている
            self.tree.xview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            self.tree.yview_scroll(int(-1 * (event.delta / 120)), "units")
    
    def _on_space_key(self, event):
        """
        SPACEキーが押された時の処理
        選択中のセルの詳細ダイアログを開く
        """
        selected_items = self.tree.selection()
        if not selected_items:
            return
        
        row_id = selected_items[0]
        
        # フォーカスされている列を取得
        focus_item = self.tree.focus()
        if not focus_item:
            return
        
        # 現在選択されている列のインデックスを取得
        # （Treeviewは列のフォーカスを直接取得できないため、
        # 最後にクリックされた列を使用）
        if not hasattr(self, 'selected_column_id') or not self.selected_column_id:
            return
        
        col_id = self.selected_column_id
        col_index = int(col_id[1:]) - 1
        
        # 編集不可のセルをチェック
        items = self.tree.get_children()
        if len(items) < 2:
            return
        
        total_row_id = items[-2]
        summary_row_id = items[-1]
        
        # 合計行は編集不可
        if row_id == total_row_id:
            return
        
        # まとめ行は収入列のみ編集可能
        if row_id == summary_row_id and col_id != "#4":
            return
        
        # 日付列は編集不可
        if col_id == "#1":
            return
        
        # +ボタン列は編集不可
        all_columns = self.get_all_columns()
        if col_index >= len(all_columns):
            return
        
        # 行データを取得
        row_vals = self.tree.item(row_id, 'values')
        if not row_vals:
            return
        
        # 日付と列名を特定
        if row_id == summary_row_id:
            day = 0
            col_index = 3  # 収入列
            col_name = "収入"
        else:
            try:
                m = re.search(r'\d+', str(row_vals[0]))
                if not m:
                    return
                day = int(m.group())
            except:
                return
            col_name = self.tree.heading(col_id, "text")
        
        # 取引詳細ダイアログを開く
        dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_index}"
        TransactionDialog(self.root, self, dict_key, col_name)
    
    def _reset_all_column_widths(self):
        """指定された列の幅をデフォルトにリセットする"""
        all_columns = self.get_all_columns() + ["+"]
        for i, col_name in enumerate(all_columns):
            col_id = f"#{i + 1}"
            if col_name in self.default_column_widths:
                self.tree.column(col_id, width=self.default_column_widths[col_name])

    def _add_column(self):
        """新しい列を追加する"""
        dialog = tk.Toplevel(self.root)
        dialog.title("列の追加")
        dialog.resizable(False, False)
        
        # ダイアログを中央に配置
        dialog_width = 300
        dialog_height = 120
        
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog_width) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog_height) // 2
        
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text="新しい列名を入力してください:", font=('Arial', 11)).pack(pady=10)
        
        entry = tk.Entry(dialog, font=('Arial', 11), width=25)
        entry.pack(pady=5)
        entry.focus_set()
        
        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=10)
        
        def on_ok():
            column_name = entry.get().strip()
            if column_name:
                all_columns = self.get_all_columns()
                if column_name not in all_columns:
                    self.data_manager.add_custom_column(column_name)
                    dialog.destroy()
                    self._recreate_treeview()
                    self._show_month(self.current_month)
                else:
                    messagebox.showwarning("警告", "その列名は既に存在します。", parent=dialog)
            else:
                messagebox.showwarning("警告", "列名を入力してください。", parent=dialog)
        
        tk.Button(button_frame, text="OK", command=on_ok, width=8).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="キャンセル", command=dialog.destroy, width=8).pack(side=tk.LEFT, padx=5)
        
        entry.bind('<Return>', lambda e: on_ok())
        dialog.bind('<Escape>', lambda e: dialog.destroy())

    def _recreate_treeview(self):
        """Treeviewを再作成する"""
        if self.tree:
            tree_parent = self.tree.master
            self.tree.destroy()
            self.tree = None
            self._create_treeview(tree_parent)

    def _edit_column_name(self, col_index=None):
        """カスタム列の名前を編集する"""
        if col_index is None:
            col_index = getattr(self, 'selected_column_index', None)
        
        if col_index is None or col_index < len(DefaultColumns.ITEMS):
            return
        
        custom_index = col_index - len(DefaultColumns.ITEMS)
        if custom_index >= len(self.data_manager.custom_columns):
            return
        
        old_name = self.data_manager.custom_columns[custom_index]
        
        # 編集ダイアログを表示
        dialog = tk.Toplevel(self.root)
        dialog.title("列名の編集")
        dialog.resizable(False, False)
        
        dialog_width = 300
        dialog_height = 120
        
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog_width) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog_height) // 2
        
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text="新しい列名を入力してください:", font=('Arial', 11)).pack(pady=10)
        
        entry = tk.Entry(dialog, font=('Arial', 11), width=25)
        entry.pack(pady=5)
        entry.insert(0, old_name)
        entry.select_range(0, tk.END)
        entry.focus_set()
        
        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=10)
        
        def on_ok():
            new_name = entry.get().strip()
            if new_name and new_name != old_name:
                all_columns = self.get_all_columns()
                if new_name not in all_columns:
                    self.data_manager.edit_custom_column(old_name, new_name)
                    dialog.destroy()
                    self._recreate_treeview()
                    self._show_month(self.current_month)
                else:
                    messagebox.showwarning("警告", "その列名は既に存在します。", parent=dialog)
            else:
                dialog.destroy()
        
        tk.Button(button_frame, text="OK", command=on_ok, width=8).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="キャンセル", command=dialog.destroy, width=8).pack(side=tk.LEFT, padx=5)
        
        entry.bind('<Return>', lambda e: on_ok())
        dialog.bind('<Escape>', lambda e: dialog.destroy())

    def _delete_column(self):
        """カスタム列を削除する"""
        col_index = getattr(self, 'selected_column_index', None)
        if col_index is None or col_index < len(DefaultColumns.ITEMS):
            return
        
        custom_index = col_index - len(DefaultColumns.ITEMS)
        if custom_index >= len(self.data_manager.custom_columns):
            return
        
        col_name = self.data_manager.custom_columns[custom_index]
        
        # 削除確認ダイアログ
        if messagebox.askyesno("確認", f"列 '{col_name}' を削除しますか?\n※この列のデータもすべて削除されます。"):
            # 列をリストから削除
            self.data_manager.delete_custom_column(col_name)
            
            # 関連するデータを削除
            self.data_manager.delete_column_data(col_index)
            
            # Treeviewを再作成して変更を反映
            self._recreate_treeview()
            self._show_month(self.current_month)

    def update_parent_cell(self, dict_key_day, col_index, new_value):
        """親画面のセル表示を更新する"""
        # キーから年月日を抽出
        y, mo, d = dict_key_day.split("-")
        y, mo, d = int(y), int(mo), int(d)
        
        # 現在表示中の年月と一致する場合のみ更新
        if (self.current_year == y) and (self.current_month == mo):
            items = self.tree.get_children()
            if len(items) < 2:
                return
            
            summary_row_id = items[-1]  # まとめ行
            
            # 該当する日付の行を検索
            for row_id in items[:-2]:  # 日付行のみ対象
                row_vals = list(self.tree.item(row_id, 'values'))
                m = re.search(r'\d+', str(row_vals[0])) if row_vals else None
                if m and int(m.group()) == d:
                    # 列数を確認して必要に応じて拡張
                    all_columns = self.get_all_columns()
                    while len(row_vals) < len(all_columns) + 1:
                        row_vals.append("")
                    
                    # 表示値をフォーマット(パディング付き)
                    display_value = "  "
                    if new_value and str(new_value).strip() != "" and str(new_value) != "0":
                        display_value = f" {new_value} "
                    
                    # 値を更新
                    row_vals[col_index] = display_value
                    self.tree.item(row_id, values=row_vals)
                    break
            
            # まとめ行(収入)の更新
            if d == 0:
                sum_vals = list(self.tree.item(summary_row_id, 'values'))
                all_columns = self.get_all_columns()
                while len(sum_vals) < len(all_columns) + 1:
                    sum_vals.append("")
                
                display_value = "  "
                if new_value and str(new_value).strip() != "" and str(new_value) != "0":
                    display_value = f" {new_value} "
                
                sum_vals[col_index] = display_value
                self.tree.item(summary_row_id, values=sum_vals)
            
            # 合計とまとめ行を再計算
            self._update_totals()

    def _get_selected_cells(self):
        """
        現在選択されているセルの情報を取得
        
        Shift選択の場合：矩形範囲内のすべてのセル
        Ctrl選択の場合：個別に選択されたセルのみ
        通常選択：現在の行と列
        
        Returns:
            list: [(row_id, col_id, day, col_idx), ...]
        """
        selected_items = self.tree.selection()
        if not selected_items:
            return []
        
        cells = []
        items = self.tree.get_children()
        total_row_id = items[-2] if len(items) >= 2 else None
        summary_row_id = items[-1] if len(items) >= 1 else None
        all_columns = self.get_all_columns()
        
        # Ctrl選択の場合：個別に記録されたセルを使用
        if self.ctrl_selected_cells and len(self.ctrl_selected_cells) > 1:
            for row_id, col_id in self.ctrl_selected_cells:
                # 合計行はスキップ
                if row_id == total_row_id:
                    continue
                
                col_idx = int(col_id[1:]) - 1
                
                # 範囲チェック
                if col_idx <= 0 or col_idx >= len(all_columns):
                    continue
                
                # まとめ行の場合、収入列のみ許可
                if row_id == summary_row_id:
                    if col_idx == 3:
                        cells.append((row_id, col_id, 0, 3))
                    continue
                
                # 日付を取得
                row_vals = self.tree.item(row_id, 'values')
                try:
                    m = re.search(r'\d+', str(row_vals[0]))
                    if not m:
                        continue
                    day = int(m.group())
                except ValueError:
                    continue
                
                cells.append((row_id, col_id, day, col_idx))
            
            return cells
        
        # 範囲選択の場合（複数行が選択され、開始列が記録されている）
        if len(selected_items) > 1 and self.selection_start_col and self.selected_column_id:
            # 列の範囲を決定
            start_col_idx = int(self.selection_start_col[1:]) - 1
            end_col_idx = int(self.selected_column_id[1:]) - 1
            
            # 正規化（小さい方が先）
            if start_col_idx > end_col_idx:
                start_col_idx, end_col_idx = end_col_idx, start_col_idx
            
            # 範囲内のすべてのセルを追加
            for row_id in selected_items:
                # 合計行はスキップ
                if row_id == total_row_id:
                    continue
                
                # 日付を取得
                row_vals = self.tree.item(row_id, 'values')
                
                if row_id == summary_row_id:
                    # まとめ行は収入列(3)のみ
                    if start_col_idx <= 3 <= end_col_idx:
                        cells.append((row_id, "#4", 0, 3))
                    continue
                
                try:
                    m = re.search(r'\d+', str(row_vals[0]))
                    if not m:
                        continue
                    day = int(m.group())
                except ValueError:
                    continue
                
                # 列の範囲内のすべてのセルを追加
                for col_idx in range(start_col_idx, end_col_idx + 1):
                    # 日付列と+列をスキップ
                    if col_idx <= 0 or col_idx >= len(all_columns):
                        continue
                    
                    col_id = f"#{col_idx + 1}"
                    cells.append((row_id, col_id, day, col_idx))
        
        else:
            # 単一セルの場合
            for row_id in selected_items:
                # 合計行はスキップ
                if row_id == total_row_id:
                    continue
                
                # 列IDがない場合はスキップ
                if not self.selected_column_id:
                    continue
                
                col_idx = int(self.selected_column_id[1:]) - 1
                
                # 範囲チェック
                if col_idx <= 0 or col_idx >= len(all_columns):
                    continue
                
                # まとめ行の場合、収入列のみ許可
                if row_id == summary_row_id and col_idx != 3:
                    continue
                
                # 日付を取得
                row_vals = self.tree.item(row_id, 'values')
                if row_id == summary_row_id:
                    day = 0
                    col_idx = 3  # 収入列
                else:
                    try:
                        m = re.search(r'\d+', str(row_vals[0]))
                        if not m:
                            continue
                        day = int(m.group())
                    except ValueError:
                        continue
                
                cells.append((row_id, self.selected_column_id, day, col_idx))
        
        return cells
    
    def _copy_cells(self, event=None):
        """
        選択されたセルをコピー（確認なし）
        """
        cells = self._get_selected_cells()
        if not cells:
            return
        
        # データを収集
        copy_data = []
        for row_id, col_id, day, col_idx in cells:
            dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_idx}"
            data_list = self.data_manager.get_transaction_data(dict_key)
            
            # セルの位置情報と合わせて保存
            row_vals = self.tree.item(row_id, 'values')
            copy_data.append({
                'day': day,
                'col_idx': col_idx,
                'data': data_list if data_list else [],
                'display_value': str(row_vals[col_idx]).strip() if col_idx < len(row_vals) else ""
            })
        
        # JSON形式でクリップボードに保存
        self.root.clipboard_clear()
        if copy_data:
            json_str = json.dumps(copy_data, ensure_ascii=False)
            self.root.clipboard_append(json_str)
            self.root.update()
    
    def _cut_cells(self, event=None):
        """
        選択されたセルを切り取り（確認なし）
        """
        cells = self._get_selected_cells()
        if not cells:
            return
        
        # Undo用に操作前の状態を保存
        undo_data = []
        for row_id, col_id, day, col_idx in cells:
            dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_idx}"
            old_data = self.data_manager.get_transaction_data(dict_key)
            undo_data.append((dict_key, old_data[:] if old_data else None))
        
        self._save_undo_state('cut', undo_data)
        
        # まずコピー
        self._copy_cells()

        # 次に削除
        affected_keys = []
        for row_id, col_id, day, col_idx in cells:
            dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_idx}"
            self.data_manager.delete_transaction_data(dict_key)
            affected_keys.append(dict_key)

            # UI更新
            self.update_parent_cell(f"{self.current_year}-{self.current_month}-{day}", col_idx, "")

        # ファイルに保存
        self.data_manager.save_transactions(affected_keys)
    
    def _paste_cells(self, event=None):
        """
        クリップボードの内容をセルに貼り付け（確認なし）
        相対位置を保持したまま貼り付け
        
        貼り付け先：選択されたセルのうち、最も左上（行最小、列最小）のセルを基準とする
        """
        try:
            clipboard_text = self.root.clipboard_get()
        except tk.TclError:
            return
        
        # 貼り付け先のセルを取得
        selected_items = self.tree.selection()
        if not selected_items:
            return
        
        # 選択されたすべてのセルの中で最も左上のセルを見つける
        items = self.tree.get_children()
        summary_row_id = items[-1] if len(items) >= 1 else None
        all_columns = self.get_all_columns()
        
        base_day = None
        base_col_idx = None
        
        # Ctrl選択の場合
        if self.ctrl_selected_cells and len(self.ctrl_selected_cells) >= 1:
            min_row_idx = float('inf')
            min_col_idx = float('inf')
            selected_row_id = None
            
            for row_id, col_id in self.ctrl_selected_cells:
                try:
                    row_idx = items.index(row_id)
                    col_idx = int(col_id[1:]) - 1
                    
                    # より上（行インデックスが小さい）、または同じ行でより左（列インデックスが小さい）
                    if row_idx < min_row_idx or (row_idx == min_row_idx and col_idx < min_col_idx):
                        min_row_idx = row_idx
                        min_col_idx = col_idx
                        selected_row_id = row_id
                        base_col_idx = col_idx
                except (ValueError, IndexError):
                    continue
            
            if selected_row_id:
                row_vals = self.tree.item(selected_row_id, 'values')
                if selected_row_id == summary_row_id:
                    base_day = 0
                    base_col_idx = 3  # 収入列
                else:
                    try:
                        base_day = int(str(row_vals[0]).strip())
                    except ValueError:
                        return
        
        # 範囲選択またはその他の場合
        if base_day is None or base_col_idx is None:
            # 選択された行の中で最も上の行を見つける
            min_row_idx = float('inf')
            selected_row_id = None
            
            for row_id in selected_items:
                try:
                    row_idx = items.index(row_id)
                    if row_idx < min_row_idx:
                        min_row_idx = row_idx
                        selected_row_id = row_id
                except ValueError:
                    continue
            
            if not selected_row_id:
                return
            
            # 列の決定
            if self.selection_start_col and self.selected_column_id:
                # 範囲選択の場合、開始列と終了列の小さい方
                start_col_idx = int(self.selection_start_col[1:]) - 1
                end_col_idx = int(self.selected_column_id[1:]) - 1
                base_col_idx = min(start_col_idx, end_col_idx)
            elif self.selected_column_id:
                # 単一選択の場合
                base_col_idx = int(self.selected_column_id[1:]) - 1
            else:
                return
            
            # 日付を取得
            row_vals = self.tree.item(selected_row_id, 'values')
            if selected_row_id == summary_row_id:
                base_day = 0
                base_col_idx = 3  # 収入列
            else:
                try:
                    base_day = int(str(row_vals[0]).strip())
                except ValueError:
                    return
        
        # JSON形式のデータを解析
        try:
            paste_data = json.loads(clipboard_text)
            if not isinstance(paste_data, list):
                return
        except json.JSONDecodeError:
            # JSON形式でない場合は、単一セルとして扱う
            amount = parse_amount(clipboard_text)
            if amount != 0 or "0" in clipboard_text:
                # Undo用に操作前の状態を保存
                dict_key = f"{self.current_year}-{self.current_month}-{base_day}-{base_col_idx}"
                old_data = self.data_manager.get_transaction_data(dict_key)
                self._save_undo_state('paste', [(dict_key, old_data[:] if old_data else None)])
                
                new_data_list = [("貼付入力", str(amount), "")]
                
                # 既存データの確認（上書き）
                self.data_manager.set_transaction_data(dict_key, new_data_list)
                self.data_manager.save_transaction(dict_key)
                total = sum(parse_amount(row[1]) for row in new_data_list if len(row) > 1)
                self.update_parent_cell(f"{self.current_year}-{self.current_month}-{base_day}", base_col_idx, str(total))
            return
        
        # 複数セルの貼り付け：各セルの相対位置を保持
        if paste_data:
            # データ形式を判定
            is_detail_window_data = False
            if paste_data and isinstance(paste_data[0], list):
                # 詳細入力ウィンドウからのデータ形式: [["支払先", "金額", "メモ"], ...]
                is_detail_window_data = True
            
            if is_detail_window_data:
                # 詳細入力ウィンドウからのデータを現在のセルに貼り付け
                dict_key = f"{self.current_year}-{self.current_month}-{base_day}-{base_col_idx}"
                old_data = self.data_manager.get_transaction_data(dict_key)
                
                # Undo用に元のデータを保存
                self._save_undo_state('paste', [(dict_key, old_data[:] if old_data else None)])
                
                # 詳細データとして保存
                new_data_list = []
                for row in paste_data:
                    if isinstance(row, list) and len(row) >= 2:
                        # [支払先, 金額, メモ] の形式
                        safe_row = [str(v) for v in row]
                        while len(safe_row) < 3:
                            safe_row.append("")
                        new_data_list.append(tuple(safe_row[:3]))
                
                if new_data_list:
                    self.data_manager.set_transaction_data(dict_key, new_data_list)
                    self.data_manager.save_transaction(dict_key)
                    total = sum(parse_amount(row[1]) for row in new_data_list if len(row) > 1)
                    self.update_parent_cell(f"{self.current_year}-{self.current_month}-{base_day}", base_col_idx, str(total))
            else:
                # メインウィンドウからのデータ形式: [{"day": 1, "col_idx": 3, "data": [...]}, ...]
                # コピー元の最小の日と列を見つける（基準点）
                min_day = min(cell['day'] for cell in paste_data)
                min_col = min(cell['col_idx'] for cell in paste_data)
                
                days_in_month = self.get_days_in_month()
                
                # Undo用に影響を受けるすべてのセルの元データを保存
                undo_data = []
                
                # 各セルを貼り付け
                for cell_data in paste_data:
                    # 元のセルの基準点からの相対位置を計算
                    day_offset = cell_data['day'] - min_day
                    col_offset = cell_data['col_idx'] - min_col
                    
                    # 貼り付け先の位置を計算
                    target_day = base_day + day_offset
                    target_col_idx = base_col_idx + col_offset
                    
                    # 範囲チェック
                    if target_day < 0 or (target_day > days_in_month and target_day != 0):
                        continue
                    if target_col_idx <= 0 or target_col_idx >= len(all_columns):
                        continue
                    
                    # Undo用に元のデータを保存
                    dict_key = f"{self.current_year}-{self.current_month}-{target_day}-{target_col_idx}"
                    old_data = self.data_manager.get_transaction_data(dict_key)
                    undo_data.append((dict_key, old_data[:] if old_data else None))
                    
                    # データを貼り付け
                    new_data = cell_data.get('data', [])
                    
                    if new_data:
                        self.data_manager.set_transaction_data(dict_key, new_data)
                        total = sum(parse_amount(row[1]) for row in new_data if len(row) > 1)
                        self.update_parent_cell(f"{self.current_year}-{self.current_month}-{target_day}", target_col_idx, str(total))
                
                # Undo履歴に保存・ファイルに保存
                if undo_data:
                    self._save_undo_state('paste', undo_data)
                    self.data_manager.save_transactions([key for key, _ in undo_data])
    
    def _delete_cells(self, event=None):
        """
        選択されたセルを削除（確認なし）
        """
        cells = self._get_selected_cells()
        if not cells:
            return
        
        # Undo用に操作前の状態を保存
        undo_data = []
        for row_id, col_id, day, col_idx in cells:
            dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_idx}"
            old_data = self.data_manager.get_transaction_data(dict_key)
            undo_data.append((dict_key, old_data[:] if old_data else None))
        
        self._save_undo_state('delete', undo_data)

        affected_keys = []
        for row_id, col_id, day, col_idx in cells:
            dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_idx}"
            self.data_manager.delete_transaction_data(dict_key)
            affected_keys.append(dict_key)

            # UI更新
            self.update_parent_cell(f"{self.current_year}-{self.current_month}-{day}", col_idx, "")

        # ファイルに保存
        self.data_manager.save_transactions(affected_keys)

    def _save_undo_state(self, action_type, cells_data):
        """
        操作前の状態をundo stackに保存
        
        Args:
            action_type: 'cut', 'paste', 'delete'のいずれか
            cells_data: [(dict_key, old_data), ...] の形式
        """
        undo_entry = {
            'action': action_type,
            'cells': cells_data,
            'year': self.current_year,
            'month': self.current_month
        }
        
        self.undo_stack.append(undo_entry)
        
        # 最大保持数を超えたら古いものから削除
        if len(self.undo_stack) > self.max_undo_count:
            self.undo_stack.pop(0)
    
    def _undo(self, event=None):
        """
        最後の操作を元に戻す (Ctrl+Z)
        """
        if not self.undo_stack:
            return
        
        undo_entry = self.undo_stack.pop()
        action = undo_entry['action']
        cells = undo_entry['cells']
        year = undo_entry['year']
        month = undo_entry['month']
        
        # 年月が異なる場合は表示を切り替え
        if self.current_year != year or self.current_month != month:
            self.current_year = year
            self.current_month = month
            self.update_year_display()
            self._update_month_buttons()
            self._show_month(self.current_month)
        
        if action == 'cut' or action == 'delete':
            # 切り取り/削除の取り消し：データを復元
            for dict_key, old_data in cells:
                if old_data:
                    self.data_manager.set_transaction_data(dict_key, old_data)
                    # UI更新
                    parts = dict_key.split('-')
                    if len(parts) == 4:
                        y, m, d, col_idx = int(parts[0]), int(parts[1]), int(parts[2]), int(parts[3])
                        total = sum(parse_amount(row[1]) for row in old_data if len(row) > 1)
                        self.update_parent_cell(f"{y}-{m}-{d}", col_idx, str(total))
        
        elif action == 'paste' or action == 'edit_detail':
            # 貼り付け/詳細編集の取り消し：貼り付けたデータを削除し、元のデータを復元
            for dict_key, old_data in cells:
                if old_data is None:
                    # 元々データがなかった場合は削除
                    self.data_manager.delete_transaction_data(dict_key)
                    parts = dict_key.split('-')
                    if len(parts) == 4:
                        y, m, d, col_idx = int(parts[0]), int(parts[1]), int(parts[2]), int(parts[3])
                        self.update_parent_cell(f"{y}-{m}-{d}", col_idx, "")
                else:
                    # 元のデータがあった場合は復元
                    self.data_manager.set_transaction_data(dict_key, old_data)
                    parts = dict_key.split('-')
                    if len(parts) == 4:
                        y, m, d, col_idx = int(parts[0]), int(parts[1]), int(parts[2]), int(parts[3])
                        total = sum(parse_amount(row[1]) for row in old_data if len(row) > 1)
                        self.update_parent_cell(f"{y}-{m}-{d}", col_idx, str(total))