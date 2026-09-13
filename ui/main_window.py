# ui/main_window.py
"""
メインウィンドウの実装
家計管理アプリケーションのメインUI
"""
import tkinter as tk
from models.transactions import cash_amount, is_pending
from models.clipboard import (
    decode_clipboard, PLAIN_AMOUNT, DETAIL_ROWS, CELL_BLOCK,
)
import json
from tkinter import ttk, messagebox
import tkinter.font as tkfont
from config import (
    WindowConfig, ColorTheme, TreeviewConfig, DefaultColumns,
    FontConfig, DialogConfig, APP_TITLE,
    get_current_year, get_current_month
)
from models.data_manager import DataManager
from ui.tooltip import TreeviewTooltip
from ui.transaction_dialog import TransactionDialog
from ui.search_dialog import SearchDialog
from ui.chart_dialog import ChartDialog
from ui.csv_dialog import CsvDialog
from ui.subscription_dialog import SubscriptionDialog
from ui.receipt_dialog import ReceiptDialog
from ui.base_dialog import BaseDialog
from ui.productivity_dialogs import LedgerListDialog, BudgetDialog, RecoveryDialog
from ui.month_picker import MonthPicker
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
        
        self.current_month_button = None
        
        # 初期化
        self._setup_window()
        self._load_data()
        self._create_ui()
        self._show_month(self.current_month)
        self._check_subscriptions()
        
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

    def _check_subscriptions(self):
        try:
            # 編集中のダイアログと自動生成が同じセルを更新しないようにする。
            if self.root.grab_current() is None:
                if self.data_manager.generate_due_subscriptions():
                    self._show_month(self.current_month)
        except (OSError, ValueError) as error:
            messagebox.showerror('定期支払いの自動登録エラー', str(error), parent=self.root)
        finally:
            self.root.after(60000, self._check_subscriptions)
    
    def _setup_window(self):
        """メインウィンドウの基本設定を行う"""
        self.root.title(APP_TITLE)
        
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
        if self.data_manager.load_warnings:
            messagebox.showwarning(
                'データ読み込みの警告', '\n\n'.join(self.data_manager.load_warnings),
                parent=self.root)
    
    def _save_data(self):
        """データと設定を保存する"""
        self.data_manager.save_data()
        self.data_manager.save_settings()
    
    def _on_closing(self):
        """ウィンドウが閉じられる時の処理"""
        try:
            self._save_data()
            self.data_manager.save_backup()
        except OSError as error:
            messagebox.showerror('保存エラー', f'終了前の保存に失敗しました。\n{error}', parent=self.root)
            return
        self.root.destroy()
    
    def _create_ui(self):
        """メインウィンドウのUI要素を作成する"""
        self._create_menu()
        # メインコンテナ
        main_container = tk.Frame(self.root, bg=self.colors['bg_primary'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        # ヘッダーセクション
        header = tk.Frame(main_container, bg=self.colors['bg_secondary'])
        header.pack(fill=tk.X, pady=(0, 8))
        
        header_inner = tk.Frame(header, bg=self.colors['bg_secondary'])
        header_inner.pack(fill=tk.X, padx=15, pady=8)
        
        self._create_year_controls(header_inner)
        self._create_month_buttons(header_inner)
        self.month_picker = MonthPicker(header_inner, self.current_year, self.current_month, self._select_year_month, anchor=self.year_label, theme='main')
        
        # 検索ボタン
        self._create_search_button(header_inner)
        
        # 図表ボタン
        self._create_chart_button(header_inner)

        # CSV入出力ボタン
        self._create_csv_button(header_inner)
        
        # 現在月表示(右側、クリック可能)
        self._create_current_month_button(header_inner)

        # メインテーブルセクション
        tree_section = tk.Frame(main_container, bg=self.colors['bg_secondary'])
        tree_section.pack(fill=tk.BOTH, expand=True)
        
        self._create_treeview(tree_section)
        self._update_month_buttons()

    def _create_menu(self):
        """機能を用途別にまとめたメニューバーを作成する。"""
        self.menu_bar = tk.Menu(self.root)
        sections = (
            ('データ', (
                ('CSVインポート・エクスポート…', CsvDialog, ''),
                ('バックアップから復元…', RecoveryDialog, ''),
            )),
            ('入力・管理', (
                ('未確定の一括確認…', lambda parent, app: LedgerListDialog(parent, app, pending=True), ''),
                ('レシートOCR…', ReceiptDialog, ''),
                ('定期支払い…', SubscriptionDialog, ''),
            )),
            ('検索・分析', (
                ('検索…', SearchDialog, 'Ctrl+F'),
                ('明細一覧…', LedgerListDialog, ''),
                ('予算・残額・前月比較…', BudgetDialog, ''),
                ('図表…', ChartDialog, ''),
            )),
        )
        for label, entries in sections:
            menu = tk.Menu(self.menu_bar, tearoff=False)
            for title, dialog, accelerator in entries:
                menu.add_command(
                    label=title, accelerator=accelerator,
                    command=lambda dialog=dialog: dialog(self.root, self),
                )
            self.menu_bar.add_cascade(label=label, menu=menu)
        self.root.configure(menu=self.menu_bar)
    
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
    
    def _open_year_input_dialog(self, event=None):
        self.month_picker.set(self.current_year, self.current_month)
        self.month_picker.open()

    def _prev_month(self):
        self.month_picker.move(-1)

    def _next_month(self):
        self.month_picker.move(1)

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

    def _create_csv_button(self, parent):
        """CSVインポート・エクスポートボタンを作成する。"""
        csv_container = tk.Frame(parent, bg=self.colors['bg_secondary'])
        csv_container.pack(side=tk.LEFT, padx=(10, 0))
        ttk.Button(csv_container, text="CSV", width=7, style='Accent.TButton',
                   command=lambda: CsvDialog(self.root, self)).pack()
    
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
        self.tree.tag_configure('pending', foreground='red')
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
    
    def update_year_display(self):
        """年表示を更新"""
        self.month_picker.set(self.current_year, self.current_month)
        self.year_label.configure(text=str(self.current_year))

    def _select_year_month(self, year, month):
        self.current_year = year
        self.select_month(month)
    
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
        """年月選択の表示を更新する。"""
        self.update_year_display()
        for month, button in enumerate(self.month_buttons, 1):
            button.configure(style="Selected.TButton" if month == self.current_month else "Modern.TButton")
        if self.current_month_button:
            self.current_month_button.configure(text=f"📅 {self.current_month:02d}月")
    
    def _open_monthly_data(self):
        """選択中の月に絞った統合明細一覧を開く。"""
        LedgerListDialog(self.root, self)
    
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
        self.tree.tag_configure('pending', foreground='red')
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

            pending = any(is_pending(row) for col in range(1, len(all_columns))
                          for row in self.data_manager.get_transaction_data(
                              f'{self.current_year}-{self.current_month}-{day}-{col}'))
            self.tree.insert("", "end", values=formatted_values, tags=('pending', tag) if pending else (tag,))
        
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
                total = sum(cash_amount(row) for row in data_list if len(row) > 1)
                if total != 0:
                    totals[col_index] = str(total)
                elif any(is_pending(row) for row in data_list):
                    totals[col_index] = '未確定'
        
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
        income_column = DefaultColumns.INCOME_COLUMN_INDEX
        dict_key = f"{self.current_year}-{self.current_month}-0-{income_column}"
        data_list = self.data_manager.get_transaction_data(dict_key)
        if data_list:
            return sum(cash_amount(row) for row in data_list if len(row) > 1)
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
            income_index = DefaultColumns.INCOME_COLUMN_INDEX
            income_str = (str(summary_vals[income_index]).strip()
                          if len(summary_vals) > income_index else "")
            income_val = int(income_str) if income_str else 0
        except:
            income_val = 0
        
        # 収支差額と総支出を更新
        balance = income_val - grand_total
        summary_vals[1] = f" {balance} " if balance != 0 else "  "
        summary_vals[DefaultColumns.EXPENSE_COLUMN_INDEX] = (
            f" {grand_total} " if grand_total != 0 else "  ")
        
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
        col_id = self.tree.identify_column(event.x)
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            col_index = self._column_index(col_id)
            if col_index is None:
                return
            all_columns = self.get_all_columns()
            if col_index == len(all_columns):  # +ボタン
                self._add_column()
            elif col_index >= len(DefaultColumns.ITEMS):  # カスタム列
                self._edit_column_name(col_index)
            return

        row_id = self.tree.identify_row(event.y)
        self._open_transaction_dialog(row_id, col_id, wait=True)

    def _transaction_target(self, row_id, col_id):
        """編集可能なセルから（データキー、表示分類名）を取得する。"""
        if not row_id or not col_id:
            return None
        items = self.tree.get_children()
        if len(items) < 2 or row_id == items[-2]:
            return None
        column = self._column_index(col_id)
        if column is None or not 0 < column < len(self.get_all_columns()):
            return None
        if row_id == items[-1]:
            if column != DefaultColumns.INCOME_COLUMN_INDEX:
                return None
            day = 0
            column_name = "収入"
        else:
            day = self._row_day(row_id)
            if day is None:
                return None
            column_name = self.tree.heading(col_id, "text")
        key = f"{self.current_year}-{self.current_month}-{day}-{column}"
        return key, column_name

    def _open_transaction_dialog(self, row_id, col_id, wait=False):
        target = self._transaction_target(row_id, col_id)
        if target is None:
            return
        dialog = TransactionDialog(self.root, self, *target)
        if wait:
            self.root.wait_window(dialog)
    
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
        """選択中の編集可能セルをSpaceキーで開く。"""
        selected_items = self.tree.selection()
        if selected_items and self.tree.focus():
            self._open_transaction_dialog(selected_items[0], self.selected_column_id)
    
    def _reset_all_column_widths(self):
        """指定された列の幅をデフォルトにリセットする"""
        all_columns = self.get_all_columns() + ["+"]
        for i, col_name in enumerate(all_columns):
            col_id = f"#{i + 1}"
            if col_name in self.default_column_widths:
                self.tree.column(col_id, width=self.default_column_widths[col_name])

    def _add_column(self):
        """新しい列を追加する"""
        dialog = BaseDialog(
            self.root,
            "列の追加",
            DialogConfig.COLUMN_EDIT_WIDTH,
            DialogConfig.COLUMN_EDIT_HEIGHT,
        )
        dialog.resizable(False, False)
        
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
                    try:
                        self.data_manager.add_custom_column(column_name)
                    except OSError as error:
                        messagebox.showerror('列の追加エラー', str(error), parent=dialog)
                        return
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
        dialog.show_ready(entry)

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
        dialog = BaseDialog(
            self.root,
            "列名の編集",
            DialogConfig.COLUMN_EDIT_WIDTH,
            DialogConfig.COLUMN_EDIT_HEIGHT,
        )
        dialog.resizable(False, False)
        
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
                    try:
                        self.data_manager.edit_custom_column(old_name, new_name)
                    except OSError as error:
                        messagebox.showerror('列名の保存エラー', str(error), parent=dialog)
                        return
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
        dialog.show_ready(entry)

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
            try:
                self.data_manager.remove_custom_column(col_name)
            except OSError as error:
                messagebox.showerror('分類の削除エラー', str(error), parent=self.root)
                return
            # 列番号を保持したUndoやセル選択は削除後には安全に適用できない。
            self.undo_stack.clear()
            self.ctrl_selected_cells.clear()
            self.selection_start_row = self.selection_start_col = None
            self.selected_column_id = None
            
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
        """現在の単一・矩形・Ctrl選択を共通のセル情報へ変換する。"""
        selected_items = self.tree.selection()
        if not selected_items:
            return []

        items = self.tree.get_children()
        total_row_id = items[-2] if len(items) >= 2 else None
        summary_row_id = items[-1] if len(items) >= 1 else None
        column_count = len(self.get_all_columns())

        if self.ctrl_selected_cells and len(self.ctrl_selected_cells) > 1:
            candidates = list(self.ctrl_selected_cells)
        elif (len(selected_items) > 1 and self.selection_start_col
              and self.selected_column_id):
            start = self._column_index(self.selection_start_col)
            end = self._column_index(self.selected_column_id)
            if start is None or end is None:
                return []
            candidates = [
                (row_id, f"#{column + 1}")
                for row_id in selected_items
                for column in range(min(start, end), max(start, end) + 1)
            ]
        else:
            if not self.selected_column_id:
                return []
            candidates = [(row_id, self.selected_column_id) for row_id in selected_items]

        cells = []
        for row_id, col_id in candidates:
            column = self._column_index(col_id)
            if row_id == total_row_id or column is None or not 0 < column < column_count:
                continue
            if row_id == summary_row_id:
                if column == DefaultColumns.INCOME_COLUMN_INDEX:
                    cells.append((row_id, "#4", 0, DefaultColumns.INCOME_COLUMN_INDEX))
                continue
            day = self._row_day(row_id)
            if day is not None:
                cells.append((row_id, col_id, day, column))
        return cells

    @staticmethod
    def _column_index(column_id):
        """Treeviewの列ID（#1など）を0始まりの列番号へ変換する。"""
        try:
            return int(column_id[1:]) - 1
        except (TypeError, ValueError, IndexError):
            return None

    def _row_day(self, row_id):
        """日付行から日を取得する。合計などの日付行以外はNoneを返す。"""
        values = self.tree.item(row_id, 'values')
        match = re.search(r'\d+', str(values[0])) if values else None
        return int(match.group()) if match else None
    
    def _copy_cells(self, event=None):
        """
        選択されたセルをコピー（確認なし）
        """
        cells = self._get_selected_cells()
        if not cells:
            return
        
        # データを収集
        copy_data = []
        for _, _, day, col_idx in cells:
            dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_idx}"
            data_list = self.data_manager.get_transaction_data(dict_key)
            
            # セルの位置情報と合わせて保存
            copy_data.append({
                'day': day,
                'col_idx': col_idx,
                'data': data_list if data_list else [],
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
        for _, _, day, col_idx in cells:
            dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_idx}"
            old_data = self.data_manager.get_transaction_data(dict_key)
            undo_data.append((dict_key, old_data[:] if old_data else None))
        
        self._save_undo_state('cut', undo_data)
        
        # まずコピー
        self._copy_cells()

        # 次に削除
        affected_keys = []
        for _, _, day, col_idx in cells:
            dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_idx}"
            self.data_manager.delete_transaction_data(dict_key)
            affected_keys.append(dict_key)

            # UI更新
            self.update_parent_cell(f"{self.current_year}-{self.current_month}-{day}", col_idx, "")

        # ファイルに保存
        try:
            self.data_manager.save_transactions(affected_keys)
        except OSError as error:
            for dict_key, old_data in undo_data:
                self.data_manager.set_transaction_data(dict_key, old_data or [])
            self.undo_stack.pop()
            self._show_month(self.current_month)
            messagebox.showerror('切り取りエラー', str(error), parent=self.root)

    def _replace_cell_transactions(self, day, column, rows):
        """1セルの明細を置換し、保存失敗時はメモリとUndo履歴を戻す。"""
        key = f"{self.current_year}-{self.current_month}-{day}-{column}"
        previous = self.data_manager.get_transaction_data(key)
        undo_value = previous[:] if previous else None
        self._save_undo_state('paste', [(key, undo_value)])
        self.data_manager.set_transaction_data(key, rows)
        try:
            self.data_manager.save_transaction(key)
        except OSError as error:
            self.data_manager.set_transaction_data(key, previous or [])
            self.undo_stack.pop()
            messagebox.showerror('貼り付けエラー', str(error), parent=self.root)
            return False
        total = sum(cash_amount(row) for row in rows)
        self.update_parent_cell(
            f"{self.current_year}-{self.current_month}-{day}",
            column,
            str(total) if total else "",
        )
        return True

    def _paste_cell_block(self, base_day, base_column, cells):
        """コピー元の相対位置を保って複数セルを貼り付ける。"""
        first_day = min(cell['day'] for cell in cells)
        first_column = min(cell['col_idx'] for cell in cells)
        column_count = len(self.get_all_columns())
        days_in_month = self.get_days_in_month()
        undo_data = []

        for cell in cells:
            target_day = base_day + cell['day'] - first_day
            target_column = base_column + cell['col_idx'] - first_column
            if not 0 <= target_day <= days_in_month:
                continue
            if not 0 < target_column < column_count:
                continue
            if (target_day == 0 and
                    target_column != DefaultColumns.INCOME_COLUMN_INDEX):
                continue

            key = f"{self.current_year}-{self.current_month}-{target_day}-{target_column}"
            previous = self.data_manager.get_transaction_data(key)
            undo_data.append((key, previous[:] if previous else None))
            rows = cell['data']
            self.data_manager.set_transaction_data(key, rows)
            total = sum(cash_amount(row) for row in rows)
            self.update_parent_cell(
                f"{self.current_year}-{self.current_month}-{target_day}",
                target_column,
                str(total) if total else "",
            )

        if not undo_data:
            return
        self._save_undo_state('paste', undo_data)
        try:
            self.data_manager.save_transactions([key for key, _ in undo_data])
        except OSError as error:
            for key, previous in undo_data:
                self.data_manager.set_transaction_data(key, previous or [])
            self.undo_stack.pop()
            try:
                self.data_manager.save_transactions(
                    [key for key, _ in undo_data], create_snapshot=False)
            except OSError:
                pass
            self._show_month(self.current_month)
            messagebox.showerror('貼り付けエラー', str(error), parent=self.root)
    
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
        
        selected_cells = self._get_selected_cells()
        if not selected_cells:
            if self._column_index(self.selected_column_id) == 0:
                messagebox.showwarning(
                    '貼り付けエラー', '日付列には金額を貼り付けできません。',
                    parent=self.root,
                )
            return
        item_order = {item: index for index, item in enumerate(self.tree.get_children())}
        _, _, base_day, base_col_idx = min(
            selected_cells,
            key=lambda cell: (item_order.get(cell[0], len(item_order)), cell[3]),
        )

        try:
            paste_kind, paste_data = decode_clipboard(clipboard_text)
        except (TypeError, ValueError) as error:
            messagebox.showwarning('貼り付けエラー', str(error), parent=self.root)
            return
        if paste_kind == PLAIN_AMOUNT:
            self._replace_cell_transactions(base_day, base_col_idx, paste_data)
            return

        if paste_kind == DETAIL_ROWS:
            self._replace_cell_transactions(base_day, base_col_idx, paste_data)
        elif paste_kind == CELL_BLOCK:
            self._paste_cell_block(base_day, base_col_idx, paste_data)
    
    def _delete_cells(self, event=None):
        """
        選択されたセルを削除（確認なし）
        """
        cells = self._get_selected_cells()
        if not cells:
            return
        
        # Undo用に操作前の状態を保存
        undo_data = []
        for _, _, day, col_idx in cells:
            dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_idx}"
            old_data = self.data_manager.get_transaction_data(dict_key)
            undo_data.append((dict_key, old_data[:] if old_data else None))
        
        self._save_undo_state('delete', undo_data)

        affected_keys = []
        for _, _, day, col_idx in cells:
            dict_key = f"{self.current_year}-{self.current_month}-{day}-{col_idx}"
            self.data_manager.delete_transaction_data(dict_key)
            affected_keys.append(dict_key)

            # UI更新
            self.update_parent_cell(f"{self.current_year}-{self.current_month}-{day}", col_idx, "")

        # ファイルに保存
        try:
            self.data_manager.save_transactions(affected_keys)
        except OSError as error:
            for dict_key, old_data in undo_data:
                self.data_manager.set_transaction_data(dict_key, old_data or [])
            self.undo_stack.pop()
            self._show_month(self.current_month)
            messagebox.showerror('削除エラー', str(error), parent=self.root)

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
                        total = sum(cash_amount(row) for row in old_data if len(row) > 1)
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
                        total = sum(cash_amount(row) for row in old_data if len(row) > 1)
                        self.update_parent_cell(f"{y}-{m}-{d}", col_idx, str(total))

        # Undo後の状態も即時保存し、強制終了時の巻き戻りを防ぐ。
        self.data_manager.save_transactions([dict_key for dict_key, _ in cells])
        self._show_month(self.current_month)
