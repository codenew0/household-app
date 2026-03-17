# ui/search_dialog.py
"""
取引データを検索するダイアログ
"""
import tkinter as tk
from tkinter import ttk, messagebox
from ui.base_dialog import BaseDialog
from config import parse_amount


class SearchDialog(BaseDialog):
    """
    取引データを検索するためのダイアログ。
    
    全期間のデータから、支払先名、金額、詳細のいずれかに
    検索文字列が含まれる取引を抽出して表示する。
    大文字小文字を区別しない部分一致検索を行う。
    """
    
    def __init__(self, parent, parent_app):
        """
        検索ダイアログを初期化する。
        
        Args:
            parent: 親ウィンドウ
            parent_app: メインアプリケーションのインスタンス
        """
        self.parent_app = parent_app
        self.search_results = []
        
        super().__init__(parent, "検索")
        
        self._create_widgets()
    
    def _create_widgets(self):
        """ダイアログ内のUI要素を作成する"""
        # グリッドレイアウトの設定
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # 検索入力セクション
        search_frame = tk.Frame(self, bg='#f0f0f0')
        search_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        search_frame.grid_columnconfigure(1, weight=1)
        
        # 検索ラベル
        tk.Label(search_frame, text="検索文字列:", font=('Arial', 12), bg='#f0f0f0').grid(
            row=0, column=0, padx=(0, 10), sticky="w")
        
        # 検索入力フィールド
        self.search_entry = tk.Entry(search_frame, font=('Arial', 12), width=30)
        self.search_entry.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        self.search_entry.focus_set()
        
        # 検索ボタン
        search_btn = tk.Button(search_frame, text="検索", font=('Arial', 12),
                               bg='#2196f3', fg='white', relief='raised', bd=2,
                               activebackground='#1976d2', command=self._search)
        search_btn.grid(row=0, column=2, padx=(0, 10), ipady=3)
        
        # クリアボタン
        clear_btn = tk.Button(search_frame, text="クリア", font=('Arial', 12),
                              bg='#ff9800', fg='white', relief='raised', bd=2,
                              activebackground='#f57c00', command=self._clear_results)
        clear_btn.grid(row=0, column=3, ipady=3)
        
        # 結果表示セクション
        result_frame = tk.Frame(self, bg='#f0f0f0')
        result_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        result_frame.grid_rowconfigure(0, weight=1)
        result_frame.grid_columnconfigure(0, weight=1)
        
        # 結果表示用Treeview
        columns = ["年月日", "項目", "支払先", "金額(円)", "メモ"]
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=15)
        
        # 列の設定
        self.default_column_widths = {}
        widths = {
            "年月日": 100,
            "項目": 120,
            "支払先": 150,
            "金額(円)": 100,
            "メモ": 200
        }

        self.sort_column = None
        self.sort_reverse = False
        
        for col in columns:
            self.result_tree.heading(col, text=col, command=lambda c=col: self._sort_by_column(c))
            width = widths.get(col, 100)
            self.result_tree.column(col, anchor="center", width=width, minwidth=int(width * 0.8))
            self.default_column_widths[col] = width
        
        self.result_tree.grid(row=0, column=0, sticky="nsew")
        
        # スクロールバー
        v_scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        self.result_tree.configure(yscrollcommand=v_scrollbar.set)
        
        h_scrollbar = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        self.result_tree.configure(xscrollcommand=h_scrollbar.set)

        # 統計情報表示エリア
        stats_frame = tk.Frame(self, bg='#f0f0f0')
        stats_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=5)
        
        self.stats_label = tk.Label(stats_frame, text="", font=('Arial', 12, 'bold'), bg='#f0f0f0')
        self.stats_label.pack()
        
        # 結果カウンター
        self.result_label = tk.Label(self, text="検索結果: 0 件",
                                     font=('Arial', 10), bg='#f0f0f0', fg='#666666')
        self.result_label.grid(row=3, column=0, sticky="w", padx=10, pady=(5, 10))
        
        # 閉じるボタン
        close_btn = tk.Button(self, text="閉じる", font=('Arial', 12),
                              bg='#f44336', fg='white', relief='raised', bd=2,
                              activebackground='#d32f2f', command=self.destroy)
        close_btn.grid(row=4, column=0, pady=(0, 10), ipady=5)
        
        # キーボードショートカット
        self.search_entry.bind('<Return>', lambda e: self._search())
        self.bind('<Escape>', lambda e: self.destroy())
        self.bind('<Control-f>', lambda e: self.search_entry.focus_set())
        self.result_tree.bind("<MouseWheel>", self._on_mousewheel)
        self.result_tree.bind("<Double-1>", self._on_double_click)
        self.result_tree.bind("<Button-3>", self._on_header_right_click)

    def _sort_by_column(self, column):
        """
        指定された列でデータをソートする
        
        Args:
            column: ソート対象の列名
        """
        if self.sort_column == column:
            # 同じ列をクリックした場合は昇順/降順を切り替え
            self.sort_reverse = not self.sort_reverse
        else:
            # 新しい列の場合は昇順でソート
            self.sort_column = column
            self.sort_reverse = False
        
        # ソートキーのマッピング
        sort_key_map = {
            "年月日": lambda x: (x['year'], x['month'], x['day']),
            "項目": lambda x: x['column'],
            "支払先": lambda x: x['partner'],
            "金額(円)": lambda x: parse_amount(x['amount']),
            "メモ": lambda x: x['detail']
        }
        
        # データをソート
        self.search_results.sort(key=sort_key_map.get(column, lambda x: ""),
                                reverse=self.sort_reverse)
        
        # Treeviewを更新
        self._refresh_treeview()
        
        # 列ヘッダーを更新（ソート方向を表示）
        self._update_column_headers()
    
    def _refresh_treeview(self):
        """ソート後のデータでTreeviewを再表示する"""
        # 既存のアイテムをクリア
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        
        # ソート済みデータを再表示
        for result in self.search_results:
            values = [result['date'], result['column'], result['partner'],
                      result['amount'], result['detail']]
            self.result_tree.insert("", "end", values=values)
    
    def _update_column_headers(self):
        """ソート状態を示すため、列ヘッダーに矢印を表示する"""
        columns = ["年月日", "項目", "支払先", "金額(円)", "メモ"]
        
        for col in columns:
            if col == self.sort_column:
                # ソート中の列には矢印を表示
                arrow = " ▼" if self.sort_reverse else " ▲"
                self.result_tree.heading(col, text=f"{col}{arrow}")
            else:
                # その他の列は通常表示
                self.result_tree.heading(col, text=col)
    
    def _on_header_right_click(self, event):
        """ヘッダーの右クリックイベントを処理する"""
        region = self.result_tree.identify_region(event.x, event.y)
        if region != "heading":
            return
        
        col_id = self.result_tree.identify_column(event.x)
        if not col_id:
            return
        
        context_menu = tk.Menu(self, tearoff=0)
        context_menu.add_command(label="全ての列幅をリセット",
                                 command=self._reset_all_column_widths)
        context_menu.post(event.x_root, event.y_root)
    
    def _reset_all_column_widths(self):
        """全ての列の幅をデフォルトにリセットする"""
        columns = ["年月日", "項目", "支払先", "金額(円)", "メモ"]
        
        for i, col_name in enumerate(columns):
            col_id = f"#{i + 1}"
            if col_name in self.default_column_widths:
                self.result_tree.column(col_id, width=self.default_column_widths[col_name])
    
    def _on_mousewheel(self, event):
        """マウスホイールによるスクロール処理"""
        self.result_tree.yview_scroll(int(-1 * (event.delta / 120)), "units")
    
    def _on_double_click(self, event):
        """検索結果の行をダブルクリックした時の処理"""
        item = self.result_tree.selection()
        if not item:
            return
        
        try:
            all_items = self.result_tree.get_children()
            row_index = all_items.index(item[0])
        except (ValueError, IndexError):
            return
        
        if row_index >= len(self.search_results):
            return
        
        data = self.search_results[row_index]
        year = data['year']
        month = data['month']
        day = data['day']
        col_index = data['col_index']
        
        # 親アプリケーションの年月を変更（月が変わる場合はツリーが再描画される）
        need_redraw = (self.parent_app.current_year != year or
                       self.parent_app.current_month != month)
        if need_redraw:
            self.parent_app.current_year = year
            self.parent_app.current_month = month
            self.parent_app.update_year_display()
            self.parent_app.select_month(month)
        
        # 再描画後にセル選択（月切り替え時は再描画完了を待って遅延実行）
        if need_redraw:
            self.after(50, lambda: self.parent_app.navigate_to_cell(day, col_index))
        else:
            self.parent_app.navigate_to_cell(day, col_index)
    
    def _search(self):
        """検索を実行する"""
        search_text = self.search_entry.get().strip()
        if not search_text:
            messagebox.showwarning("警告", "検索文字列を入力してください。")
            return
        
        # 前回の検索結果をクリア
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        
        self.search_results = []
        
        # データマネージャーから検索
        results = self.parent_app.data_manager.search_transactions(search_text)
        total_amount = 0  # 合計金額用（ループ内で計算して効率化）
        
        # 結果を整形
        all_columns = self.parent_app.get_all_columns()
        for result in results:
            year = result['year']
            month = result['month']
            day = result['day']
            col_index = result['col_index']
            date_str = f"{year}/{month:02d}/{day:02d}"
            
            if day == 0:
                column_name = "収入"
                date_str = f"{year}/{month:02d}"
            elif col_index < len(all_columns):
                column_name = all_columns[col_index]
            else:
                column_name = f"列{col_index}"

            # 金額の加算（統計計算をループ内で実行して二重ループを回避・効率化）
            amount_val = parse_amount(result['amount'])
            total_amount += amount_val
            
            search_result = {
                'year': year,
                'month': month,
                'day': day,
                'col_index': col_index,
                'date': date_str,
                'column': column_name,
                'partner': result['partner'],
                'amount': result['amount'],
                'detail': result['detail']
            }
            self.search_results.append(search_result)
        
        # 結果を日付順にソート（デフォルト）
        self.sort_column = "年月日"
        self.sort_reverse = False
        self.search_results.sort(key=lambda x: (x['year'], x['month'], x['day'], x['col_index']))
        
        # 結果を表示
        for result in self.search_results:
            values = [result['date'], result['column'], result['partner'],
                      result['amount'], result['detail']]
            self.result_tree.insert("", "end", values=values)
        
        # 列ヘッダーを更新
        self._update_column_headers()
        
        # 結果カウンターを更新
        count = len(self.search_results)
        self.result_label.config(text=f"検索結果: {len(self.search_results)} 件")

        # 統計情報の更新
        if count > 0:
            avg_amount = total_amount / count
            self.stats_label.config(text=f"合計金額: ¥{total_amount:,} | 平均金額: ¥{avg_amount:.0f}")
        else:
            self.stats_label.config(text="")
    
    def _clear_results(self):
        """検索結果と入力フィールドをクリアする"""
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        
        self.search_results = []
        self.result_label.config(text="検索結果: 0 件")
        self.stats_label.config(text="")
        self.search_entry.delete(0, tk.END)
        self.search_entry.focus_set()