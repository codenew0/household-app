# ui/monthly_data_dialog.py
"""
月間データの詳細を表示するダイアログ
"""
import tkinter as tk
from tkinter import ttk
from ui.base_dialog import BaseDialog
from config import parse_amount


class MonthlyDataDialog(BaseDialog):
    """
    月間データの詳細を表示するダイアログ。
    
    指定された年月のすべての取引データを一覧表示し、
    項目ごとにソート可能な機能を提供する。
    合計金額や平均金額などの統計情報も表示する。
    """
    
    def __init__(self, parent, parent_app, year, month):
        """
        月間データダイアログを初期化する。
        
        Args:
            parent: 親ウィンドウ
            parent_app: メインアプリケーションのインスタンス
            year: 表示する年
            month: 表示する月
        """
        self.parent_app = parent_app
        self.year = year
        self.month = month
        self.monthly_data = []
        self.sort_column = None
        self.sort_reverse = False
        
        super().__init__(parent, f"月間データ詳細 - {year}年{month:02d}月")
        
        self._create_widgets()
        self._load_monthly_data()
    
    def _create_widgets(self):
        """ダイアログ内のUI要素を作成する"""
        # グリッドレイアウトの設定
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # タイトルセクション
        title_frame = tk.Frame(self, bg='#f0f0f0')
        title_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        
        title_label = tk.Label(title_frame,
                               text=f"{self.year}年{self.month:02d}月の詳細データ",
                               font=('Arial', 16, 'bold'),
                               bg='#f0f0f0')
        title_label.pack()
        
        # データ表示セクション
        result_frame = tk.Frame(self, bg='#f0f0f0')
        result_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        result_frame.grid_rowconfigure(0, weight=1)
        result_frame.grid_columnconfigure(0, weight=1)
        
        # Treeview
        columns = ["年月日", "項目", "支払先", "金額(円)", "メモ"]
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=15)
        
        # 重複データ用のタグを設定
        self.result_tree.tag_configure("duplicate", background="#ffcccc")
        
        # 列設定
        self.default_column_widths = {}
        widths = {
            "年月日": 100,
            "項目": 120,
            "支払先": 150,
            "金額(円)": 100,
            "メモ": 200
        }
        
        for col in columns:
            self.result_tree.heading(col, text=col, command=lambda c=col: self._sort_by_column(c))
            width = widths.get(col, 100)
            self.result_tree.column(col, anchor="center", width=width, minwidth=int(width*0.8))
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
        
        # データ件数表示
        self.result_label = tk.Label(self, text="データ: 0 件",
                                     font=('Arial', 10), bg='#f0f0f0', fg='#666666')
        self.result_label.grid(row=3, column=0, sticky="w", padx=10, pady=(5, 10))
        
        # 閉じるボタン
        close_btn = tk.Button(self, text="閉じる", font=('Arial', 12),
                              bg='#f44336', fg='white', relief='raised', bd=2,
                              activebackground='#d32f2f', command=self.destroy)
        close_btn.grid(row=4, column=0, pady=(0, 10), ipady=5)
        
        # キーボードショートカット
        self.bind('<Escape>', lambda e: self.destroy())
        self.result_tree.bind("<MouseWheel>", self._on_mousewheel)
        self.result_tree.bind("<Double-1>", self._on_double_click)
        self.result_tree.bind("<Button-3>", self._on_header_right_click)
    
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
    
    def _on_double_click(self, event):
        """行をダブルクリックした時の処理"""
        item = self.result_tree.selection()
        if not item:
            return
        
        row_index = self.result_tree.index(item[0])
        if row_index >= len(self.monthly_data):
            return
        
        data = self.monthly_data[row_index]
        year, month, day, col_index = data['sort_key']
        
        # 親アプリケーションの年月を変更（通常は同じ月だが念のため対応）
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
    
    def _on_mousewheel(self, event):
        """マウスホイールによるスクロール処理"""
        self.result_tree.yview_scroll(int(-1 * (event.delta / 120)), "units")
    
    def _sort_by_column(self, column):
        """指定された列でデータをソートする"""
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = False
        
        sort_key_map = {
            "年月日": lambda x: x['sort_key'],
            "項目": lambda x: x['column'],
            "支払先": lambda x: x['partner'],
            "金額(円)": lambda x: x['amount_value'],
            "メモ": lambda x: x['detail']
        }
        
        self.monthly_data.sort(key=sort_key_map.get(column, lambda x: ""),
                               reverse=self.sort_reverse)
        
        self._refresh_treeview()
        self._update_column_headers()
    
    def _refresh_treeview(self):
        """ソート後のデータでTreeviewを再表示する"""
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        
        for result in self.monthly_data:
            values = [result['date'], result['column'], result['partner'],
                      result['amount'], result['detail']]
            self.result_tree.insert("", "end", values=values)
        
        self._highlight_duplicates()
        self.result_label.config(text=f"データ: {len(self.monthly_data)} 件")
    
    def _update_column_headers(self):
        """ソート状態を示すため、列ヘッダーに矢印を表示する"""
        columns = ["年月日", "項目", "支払先", "金額(円)", "メモ"]
        
        for col in columns:
            if col == self.sort_column:
                arrow = " ▼" if self.sort_reverse else " ▲"
                self.result_tree.heading(col, text=f"{col}{arrow}")
            else:
                self.result_tree.heading(col, text=col)
    
    def _highlight_duplicates(self):
        """重複データを検出して薄い赤色でハイライトする（最適化版）"""
        items = self.result_tree.get_children()
        seen_data = {}
        
        # 1回のループで重複検出（2回ループを1回に削減して効率化）
        for item in items:
            values = self.result_tree.item(item, 'values')
            if not values:
                continue
            
            # タプルで直接比較（不要な文字列化を削減）
            data_key = tuple(values[:5])
            
            if data_key in seen_data:
                seen_data[data_key].append(item)
            else:
                seen_data[data_key] = [item]
        
        # 重複しているアイテムにタグを設定
        for item_list in seen_data.values():
            if len(item_list) > 1:
                for item_id in item_list:
                    current_tags = self.result_tree.item(item_id, 'tags')
                    
                    if isinstance(current_tags, str):
                        tag_list = [current_tags] if current_tags else []
                    else:
                        tag_list = list(current_tags) if current_tags else []

                    # duplicateタグを追加
                    if "duplicate" not in tag_list:
                        tag_list.append("duplicate")

                    # タグをタプルとして設定（Treeviewが期待する形式）
                    self.result_tree.item(item_id, tags=tuple(tag_list))

    def _load_monthly_data(self):
        """
        指定された年月のデータを読み込む。
        """
        # 既存の表示をクリア
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        self.monthly_data = []
        total_amount = 0  # 月間合計金額
        total_count = 0  # 取引件数

        # データ取得元を data_manager に変更
        for dict_key, data_list in self.parent_app.data_manager.data.items():
            try:
                # キーを解析（形式: "年-月-日-列インデックス"）
                parts = dict_key.split("-")
                if len(parts) == 4:
                    year = int(parts[0])
                    month = int(parts[1])
                    day = int(parts[2])
                    col_index = int(parts[3])

                    # まとめ行（day=0）は収入データなので除外
                    if day == 0:
                        continue

                    # 指定された年月のデータのみ処理
                    if year == self.year and month == self.month:
                        # 項目名を取得
                        all_columns = self.parent_app.get_all_columns()
                        column_name = all_columns[col_index] if col_index < len(all_columns) else f"列{col_index}"
                        date_str = f"{year}/{month:02d}/{day:02d}"

                        # 各取引データを処理
                        for row in data_list:
                            if len(row) >= 3:
                                partner = str(row[0]).strip() if row[0] else ""
                                amount_str = str(row[1]).strip() if row[1] else ""
                                detail = str(row[2]).strip() if row[2] else ""

                                # 【修正箇所】ここです！ self._parse_amount ではなく parse_amount を使います
                                amount_value = parse_amount(amount_str)

                                # 結果データを構造化
                                result = {
                                    'date': date_str,
                                    'column': column_name,
                                    'partner': partner,
                                    'amount': amount_str,
                                    'detail': detail,
                                    'amount_value': amount_value,
                                    'sort_key': (year, month, day, col_index)
                                }
                                self.monthly_data.append(result)
                                total_amount += amount_value
                                total_count += 1
            except (ValueError, IndexError):
                continue
            except Exception as e:
                # その他のエラー（AttributeErrorなど）をコンソールに出力して確認できるようにする
                print(f"Error loading row: {e}")
                continue

        # デフォルトで日付順にソート
        self.monthly_data.sort(key=lambda x: x['sort_key'])
        self.sort_column = "年月日"
        self.sort_reverse = False

        # データをTreeviewに表示
        for result in self.monthly_data:
            values = [result['date'], result['column'], result['partner'],
                      result['amount'], result['detail']]
            self.result_tree.insert("", "end", values=values)

        # 重複データをチェックしてハイライト
        self._highlight_duplicates()

        # 列ヘッダーを更新
        self._update_column_headers()

        # 統計情報を計算して表示
        if total_count > 0:
            avg_amount = total_amount / total_count
            self.stats_label.config(text=f"合計金額: ¥{total_amount:,} | 平均金額: ¥{avg_amount:.0f}")
        else:
            self.stats_label.config(text="データなし")

        # データ件数を表示
        self.result_label.config(text=f"データ: {len(self.monthly_data)} 件")