# ui/transaction_dialog.py
"""
取引詳細を入力・編集するダイアログ
"""
import tkinter as tk
from tkinter import ttk, messagebox
import json
from ui.base_dialog import BaseDialog
from config import DialogConfig, parse_amount


class TransactionDialog(BaseDialog):
    """
    取引詳細を入力・編集するダイアログ。
    
    特定の日付・項目の取引データ(支払先、金額、詳細)を
    テーブル形式で入力・編集できる。
    支払先の入力時には過去の履歴から候補を表示する。
    """
    
    def __init__(self, parent, parent_app, dict_key, col_name):
        """
        取引詳細ダイアログを初期化する。
        
        Args:
            parent: 親ウィンドウ
            parent_app: メインアプリケーションのインスタンス
            dict_key: データのキー(年-月-日-列インデックス)
            col_name: 項目名(表示用)
        """
        self.parent_app = parent_app
        self.dict_key = dict_key
        self.col_name = col_name
        self.entry_editor = None
        
        # 自動補完用の変数
        self.autocomplete_candidates = []  # 現在の候補リスト
        self.autocomplete_index = -1  # 現在の候補インデックス
        self.autocomplete_original_text = ""  # 元のテキスト
        
        # メモ候補をキャッシュ（初期化時に1回だけ収集・効率化）
        self._memo_candidates_cache = None
        
        # 元に戻す機能用
        self.undo_stack = []  # 操作履歴 [(action_type, data), ...]
        self.max_undo_count = 50  # 最大保持数
        
        # キーを解析して年月日と列インデックスを取得
        parts = dict_key.split("-")
        if len(parts) == 4:
            self.year = int(parts[0])
            self.month = int(parts[1])
            self.day = int(parts[2])
            self.col_index = int(parts[3])
        else:
            self.year = self.month = self.day = self.col_index = 0
        
        super().__init__(parent, f"支出・収入詳細 - {col_name}",
                         DialogConfig.TRANSACTION_WIDTH,
                         DialogConfig.TRANSACTION_HEIGHT)
        
        self._create_widgets()
    
    def _create_widgets(self):
        """取引詳細ダイアログのUI要素を作成する"""
        # グリッドレイアウトの設定
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # タイトル
        title_label = tk.Label(self, text=f"{self.col_name} の詳細入力",
                               font=('Arial', 16, 'bold'), bg='#f0f0f0')
        title_label.grid(row=0, column=0, pady=(10, 15), sticky="ew")
        
        # Treeviewコンテナ
        tree_container = tk.Frame(self, bg='#f0f0f0')
        tree_container.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)
        
        style = ttk.Style()
        style.map("Treeview",
                  background=[('selected', '#0078d4')],
                  foreground=[('selected', 'yellow')])


        # 取引データ編集用Treeview
        columns = ["支払先", "金額(円)", "メモ"]
        self.tree = ttk.Treeview(tree_container, columns=columns, show="headings", selectmode='extended')
        
        # 列の設定
        self.tree.heading("支払先", text="支払先")
        self.tree.heading("金額(円)", text="金額(円)")
        self.tree.heading("メモ", text="メモ")
        
        self.tree.column("支払先", anchor="center", width=150, minwidth=100)
        self.tree.column("金額(円)", anchor="center", width=100, minwidth=80)
        self.tree.column("メモ", anchor="center", width=200, minwidth=150)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        
        # スクロールバー
        scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # イベントバインド
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Button-3>", self._on_right_click)

        # キーボードショートカット
        self.tree.bind("<Control-c>", lambda e: self._copy_rows())
        self.tree.bind("<Control-x>", lambda e: self._cut_rows())
        self.tree.bind("<Control-v>", lambda e: self._paste_rows())
        self.tree.bind("<Control-z>", lambda e: self._undo())
        self.tree.bind("<Delete>", lambda e: self._delete_row())
        self.tree.bind("<space>", self._on_space_key)
        self.tree.bind("<Tab>", self._on_tab_key)
        
        # ボタンコンテナ
        button_frame = tk.Frame(self, bg='#f0f0f0')
        button_frame.grid(row=2, column=0, sticky="ew", pady=(15, 10), padx=10)
        
        # 行追加ボタン
        add_btn = tk.Button(button_frame, text="行を追加", font=('Arial', 12),
                            bg='#4caf50', fg='white', relief='raised', bd=2,
                            activebackground='#45a049', command=self._add_row)
        add_btn.pack(side=tk.LEFT, padx=(0, 10), ipady=5)
        
        # 右側のボタングループ
        right_button_frame = tk.Frame(button_frame, bg='#f0f0f0')
        right_button_frame.pack(side=tk.RIGHT)
        
        # キャンセルボタン
        cancel_btn = tk.Button(right_button_frame, text="キャンセル", font=('Arial', 12),
                               bg='#f44336', fg='white', relief='raised', bd=2,
                               activebackground='#d32f2f', command=self.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=(10, 0), ipady=5)
        
        # OKボタン
        ok_btn = tk.Button(right_button_frame, text="OK", font=('Arial', 12, 'bold'),
                           bg='#2196f3', fg='white', relief='raised', bd=2,
                           activebackground='#1976d2', command=self._on_ok)
        ok_btn.pack(side=tk.RIGHT, ipady=5)
        
        # 右クリックメニュー
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="コピー (Ctrl+C)", command=self._copy_rows)
        self.context_menu.add_command(label="切り取り (Ctrl+X)", command=self._cut_rows)
        self.context_menu.add_command(label="貼り付け (Ctrl+V)", command=self._paste_rows)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="元に戻す (Ctrl+Z)", command=self._undo)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="行を削除 (Delete)", command=self._delete_row)
        
        # 使用方法のヒント
        hint_label = tk.Label(self, text="使い方: ダブルクリック/SPACEで編集、TABで自動補完/次のセル移動、ENTERで確定",
                              font=('Arial', 10), fg='#666666', bg='#f0f0f0')
        hint_label.grid(row=3, column=0, pady=(5, 10))
        
        # キーボードショートカット - ENTERキーの処理を変更
        self.bind('<Return>', self._on_enter_key)
        self.bind('<Escape>', lambda e: self.destroy())
        self.tree.bind("<MouseWheel>", self._on_mousewheel)
        
        # データを読み込む
        self.after(50, self._load_data)
    
    # ENTERキーの処理
    def _on_enter_key(self, event):
        """
        ENTERキーが押された時の処理
        - 編集中の場合: 編集を確定
        - 編集中でない場合: OKボタンと同じ処理
        """
        if self.entry_editor:
            # 編集中の場合は何もしない（エディタ自身のReturnイベントで処理）
            return
        else:
            # 編集中でない場合はOK処理
            self._on_ok()
    
    def _on_mousewheel(self, event):
        """マウスホイールによるスクロール処理"""
        self.tree.yview_scroll(int(-1 * (event.delta / 120)), "units")
    
    def _load_data(self):
        """既存のデータを読み込んでTreeviewに表示する"""
        # 既存の表示をクリア
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # データを取得
        data_list = self.parent_app.data_manager.get_transaction_data(self.dict_key)
        
        if not data_list:
            # データがない場合は空行を1つ追加
            self.tree.insert("", "end", values=["", "", ""])
        else:
            # 既存データを表示
            for row in data_list:
                row_data = list(row) if row else ["", "", ""]
                while len(row_data) < 3:
                    row_data.append("")
                self.tree.insert("", "end", values=row_data)
            
            # 最後に空行を追加(新規入力用)
            self.tree.insert("", "end", values=["", "", ""])
    
    def _add_row(self):
        """新しい空行を追加する"""
        item_id = self.tree.insert("", "end", values=["", "", ""])
        self.tree.selection_set(item_id)
        self.tree.see(item_id)
    
    # 自動補完機能のメソッド群
    def _get_autocomplete_candidates(self, text, col_idx):
        """
        入力テキストに基づいて自動補完候補を取得する
        
        Args:
            text: 入力されたテキスト
            col_idx: 列インデックス（0:支払先、2:メモ）
            
        Returns:
            list: マッチする候補のリスト
        """
        if not text:
            return []
        
        text_lower = text.lower()
        
        if col_idx == 0:  # 支払先列
            # 支払先の履歴から候補を取得
            all_partners = self.parent_app.data_manager.get_transaction_partners_list()
            candidates = [p for p in all_partners if p.lower().startswith(text_lower)]
        elif col_idx == 2:  # メモ列
            # キャッシュされたメモ候補を使用（効率化）
            all_memos = self._get_memo_candidates()
            candidates = [m for m in all_memos if m.lower().startswith(text_lower)]
        else:
            candidates = []
        
        return sorted(candidates)
    
    def _collect_all_memos(self):
        """全取引データからメモを収集する（初期化時に1回だけ実行・効率化）"""
        memos = set()
        for data_list in self.parent_app.data_manager.data.values():
            for row in data_list:
                if len(row) >= 3 and row[2] and str(row[2]).strip():
                    memos.add(str(row[2]).strip())
        return sorted(list(memos))
    
    def _get_memo_candidates(self):
        """メモ候補を取得（キャッシュを使用）"""
        if self._memo_candidates_cache is None:
            self._memo_candidates_cache = self._collect_all_memos()
        return self._memo_candidates_cache
    
    def _handle_autocomplete_tab(self, event, item_id, col_idx):
        """
        TABキーによる自動補完を処理する
        
        Args:
            event: イベントオブジェクト
            item_id: 現在編集中のアイテムID
            col_idx: 列インデックス
            
        Returns:
            str: "break"を返してデフォルトのTAB動作を抑制
        """
        if not self.entry_editor:
            return
        
        current_text = self.entry_editor.get()
        
        # 初回のTAB押下時、または入力テキストが変更された場合
        if (self.autocomplete_index == -1 or 
            current_text != self.autocomplete_original_text):
            # 候補を新規取得
            self.autocomplete_candidates = self._get_autocomplete_candidates(current_text, col_idx)
            self.autocomplete_original_text = current_text
            self.autocomplete_index = -1
        
        # 候補がない場合は何もしない
        if not self.autocomplete_candidates:
            return "break"
        
        # 次の候補に進む
        self.autocomplete_index = (self.autocomplete_index + 1) % len(self.autocomplete_candidates)
        selected_candidate = self.autocomplete_candidates[self.autocomplete_index]
        
        # エディタのテキストを更新
        self.entry_editor.delete(0, tk.END)
        self.entry_editor.insert(0, selected_candidate)
        
        return "break"  # デフォルトのTAB動作を抑制
    
    def _reset_autocomplete(self):
        """自動補完状態をリセット"""
        self.autocomplete_candidates = []
        self.autocomplete_index = -1
        self.autocomplete_original_text = ""
    
    def _on_double_click(self, event):
        """セルのダブルクリックイベントを処理する"""
        if self.entry_editor:
            self._cancel_edit()
        
        item_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        
        if not item_id or not col_id:
            return
        
        bbox = self.tree.bbox(item_id, col_id)
        if not bbox:
            return
        
        col_idx = int(col_id[1:]) - 1
        values = list(self.tree.item(item_id, 'values'))
        
        while len(values) <= col_idx:
            values.append("")
        
        x, y, w, h = bbox
        
        # 自動補完状態をリセット
        self._reset_autocomplete()
        
        if col_idx == 0:  # 支払先列の場合
            self.entry_editor = ttk.Combobox(self.tree, font=("Arial", 11))
            partner_list = self.parent_app.data_manager.get_transaction_partners_list()
            self.entry_editor['values'] = partner_list
            self.entry_editor.set(values[col_idx])
        else:
            self.entry_editor = tk.Entry(self.tree, font=("Arial", 11))
            self.entry_editor.insert(0, str(values[col_idx]))
        
        self.entry_editor.place(x=x, y=y, width=w, height=h)
        self.entry_editor.focus_set()
        self.entry_editor.select_range(0, tk.END)
        
        # イベントバインド - ENTERキーで編集を確定し、イベント伝播を停止
        self.entry_editor.bind("<Return>", lambda e: self._save_edit_and_stop(item_id, col_idx))
        # TABキーは自動補完に使用し、フィールド移動はしない
        self.entry_editor.bind("<Tab>", lambda e: self._handle_autocomplete_tab(e, item_id, col_idx))
        self.entry_editor.bind("<FocusOut>", lambda e: self._save_edit(item_id, col_idx))
        self.entry_editor.bind("<Escape>", lambda e: self._cancel_edit())
        
        # テキスト変更時に自動補完状態をリセット（Combobox以外）
        if col_idx != 0:
            self.entry_editor.bind("<KeyRelease>", lambda e: self._on_text_change(e))
        
        self.tree.bind("<Button-1>", lambda e: self._save_edit(item_id, col_idx) if self.entry_editor else None, add=True)

    def _save_edit_and_stop(self, item_id, col_idx):
        """
        エディターの内容を保存し、イベント伝播を停止する
        
        Args:
            item_id: 編集中のアイテムID
            col_idx: 列インデックス
            
        Returns:
            str: "break"を返してイベント伝播を停止
        """
        self._save_edit(item_id, col_idx)
        return "break"

    
    def _on_text_change(self, event):
        """
        テキスト変更時の処理（TABキー以外）
        
        Args:
            event: イベントオブジェクト
        """
        # TABキーの場合は何もしない
        if event.keysym == "Tab":
            return
        
        # 自動補完状態をリセット
        self._reset_autocomplete()
    
    def _save_edit(self, item_id, col_idx):
        """エディターの内容を保存する"""
        if not self.entry_editor:
            return
        
        try:
            new_value = self.entry_editor.get()
            
            # 現在の値を取得（変更チェック用）
            old_values = list(self.tree.item(item_id, 'values'))
            while len(old_values) <= col_idx:
                old_values.append("")
            old_value = old_values[col_idx]
            
            self.entry_editor.destroy()
            self.entry_editor = None
            
            # 自動補完状態をリセット
            self._reset_autocomplete()
            
            # 値が変更されていれば元に戻す履歴を保存
            if old_value != new_value:
                # ダイアログ内の元に戻すスタックに保存
                current_data = []
                for item in self.tree.get_children():
                    vals = self.tree.item(item, 'values')
                    current_data.append(tuple(vals))
                
                self.undo_stack.append({
                    'type': 'edit',
                    'item_id': item_id,
                    'col_idx': col_idx,
                    'old_value': old_value,
                    'new_value': new_value
                })
                
                if len(self.undo_stack) > self.max_undo_count:
                    self.undo_stack.pop(0)
                
                # メインウィンドウにも保存（ダイアログを閉じた後用）
                old_data = self.parent_app.data_manager.get_transaction_data(self.dict_key)
                self.parent_app._save_undo_state('detail_edit', [(self.dict_key, old_data[:] if old_data else None)])
            
            # 支払先の場合は履歴に追加
            if col_idx == 0 and new_value.strip():
                self.parent_app.data_manager.add_transaction_partner(new_value.strip())
            
            # Treeviewの値を更新
            values = list(self.tree.item(item_id, 'values'))
            while len(values) <= col_idx:
                values.append("")
            
            values[col_idx] = new_value
            self.tree.item(item_id, values=values)
            
            # 値が変更されていれば即座にメインウィンドウに反映
            if old_value != new_value:
                self._apply_changes_to_parent()
            
            # 最終行に入力があったら新しい空行を追加
            all_items = self.tree.get_children()
            if item_id == all_items[-1] and any(str(cell).strip() for cell in values):
                self._add_row()
            
            self.tree.unbind("<Button-1>")
        except Exception:
            if self.entry_editor:
                self.entry_editor.destroy()
                self.entry_editor = None
            self._reset_autocomplete()
    
    def _cancel_edit(self):
        """編集をキャンセルしてエディターを削除する"""
        if self.entry_editor:
            try:
                self.entry_editor.destroy()
            except:
                pass
            self.entry_editor = None
        
        # 自動補完状態をリセット
        self._reset_autocomplete()
        
        try:
            self.tree.unbind("<Button-1>")
        except:
            pass
    
    def _on_right_click(self, event):
        """右クリックイベントを処理する"""
        # クリック位置のアイテムを選択状態にする（選択がない場合のみ）
        item_id = self.tree.identify_row(event.y)
        if item_id:
            # すでに選択されているアイテムの上で右クリックした場合は選択を維持
            if item_id not in self.tree.selection():
                self.tree.selection_set(item_id)
            self.context_menu.post(event.x_root, event.y_root)
    
    def _delete_row(self):
        """選択された行を削除する（確認なし）"""
        selected_items = self.tree.selection()
        if not selected_items:
            return
        
        # 元に戻す用に削除前のデータを保存
        undo_data = []
        for item in selected_items:
            values = self.tree.item(item, 'values')
            index = self.tree.index(item)
            undo_data.append((index, list(values)))
        
        self._save_undo_state('delete', undo_data)
        
        # 削除を実行
        for item in selected_items:
            self.tree.delete(item)

    def _on_space_key(self, event):
        """
        SPACEキーが押された時の処理
        編集モードでない場合、選択中のセルを編集モードにする
        
        Returns:
            str: "break"を返してデフォルトのSPACE動作を抑制
        """
        # 既に編集モードの場合は何もしない
        if self.entry_editor:
            return "break"
        
        selected_items = self.tree.selection()
        if not selected_items:
            return "break"
        
        item_id = selected_items[0]
        
        # 現在フォーカスされている列を取得
        # Treeviewのfocusとselectionだけでは列が分からないため、
        # 最初の列（支払先）を編集する
        col_id = "#1"  # 支払先列
        
        # ダブルクリックイベントを模擬
        bbox = self.tree.bbox(item_id, col_id)
        if bbox:
            # 疑似イベントオブジェクトを作成
            class FakeEvent:
                def __init__(self, x, y):
                    self.x = x
                    self.y = y
            
            x, y, w, h = bbox
            fake_event = FakeEvent(x + w//2, y + h//2)
            self._on_double_click(fake_event)
        
        return "break"
    
    def _on_tab_key(self, event):
        """
        TABキーが押された時の処理
        編集モードでない場合、次のセルに移動する
        
        Returns:
            str: "break"を返してデフォルトのTAB動作を抑制
        """
        # 編集モードの場合は自動補完処理に任せる
        if self.entry_editor:
            return
        
        selected_items = self.tree.selection()
        if not selected_items:
            return "break"
        
        current_item = selected_items[0]
        all_items = self.tree.get_children()
        
        if not all_items:
            return "break"
        
        # 現在のアイテムのインデックスを取得
        try:
            current_index = all_items.index(current_item)
        except ValueError:
            return "break"
        
        # 次のアイテムに移動
        if current_index < len(all_items) - 1:
            next_item = all_items[current_index + 1]
        else:
            # 最後の行の場合は最初の行に戻る
            next_item = all_items[0]
        
        # 選択を更新
        self.tree.selection_set(next_item)
        self.tree.focus(next_item)
        self.tree.see(next_item)
        
        return "break"
    
    def _copy_rows(self):
        """選択された行をコピーする"""
        selected_items = self.tree.selection()
        if not selected_items:
            return
            
        # データを収集（リストのリスト形式）
        rows_data = []
        for item in selected_items:
            values = self.tree.item(item, 'values')
            # 値を文字列として取得
            row = [str(v) for v in values]
            rows_data.append(row)
            
        if rows_data:
            # JSON形式でクリップボードにコピー
            json_str = json.dumps(rows_data, ensure_ascii=False)
            self.clipboard_clear()
            self.clipboard_append(json_str)
            self.update()

    def _paste_rows(self):
        """クリップボードから行を貼り付ける（メインウィンドウからの貼り付けもサポート）"""
        try:
            clipboard_text = self.clipboard_get()
        except tk.TclError:
            return

        new_data = []
        is_main_window_data = False
        
        # JSON形式として解析
        try:
            parsed = json.loads(clipboard_text)
            if isinstance(parsed, list):
                # メインウィンドウからのコピーデータかチェック
                if parsed and isinstance(parsed[0], dict) and 'data' in parsed[0]:
                    # メインウィンドウからのデータ形式
                    is_main_window_data = True
                    # 各セルのdataフィールドから行データを取得
                    for cell in parsed:
                        cell_data = cell.get('data', [])
                        for row in cell_data:
                            if len(row) >= 2:
                                new_data.append(row[:3] if len(row) >= 3 else row + [""])
                else:
                    # 詳細入力ウィンドウからのデータ形式
                    new_data = parsed
        except json.JSONDecodeError:
            pass
            
        # JSONでなければタブ区切りテキスト（Excel等）として解析
        if not new_data and clipboard_text:
            lines = clipboard_text.strip().split('\n')
            for line in lines:
                cols = line.split('\t')
                # 少なくとも1つの列があれば採用
                if cols:
                    # 3列になるように調整（支払先, 金額, メモ）
                    while len(cols) < 3:
                        cols.append("")
                    new_data.append(cols[:3])

        # 元に戻す用に貼り付け前の状態を保存
        if new_data:
            undo_data = []
            insert_count = 0
            
            for row in new_data:
                # [支払先, 金額, メモ] の形式であることを確認
                if isinstance(row, list) and len(row) >= 2: # 少なくとも支払先と金額
                    safe_row = [str(val).strip() for val in row]
                    while len(safe_row) < 3:
                        safe_row.append("")
                    
                    # 空行（入力用）の直前に挿入するのが理想的だが、末尾追加でOK
                    # 現在の末尾の空行を探す
                    items = self.tree.get_children()
                    last_item = items[-1] if items else None
                    last_vals = self.tree.item(last_item, 'values') if last_item else None
                    
                    # 末尾が完全な空行なら、その手前に挿入
                    if last_vals and all(v == "" for v in last_vals):
                        insert_index = items.index(last_item)
                        self.tree.insert("", insert_index, values=safe_row)
                    else:
                        insert_index = len(items)
                        self.tree.insert("", "end", values=safe_row)
                    
                    # 元に戻す用にインデックスを記録
                    undo_data.append((insert_index + insert_count, safe_row))
                    insert_count += 1
            
            # 元に戻すスタックに保存
            if undo_data:
                self._save_undo_state('paste', undo_data)
    

    def _cut_rows(self):
        """選択された行を切り取る"""
        selected_items = self.tree.selection()
        if not selected_items:
            return
        
        # まずコピー
        self._copy_rows()
        
        # 元に戻す用にデータを保存
        undo_data = []
        for item in selected_items:
            values = self.tree.item(item, 'values')
            index = self.tree.index(item)
            undo_data.append((index, list(values)))
        
        self._save_undo_state('cut', undo_data)
        
        # 削除
        for item in selected_items:
            self.tree.delete(item)
    
    def _save_undo_state(self, action_type, rows_data):
        """
        操作前の状態をundo stackに保存
        
        Args:
            action_type: 'cut', 'paste', 'delete'のいずれか
            rows_data: [(index, values), ...] の形式
        """
        undo_entry = {
            'action': action_type,
            'rows': rows_data
        }
        
        self.undo_stack.append(undo_entry)
        
        # 最大保持数を超えたら古いものから削除
        if len(self.undo_stack) > self.max_undo_count:
            self.undo_stack.pop(0)

    def _apply_changes_to_parent(self):
        """
        現在のダイアログのデータをメインウィンドウに即座に反映
        """
        # すべての行データを収集
        all_rows = []
        for item_id in self.tree.get_children():
            values = self.tree.item(item_id, 'values')
            row = list(values)
            while len(row) < 3:
                row.append("")
            all_rows.append(tuple(row))
        
        # 空行を除去
        filtered_rows = [row for row in all_rows if any(str(cell).strip() for cell in row)]
        
        if not filtered_rows:
            # データが空の場合
            dict_key_day = f"{self.year}-{self.month}-{self.day}"
            self.parent_app.update_parent_cell(dict_key_day, self.col_index, "")
            self.parent_app.data_manager.delete_transaction_data(self.dict_key)
        else:
            # データがある場合
            self.parent_app.data_manager.set_transaction_data(self.dict_key, filtered_rows)
            
            # 金額列(インデックス1)を合計
            total = sum(parse_amount(row[1]) for row in filtered_rows if len(row) > 1)
            
            # 親セルを更新
            dict_key_day = f"{self.year}-{self.month}-{self.day}"
            display_value = str(total) if total != 0 else ""
            self.parent_app.update_parent_cell(dict_key_day, self.col_index, display_value)
    
    def _undo(self):
        """
        最後の操作を元に戻す (Ctrl+Z)
        """
        if not self.undo_stack:
            return
        
        undo_entry = self.undo_stack.pop()
        undo_type = undo_entry['type']
        
        if undo_type == 'edit':
            # セル編集の取り消し
            item_id = undo_entry['item_id']
            col_idx = undo_entry['col_idx']
            old_value = undo_entry['old_value']
            
            # 値を元に戻す
            try:
                values = list(self.tree.item(item_id, 'values'))
                values[col_idx] = old_value
                self.tree.item(item_id, values=values)
                
                # メインウィンドウにも反映
                self._apply_changes_to_parent()
                
                # メインウィンドウの元に戻すスタックからも削除
                if self.parent_app.undo_stack:
                    last = self.parent_app.undo_stack[-1]
                    if last['action'] == 'detail_edit':
                        self.parent_app.undo_stack.pop()
            except:
                pass

    def _reload_tree_without_clearing_selection(self):
        """
        選択状態を保持したままTreeviewのデータを再読み込み
        """
        # 現在の選択を保存
        selected_items = self.tree.selection()
        selected_indices = []
        for item in selected_items:
            try:
                selected_indices.append(self.tree.index(item))
            except:
                pass
        
        # 既存の表示をクリア
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # データを取得
        data_list = self.parent_app.data_manager.get_transaction_data(self.dict_key)
        
        if not data_list:
            # データがない場合は空行を1つ追加
            self.tree.insert("", "end", values=["", "", ""])
        else:
            # 既存データを表示
            for row in data_list:
                row_data = list(row) if row else ["", "", ""]
                while len(row_data) < 3:
                    row_data.append("")
                self.tree.insert("", "end", values=row_data)
            
            # 最後に空行を追加(新規入力用)
            self.tree.insert("", "end", values=["", "", ""])
        
        # 選択を復元
        all_items = self.tree.get_children()
        for idx in selected_indices:
            if idx < len(all_items):
                self.tree.selection_add(all_items[idx])

    def _on_ok(self):
        """OKボタンの処理"""
        # 編集前のデータを保存（元に戻す用）
        old_data = self.parent_app.data_manager.get_transaction_data(self.dict_key)

        # データを反映（メモリのみ）
        self._apply_changes_to_parent()

        # データが変更されていればファイルに保存し、undo履歴に記録
        new_data = self.parent_app.data_manager.get_transaction_data(self.dict_key)
        if old_data != new_data:
            self.parent_app.data_manager.save_transaction(self.dict_key)
            self.parent_app._save_undo_state('edit_detail', [(self.dict_key, old_data[:] if old_data else None)])

        self.destroy()