# ui/base_dialog.py
"""
ダイアログの基底クラス
すべてのダイアログで共通の機能を提供
"""
import tkinter as tk
from config import DialogConfig


class BaseDialog(tk.Toplevel):
    """
    すべてのダイアログの基底クラス。
    
    共通の初期化処理(中央配置、モーダル設定など)を提供し、
    各ダイアログクラスはこのクラスを継承することで、
    一貫したUIと動作を実現する。
    """
    
    def __init__(self, parent, title, width=None, height=None):
        """
        ダイアログの基本設定を行う。
        
        Args:
            parent: 親ウィンドウ
            title: ダイアログのタイトル
            width: ダイアログの幅(デフォルト: DialogConfig.DEFAULT_WIDTH)
            height: ダイアログの高さ(デフォルト: DialogConfig.DEFAULT_HEIGHT)
        """
        super().__init__(parent)
        self.title(title)
        self.configure(bg='#f0f0f0')
        
        # デフォルトサイズの設定
        if width is None:
            width = DialogConfig.DEFAULT_WIDTH
        if height is None:
            height = DialogConfig.DEFAULT_HEIGHT
        
        # ダイアログを親ウィンドウの中央に配置
        self._center_on_parent(width, height)
        
        # モーダルダイアログとして設定
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)
        
        # ダイアログを前面に表示
        self.lift()
        self.focus_force()
    
    def _center_on_parent(self, width, height):
        """
        ダイアログを親ウィンドウの中央に配置する。
        
        画面外にはみ出さないように位置を調整し、
        ユーザーが見やすい位置に表示する。
        """
        # 親ウィンドウの位置とサイズを取得
        parent_x = self.master.winfo_x()
        parent_y = self.master.winfo_y()
        parent_w = self.master.winfo_width()
        parent_h = self.master.winfo_height()
        
        # 中央配置の計算
        x = parent_x + (parent_w - width) // 2
        y = parent_y + (parent_h - height) // 2
        
        # ダイアログの位置とサイズを設定
        self.geometry(f"{width}x{height}+{x}+{y}")
        
        # 最小サイズを設定(元のサイズの75%)
        min_width = int(width * DialogConfig.MIN_SIZE_RATIO)
        min_height = int(height * 0.67)
        self.minsize(min_width, min_height)

    def navigate_to_cell(self, parent_app, day, col_index, delay=False):
        """
        メインウィンドウの指定セルに移動して選択状態にする共通メソッド。

        Args:
            parent_app: MainWindowのインスタンス
            day: 移動先の日（0=まとめ行）
            col_index: 移動先の列インデックス
            delay: Trueの場合、50ms遅延してから実行（月切り替え後の再描画待ち用）
        """
        if delay:
            self.after(50, lambda: self._do_navigate(parent_app, day, col_index))
        else:
            self._do_navigate(parent_app, day, col_index)

    def _do_navigate(self, parent_app, day, col_index):
        """実際のナビゲーション処理"""
        if not parent_app.tree:
            return

        items = parent_app.tree.get_children()
        if not items:
            return

        target_item = None

        if day == 0:
            target_item = items[-1]
        else:
            for item in items[:-2]:
                values = parent_app.tree.item(item, 'values')
                if values and str(values[0]).strip() == str(day):
                    target_item = item
                    break

        if target_item:
            parent_app.tree.selection_set(target_item)
            parent_app.tree.see(target_item)
            parent_app.tree.focus(target_item)
            # 列のハイライト・コピペ等で使われる選択状態を更新
            col_id = f"#{col_index + 1}"
            parent_app.selected_column_id = col_id
            parent_app.selection_start_col = col_id
            parent_app.ctrl_selected_cells = [(target_item, col_id)]