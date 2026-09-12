# ui/tooltip.py
"""
Treeview用のツールチップ機能
"""
import tkinter as tk
from models.transactions import cash_amount, describe, normalize, is_pending, validate
from config import parse_amount


class TreeviewTooltip:
    """
    Treeviewのセルにマウスオーバーした時に詳細情報を表示するツールチップ機能。
    
    各セルの金額の内訳(支払先、金額、詳細)を表示し、
    合計行では日別の内訳、まとめ行では収入・支出の詳細を表示する。
    """
    
    def __init__(self, treeview, parent_app):
        """
        ツールチップの初期化。
        
        Args:
            treeview: ツールチップを適用するTreeviewウィジェット
            parent_app: メインアプリケーションのインスタンス(データアクセス用)
        """
        self.treeview = treeview
        self.parent_app = parent_app
        self.tooltip_window = None
        self.current_item = None
        self.current_column = None
        
        # マウスイベントをバインド
        self.treeview.bind('<Motion>', self._on_mouse_motion)
        self.treeview.bind('<Leave>', self._on_mouse_leave)
    
    def _on_mouse_motion(self, event):
        """マウスがTreeview上で移動した時の処理"""
        item = self.treeview.identify_row(event.y)
        column = self.treeview.identify_column(event.x)
        
        if not item or not column:
            self._hide_tooltip()
            return
        
        if item == self.current_item and column == self.current_column:
            return
        
        self.current_item = item
        self.current_column = column
        
        col_index = int(column[1:]) - 1
        all_columns = self.parent_app.get_all_columns()
        
        if col_index == 0 or col_index >= len(all_columns):
            self._hide_tooltip()
            return
        
        row_values = self.treeview.item(item, 'values')
        if not row_values:
            self._hide_tooltip()
            return
        
        cell_value = str(row_values[col_index]).strip() if col_index < len(row_values) else ""
        if not cell_value or cell_value == "0":
            self._hide_tooltip()
            return
        
        items = self.treeview.get_children()
        if len(items) < 2:
            self._hide_tooltip()
            return
        
        total_row_id = items[-2]
        summary_row_id = items[-1]
        
        if item == total_row_id:
            self._show_total_tooltip(event, col_index)
        elif item == summary_row_id:
            if col_index == 3:
                self._show_income_tooltip(event)
            elif col_index == 5:
                self._show_expense_tooltip(event)
            else:
                self._hide_tooltip()
        else:
            import re
            try:
                m = re.search(r'\d+', str(row_values[0]))
                if not m:
                    self._hide_tooltip()
                    return
                day = int(m.group())
                self._show_detail_tooltip(event, day, col_index)
            except ValueError:
                self._hide_tooltip()
    
    def _show_detail_tooltip(self, event, day, col_index):
        """通常セルの詳細情報をツールチップで表示"""
        dict_key = f"{self.parent_app.current_year}-{self.parent_app.current_month}-{day}-{col_index}"
        data_list = self.parent_app.data_manager.get_transaction_data(dict_key)
        
        if not data_list:
            self._hide_tooltip()
            return
        
        lines = []
        total = 0
        
        for row in data_list:
            if len(row) >= 3:
                partner = str(row[0]).strip() if row[0] else "(未入力)"
                amount_str = str(row[1]).strip() if row[1] else "0"
                detail = describe(row)
                
                try:
                    amount = cash_amount(row)
                    total += amount
                    amount_display = f"¥{amount:,}"
                except ValueError:
                    amount_display = amount_str
                
                line = f"• {partner}: {amount_display}"
                if detail:
                    line += f" ({detail})"
                lines.append(line)
        
        if len(data_list) > 1:
            lines.append("─" * 30)
            lines.append(f"合計: ¥{total:,}")
        
        self._show_tooltip(event, "\n".join(lines))
    
    def _show_total_tooltip(self, event, col_index):
        """合計行のツールチップを表示"""
        all_columns = self.parent_app.get_all_columns()
        column_name = all_columns[col_index] if col_index < len(all_columns) else "不明"
        
        days_with_data = []
        total = 0
        days_in_month = self.parent_app.get_days_in_month()
        
        for day in range(1, days_in_month + 1):
            dict_key = f"{self.parent_app.current_year}-{self.parent_app.current_month}-{day}-{col_index}"
            data_list = self.parent_app.data_manager.get_transaction_data(dict_key)
            if data_list:
                day_total = sum(cash_amount(row) for row in data_list if len(row) > 1)
                if day_total > 0:
                    days_with_data.append(f"{day}日: ¥{day_total:,}")
                    total += day_total
        
        if days_with_data:
            lines = [f"【{column_name}の内訳】"]
            lines.extend(days_with_data[:10])
            if len(days_with_data) > 10:
                lines.append(f"... 他{len(days_with_data) - 10}日分")
            lines.append("─" * 30)
            lines.append(f"合計: ¥{total:,}")
            self._show_tooltip(event, "\n".join(lines))
        else:
            self._hide_tooltip()
    
    def _show_income_tooltip(self, event):
        """収入セルのツールチップを表示"""
        dict_key = f"{self.parent_app.current_year}-{self.parent_app.current_month}-0-3"
        data_list = self.parent_app.data_manager.get_transaction_data(dict_key)
        
        if not data_list:
            self._hide_tooltip()
            return
        
        lines = ["【収入の内訳】"]
        total = 0
        
        for row in data_list:
            if len(row) >= 3:
                source = str(row[0]).strip() if row[0] else "(未入力)"
                amount = cash_amount(row)
                detail = describe(row)
                
                total += amount
                line = f"• {source}: ¥{amount:,}"
                if detail:
                    line += f" ({detail})"
                lines.append(line)
        
        if len(data_list) > 1:
            lines.append("─" * 30)
            lines.append(f"合計: ¥{total:,}")
        
        self._show_tooltip(event, "\n".join(lines))
    
    def _show_expense_tooltip(self, event):
        """支出合計のツールチップを表示"""
        lines = ["【支出の内訳】"]
        all_columns = self.parent_app.get_all_columns()
        grand_total = 0
        days_in_month = self.parent_app.get_days_in_month()
        
        for col_index in range(1, len(all_columns)):
            column_total = 0
            column_name = all_columns[col_index]
            
            for day in range(1, days_in_month + 1):
                dict_key = f"{self.parent_app.current_year}-{self.parent_app.current_month}-{day}-{col_index}"
                data_list = self.parent_app.data_manager.get_transaction_data(dict_key)
                if data_list:
                    for row in data_list:
                        if len(row) > 1:
                            column_total += cash_amount(row)
            
            if column_total > 0:
                lines.append(f"• {column_name}: ¥{column_total:,}")
                grand_total += column_total
        
        lines.append("─" * 30)
        lines.append(f"合計: ¥{grand_total:,}")
        self._show_tooltip(event, "\n".join(lines))
    
    def _show_tooltip(self, event, text):
        """ツールチップウィンドウを表示する"""
        self._hide_tooltip()
        
        self.tooltip_window = tk.Toplevel(self.treeview)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{event.x_root + 10}+{event.y_root + 10}")
        
        label = tk.Label(self.tooltip_window,
                         text=text,
                         justify=tk.LEFT,
                         background="#ffffcc",
                         relief=tk.SOLID,
                         borderwidth=1,
                         font=("Arial", 9))
        label.pack()
    
    def _hide_tooltip(self):
        """ツールチップを非表示にする"""
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None
        self.current_item = None
        self.current_column = None
    
    def _on_mouse_leave(self, event):
        """マウスがTreeviewから離れた時の処理"""
        self._hide_tooltip()
