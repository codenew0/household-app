"""メイン表で扱うクリップボード形式の解析。"""
import json

from models.transactions import validate


PLAIN_AMOUNT = "plain_amount"
DETAIL_ROWS = "detail_rows"
CELL_BLOCK = "cell_block"
EMPTY = "empty"


def decode_clipboard(text):
    """文字列を、単一金額・詳細行・メイン表セルのいずれかに変換する。"""
    if not text.strip():
        return EMPTY, []
    try:
        value = json.loads(text)
    except json.JSONDecodeError:
        return PLAIN_AMOUNT, [validate(["貼付入力", text, ""])]

    if not isinstance(value, list):
        return PLAIN_AMOUNT, [validate(["貼付入力", text, ""])]
    if not value:
        return EMPTY, []
    if all(isinstance(row, (list, tuple)) for row in value):
        if any(len(row) < 2 for row in value):
            raise ValueError("詳細明細には支払先と金額が必要です。")
        return DETAIL_ROWS, [validate(row) for row in value]
    if not all(isinstance(cell, dict) for cell in value):
        raise ValueError("クリップボードの形式が不正です。")

    cells = []
    for cell in value:
        day = cell.get("day")
        column = cell.get("col_idx")
        rows = cell.get("data", [])
        if not isinstance(day, int) or not isinstance(column, int) or not isinstance(rows, list):
            raise ValueError("クリップボードのセル情報が不正です。")
        cells.append({"day": day, "col_idx": column,
                      "data": [validate(row) for row in rows]})
    return CELL_BLOCK, cells
