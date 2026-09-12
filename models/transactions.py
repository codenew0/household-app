"""明細の共通形式。先頭3項目は既存データと互換。"""
from config import parse_amount

PENDING = "未確定"


def normalize(row):
    values = [str(value) if value is not None else "" for value in row[:7]]
    return values + [""] * (7 - len(values))


def is_pending(row):
    return len(row) > 5 and row[5] == PENDING


def cash_amount(row):
    """確定済みの現金支出 = ポイント利用前の金額 − 利用ポイント。"""
    if is_pending(row):
        return 0
    points = parse_amount(row[3]) if len(row) > 3 else 0
    return parse_amount(row[1]) - points


def describe(row):
    values = normalize(row)
    details = [values[2]]
    if values[3] and parse_amount(values[3]):
        details.append(f"ポイント利用前: {values[1]}円")
        details.append(f"ポイント利用: {values[3]} pt")
        details.append(f"お金での支払額: {parse_amount(values[1]) - parse_amount(values[3])}円")
    if values[4]:
        details.append(f"支払方法: {values[4]}")
    if is_pending(values):
        details.append(f"未確定 {values[1]}円（支出合計対象外）")
    return " / ".join(part for part in details if part)


def validate(row):
    values = normalize(row)
    for index, label in ((1, "金額"), (3, "ポイント")):
        text = values[index].replace(",", "").replace("¥", "").strip()
        try:
            number = int(text or "0")
        except ValueError:
            raise ValueError(f"{label}は整数で入力してください。")
        if index == 3 and number < 0:
            raise ValueError("ポイントは0以上で入力してください。")
        values[index] = str(number)
    if int(values[3]) > max(0, int(values[1])):
        raise ValueError("利用ポイントはポイント利用前の金額以下で入力してください。")
    return values
