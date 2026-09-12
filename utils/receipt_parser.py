"""OCR座標から商品名と金額を復元する。認識結果は候補として扱う。"""
import re
import unicodedata


def compact(text):
    return re.sub(r'\s+', '', unicodedata.normalize('NFKC', text))


def amount(text):
    text = compact(text)
    match = re.fullmatch(r'[¥￥\\]?([0-9][0-9,]*)(?:円|軽|\*|※)?', text)
    return int(match.group(1).replace(',', '')) if match else None


def join_words(words):
    # OCRが日本語の1文字ごとに挿入した空白を除く。大きな単語間隔は残す。
    words = sorted(words, key=lambda w: w['x'])
    parts = []
    previous = None
    for word in words:
        if previous and word['x'] - previous['x'] - previous['width'] > max(previous['height'], word['height']) * .4:
            parts.append(' ')
        parts.append(unicodedata.normalize('NFKC', word['text']))
        previous = word
    return ''.join(parts)


def spatial_rows(layout):
    fragments = []
    for line in layout['lines']:
        words = line['words']
        if not words:
            continue
        top = min(w['y'] for w in words)
        bottom = max(w['y'] + w['height'] for w in words)
        fragments.append(dict(words=words, top=top, bottom=bottom, center=(top + bottom) / 2))
    groups = []
    for fragment in sorted(fragments, key=lambda f: f['center']):
        height = fragment['bottom'] - fragment['top']
        target = next((g for g in reversed(groups)
                       if abs(g['center'] - fragment['center']) < min(g['height'], height) * .55), None)
        if target:
            target['words'].extend(fragment['words'])
            target['top'] = min(target['top'], fragment['top'])
            target['bottom'] = max(target['bottom'], fragment['bottom'])
        else:
            groups.append(dict(words=fragment['words'][:], top=fragment['top'], bottom=fragment['bottom'],
                               center=fragment['center'], height=height))
    return sorted(groups, key=lambda g: g['top'])


NON_PRODUCT = re.compile(r'合計|小計|消費税|税等|対象|点数|領収|お預|お釣|還元|会員|ポイント|支払|クレジット|PayPay|現金|レジ|責[:、]|[0-9]{4}年', re.I)


def parse_layout(layout):
    candidates, other = [], []
    rows = spatial_rows(layout)
    first_product = False
    finished = False
    for row in rows:
        words = sorted(row['words'], key=lambda w: w['x'])
        raw = join_words(words)
        text = compact(raw)
        if NON_PRODUCT.search(text) or re.search(r'計|言十', text):
            if first_product and re.search(r'計|言十|消費税|お預', text):
                finished = True
            other.append(raw)
            continue
        # 右側の価格列だけを切り出す。商品名中の「120G」「108エン」は価格にしない。
        split = None
        for i in range(1, len(words)):
            gap = words[i]['x'] - words[i-1]['x'] - words[i-1]['width']
            if words[i]['x'] > layout['width'] * .55 and gap > row['height'] * .9:
                split = i
        name_words = words[:split] if split else words
        price_words = words[split:] if split else []
        price_text = ''.join(w['text'] for w in price_words)
        price = amount(price_text) if price_words else None
        name = join_words(name_words)
        # 左寄せの名前を含む行を候補として残す。金額欠落は空欄で要確認。
        if finished or words[0]['x'] > layout['width'] * .4 or not re.search(r'[ぁ-んァ-ヶ一-龠A-Za-z]', name):
            other.append(raw)
            continue
        if not first_product and price is None:
            other.append(raw)
            continue
        first_product = True
        candidates.append(dict(name=name, amount=price, raw=raw,
                               note='要確認' if price is not None else '金額を確認',
                               price_x=words[split]['x'] if split else None,
                               bounds=dict(x=0, y=max(0, int(row['top'] - 5)), width=int(layout['width']),
                                           height=min(int(row['bottom'] - row['top'] + 10), int(layout['height'] - max(0, row['top'] - 5))))))
    return dict(items=candidates, other=other, raw=[line['text'] for line in layout['lines']])


def recognize_products(path):
    from utils.receipt_ocr import recognize_layout
    layout = recognize_layout(path)
    parsed = parse_layout(layout)
    if parsed['items']:
        try:
            refined = recognize_layout(path, [item['bounds'] for item in parsed['items']])
            for item, region in zip(parsed['items'], refined['regions']):
                result = parse_layout(dict(lines=region['lines'], width=layout['width'], height=layout['height']))
                if len(result['items']) == 1:
                    better = result['items'][0]
                    if item['amount'] is not None and better['amount'] is not None and item['amount'] != better['amount']:
                        item['note'] = f"金額候補 {item['amount']} / {better['amount']}：要確認"
                    else:
                        item['amount'] = better['amount'] if better['amount'] is not None else item['amount']
                    item['name'] = better['name']
            # 数字だけの価格が日本語OCRに無視される場合は、価格列を英語OCRで再認識。
            positions = [item['price_x'] for item in parsed['items'] if item['price_x'] is not None]
            if positions:
                left = max(0, int(min(positions) - 12))
                price_regions = [dict(item['bounds'], x=left, width=int(layout['width']) - left, language='en')
                                 for item in parsed['items']]
                numbers = recognize_layout(path, price_regions)
                for item, region in zip(parsed['items'], numbers['regions']):
                    text = ''.join(w['text'] for line in region['lines'] for w in sorted(line['words'], key=lambda w: w['x']))
                    match = re.match(r'^[¥￥\\]?([0-9][0-9,]*)', compact(text))
                    if item['amount'] is None and match:
                        item['amount'] = int(match.group(1).replace(',', ''))
                        item['note'] = '価格欄の再認識・要確認'
        except (RuntimeError, ValueError):
            parsed['other'].append('行ごとの再認識に失敗しました。元のOCR結果を表示しています。')
    if parsed['items']:
        _refine_preprocessed(path, layout, parsed)
    parsed['width'], parsed['height'] = layout['width'], layout['height']
    return parsed


def _refine_preprocessed(path, layout, parsed):
    """細長い感熱紙フォントを横拡大・コントラスト補正して再照合する。"""
    import tempfile
    from pathlib import Path
    from PIL import Image, ImageOps, ImageEnhance
    from utils.receipt_ocr import recognize_layout
    try:
        with tempfile.TemporaryDirectory(prefix='household_ocr_') as directory:
            with Image.open(path) as image:
                prepared = ImageEnhance.Contrast(ImageOps.grayscale(ImageOps.exif_transpose(image))).enhance(1.5)
                prepared.thumbnail((2000, 2400))
                prepared = prepared.resize((int(prepared.width * 1.6), prepared.height), Image.Resampling.LANCZOS)
                target = Path(directory) / 'contrast.png'
                prepared.save(target)
            second = recognize_layout(target)
        positions = [item['price_x'] / layout['width'] for item in parsed['items'] if item['price_x'] is not None]
        boundary = (min(positions) - .015) if positions else .7
        rows = spatial_rows(second)
        for item in parsed['items']:
            bounds = item['bounds']
            center = (bounds['y'] + bounds['height'] / 2) / layout['height']
            nearby = [row for row in rows if abs(row['center'] / second['height'] - center) < bounds['height'] / layout['height'] * .45]
            if len(nearby) != 1:
                continue
            words = nearby[0]['words']
            names = [w for w in words if w['x'] / second['width'] < boundary]
            prices = [w for w in words if w['x'] / second['width'] >= boundary]
            if names:
                item['alternative_name'] = item['name']
                item['name'] = join_words(names)
            candidate = amount(''.join(w['text'] for w in sorted(prices, key=lambda w: w['x'])))
            if candidate is not None:
                if item['amount'] is None:
                    item['amount'] = candidate
                    item['note'] = '補正画像で再認識・要確認'
                elif item['amount'] != candidate:
                    item['note'] = f"金額候補 {item['amount']} / {candidate}：要確認"
    except (OSError, RuntimeError, ValueError):
        parsed['other'].append('画像補正での再認識に失敗しました。元のOCR結果を表示しています。')
