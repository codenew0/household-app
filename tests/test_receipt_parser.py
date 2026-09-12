import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils.receipt_parser import amount, parse_layout


def line(text, x, y, width=100):
    return {'text': text, 'words': [dict(text=text, x=x, y=y, width=width, height=30)]}


class ReceiptParserTests(unittest.TestCase):
    def test_separate_columns_and_scrambled_order(self):
        layout = {'width': 1000, 'height': 800, 'lines': [
            line('119軽', 780, 143), line('108軽', 780, 100),
            line('牛乳400ml', 50, 100, 200), line('アイス', 50, 140),
            line('合計', 50, 200), line('227', 780, 203),
            line('ポイント残高', 50, 240), line('171P', 780, 244)]}
        result = parse_layout(layout)
        self.assertEqual([(i['name'], i['amount']) for i in result['items']],
                         [('牛乳400ml', 108), ('アイス', 119)])

    def test_missing_price_not_confused_with_product_size(self):
        layout = {'width': 1000, 'height': 800, 'lines': [
            line('トマト', 50, 100), line('198軽', 780, 100),
            line('菜箸33cm', 50, 140, 260)]}
        result = parse_layout(layout)
        self.assertEqual(result['items'][1]['name'], '菜箸33cm')
        self.assertIsNone(result['items'][1]['amount'])

    def test_amount_suffixes_and_ocr_digit_spacing(self):
        self.assertEqual(amount('1 1 9 軽'), 119)
        self.assertEqual(amount('￥ 1,187'), 1187)
        self.assertIsNone(amount('171P'))
        self.assertIsNone(amount('400ml'))

    def test_no_guessed_price_when_only_header(self):
        result = parse_layout({'width': 1000, 'height': 800, 'lines': [line('2020年4月24日', 100, 10)]})
        self.assertEqual(result['items'], [])


if __name__ == '__main__':
    unittest.main()
