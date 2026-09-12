import copy
import datetime
import os
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import config
from models.data_manager import DataManager
from models.transactions import cash_amount, normalize, PENDING, validate


class FeatureTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root_dir = self.directory.name
        self.patch = patch.multiple(config, JSON_DIR=self.root_dir,
                                    SETTINGS_FILE=os.path.join(self.root_dir, 'settings.json'))
        self.patch.start()
        self.addCleanup(self.patch.stop)
        self.manager = DataManager()

    def subscription(self, day=31, start='2026-01-01'):
        return self.manager.save_subscription('動画', '配信会社', '1000', day, 'カード', start)

    def test_month_end_catchup_pending_and_restart_idempotency(self):
        self.subscription()
        self.assertEqual(self.manager.generate_due_subscriptions(datetime.date(2026, 3, 1)), 2)
        self.assertIn('2026-2-28-7', self.manager.data)
        self.assertEqual(sum(cash_amount(row) for rows in self.manager.data.values() for row in rows), 0)
        other = DataManager()
        other.load_settings()
        other.load_data()
        self.assertEqual(other.generate_due_subscriptions(datetime.date(2026, 3, 1)), 0)
        self.assertEqual(len(other.data), 2)

    def test_deleted_occurrence_stays_deleted(self):
        self.subscription()
        self.manager.generate_due_subscriptions(datetime.date(2026, 1, 31))
        self.manager.delete_transaction_data('2026-1-31-7')
        self.manager.save_transaction('2026-1-31-7')
        other = DataManager()
        other.load_settings()
        other.load_data()
        self.assertEqual(other.generate_due_subscriptions(datetime.date(2026, 1, 31)), 0)
        self.assertEqual(other.data, {})

    def test_start_date_disabled_and_leap_year(self):
        s = self.subscription(start='2024-02-01')
        self.assertEqual(self.manager.generate_due_subscriptions(datetime.date(2024, 2, 28)), 0)
        self.assertEqual(self.manager.generate_due_subscriptions(datetime.date(2024, 2, 29)), 1)
        self.manager.subscriptions[0]['enabled'] = False
        self.assertEqual(self.manager.generate_due_subscriptions(datetime.date(2024, 5, 31)), 0)

    def test_points_pending_json_csv_roundtrip(self):
        rows = [['店', '1000', '商品', '200', 'カード', '', ''],
                ['動画', '1000', '契約', '0', 'カード', PENDING, 'id:2026-07']]
        self.manager.data = {'2026-7-1-7': rows}
        self.manager.save_data()
        self.manager.load_data()
        self.assertEqual(self.manager.data['2026-7-1-7'], rows)
        self.assertEqual(sum(map(cash_amount, rows)), 800)
        path = os.path.join(self.root_dir, 'roundtrip.csv')
        self.manager.export_csv(path, (2026, 7), (2026, 7), config.DefaultColumns.ITEMS)
        self.manager.import_csv(path, (2026, 7), (2026, 7), 'replace', 12)
        self.assertEqual(self.manager.data['2026-7-1-7'], rows)
        self.assertEqual(sum(int(item['amount']) for item in self.manager.search_transactions('')), 800)

    def test_save_failure_does_not_mark_occurrence_generated(self):
        self.subscription()
        with patch.object(self.manager, 'save_transactions', side_effect=OSError('disk full')):
            with self.assertRaises(OSError):
                self.manager.generate_due_subscriptions(datetime.date(2026, 1, 31))
        self.assertEqual(self.manager.data, {})
        self.assertEqual(self.manager.subscription_runs, set())
        self.assertEqual(self.manager.generate_due_subscriptions(datetime.date(2026, 1, 31)), 1)

    def test_old_three_field_rows_still_count_cash(self):
        self.assertEqual(cash_amount(['店', '1200', 'メモ']), 1200)
        self.assertEqual(normalize(['店', '1200', 'メモ'])[3:], ['', '', '', ''])

    def test_payment_methods_persist_and_delete_keeps_history(self):
        self.manager.add_payment_method('テストカード')
        self.manager.data = {'2026-7-1-1': [['店', '100', '', '0', 'テストカード']]}
        other = DataManager()
        other.load_settings()
        self.assertIn('テストカード', other.payment_methods)
        self.manager.remove_payment_method('テストカード')
        other.load_settings()
        self.assertNotIn('テストカード', other.payment_methods)
        self.assertEqual(self.manager.data['2026-7-1-1'][0][4], 'テストカード')
        with self.assertRaises(ValueError):
            self.manager.add_payment_method('  ')

    def test_points_are_deducted_and_full_points_payment_is_zero(self):
        self.assertEqual(cash_amount(validate(['店', '1,000', '商品', '200'])), 800)
        self.assertEqual(cash_amount(validate(['店', '1000', '商品', '1000'])), 0)
        self.assertEqual(cash_amount(validate(['店', '-100', '返金', '0'])), -100)
        with self.assertRaises(ValueError):
            validate(['店', '1000', '商品', '1001'])
        with self.assertRaises(ValueError):
            validate(['店', '1000', '商品', '-1'])


if __name__ == '__main__':
    unittest.main()
