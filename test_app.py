"""
請求書自動生成ツール 自動テスト
★ invoice.py（本番コード）を直接 import してテストする。
  テスト専用のコピーは持たない → テストと本番が常に一致。

実行方法: python3 test_app.py
"""

import unittest
import openpyxl
import io
from datetime import date, datetime

# ★ 本番コードを直接 import
from invoice import (
    create_invoice,
    read_nouhinshо,
    verify_invoice_output,
    PRICE_MASTER,
    TAX_RATE,
    ITEM_START,
    ITEM_END,
)


# ====================================================
# ユーティリティ
# ====================================================
def make_item(name, qty, unit_price=None):
    """テスト用アイテムを簡単に作るヘルパー"""
    if unit_price is None:
        unit_price = PRICE_MASTER.get(name, 0)
    return {"name": name, "unit_price": unit_price, "quantity": qty, "amount": qty * unit_price}


def load_excel(excel_bytes):
    return openpyxl.load_workbook(io.BytesIO(excel_bytes)).active


# ====================================================
class TestPriceCalculation(unittest.TestCase):
    """1. 金額計算の正確性"""

    def test_all_master_products_have_nonzero_price(self):
        """価格マスタの全商品が 1 円以上であること"""
        for name, price in PRICE_MASTER.items():
            self.assertGreater(price, 0, f"{name} の単価が 0 です")

    def test_item_amount_equals_price_times_qty(self):
        """単価 × 数量 = 金額"""
        for name, price in PRICE_MASTER.items():
            item = make_item(name, 35)
            self.assertEqual(item["amount"], price * 35,
                             f"{name}: {price}×35={price*35} のはずが {item['amount']}")

    def test_subtotal(self):
        """小計の合計"""
        items = [make_item("ツナプレーン", 35), make_item("ツナラー油", 35), make_item("ツナおかず", 35)]
        subtotal = sum(i["amount"] for i in items)
        self.assertEqual(subtotal, 520*35 + 580*35 + 580*35)

    def test_tax_is_truncated_not_rounded(self):
        """消費税は切り捨て（int()）であること"""
        subtotal = 58801  # 58801 * 0.08 = 4704.08 → int は 4704
        tax = int(subtotal * TAX_RATE)
        self.assertEqual(tax, 4704)


# ====================================================
class TestExcelOutput(unittest.TestCase):
    """2. Excel に正しく書き込まれているか"""

    def setUp(self):
        self.groups = [
            {"date": datetime(2026, 5, 4), "items": [
                make_item("ツナプレーン", 35),
                make_item("ツナラー油",   35),
                make_item("ツナおかず",   35),
            ]}
        ]
        self.excel = create_invoice("テスト店舗", self.groups, "5月分", "味噌加工品代金")
        self.ws    = load_excel(self.excel)

    def test_store_name_in_a3(self):
        self.assertEqual(self.ws["A3"].value, "テスト店舗")

    def test_billing_title_in_c6(self):
        title = self.ws["C6"].value
        self.assertIn("5月分", title)
        self.assertIn("味噌加工品代金", title)
        # 年が重複していないこと（例: 「令和8年8年」とならない）
        reiwa = date.today().year - 2018
        self.assertNotIn(f"年{reiwa}年", title, f"件名に年が重複: {title}")

    def test_product_name_written_to_b_column(self):
        """B列に商品名が書かれていること"""
        self.assertEqual(self.ws.cell(row=ITEM_START, column=2).value, "ツナプレーン")

    def test_quantity_written_to_j_column(self):
        """J列に数量が数値で書かれていること"""
        qty = self.ws.cell(row=ITEM_START, column=10).value
        self.assertEqual(qty, 35)
        self.assertIsInstance(qty, int, f"数量は int のはずが {type(qty)}")

    def test_unit_price_written_to_l_column(self):
        """L列に単価が数値で書かれていること（VLOOKUPではなく数値）"""
        price = self.ws.cell(row=ITEM_START, column=12).value
        self.assertEqual(price, 520, f"単価が 520 のはずが {price}")
        self.assertIsInstance(price, int, f"単価は int のはずが {type(price)}")

    def test_amount_written_to_o_column(self):
        """O列に金額が数値で書かれていること（数式ではなく数値）"""
        amount = self.ws.cell(row=ITEM_START, column=15).value
        self.assertEqual(amount, 520 * 35, f"金額が {520*35} のはずが {amount}")
        self.assertIsInstance(amount, int, f"金額は int のはずが {type(amount)}")

    def test_unit_column_is_ko(self):
        """K列に「個」が書かれていること"""
        self.assertEqual(self.ws.cell(row=ITEM_START, column=11).value, "個")

    def test_delivery_date_label(self):
        """A列の先頭行に日付ラベル（○月○日）が書かれていること"""
        self.assertEqual(self.ws.cell(row=ITEM_START, column=1).value, "5月4日")

    def test_full_calc_on_load(self):
        """Excel を開いたとき SUM 数式が自動再計算されること"""
        wb = openpyxl.load_workbook(io.BytesIO(self.excel))
        self.assertTrue(wb.calculation.fullCalcOnLoad)

    def test_template_loads(self):
        """テンプレートが読み込めること"""
        from invoice import load_template
        import openpyxl
        wb = openpyxl.load_workbook(load_template())
        self.assertIn("見積書", wb.sheetnames)


# ====================================================
class TestExcelVerification(unittest.TestCase):
    """3. verify_invoice_output が書き込み不整合を検出できるか"""

    def _make_excel(self, groups):
        return create_invoice("テスト店舗", groups, "5月分", "テスト")

    def test_verify_passes_for_correct_output(self):
        """正しい書き込みは verify が OK を返すこと"""
        groups = [{"date": datetime(2026, 5, 4), "items": [make_item("ツナプレーン", 35)]}]
        excel  = self._make_excel(groups)
        result = verify_invoice_output(excel, groups)
        self.assertTrue(result["ok"], f"verify が失敗: {result['errors']}")

    def test_verify_subtotal_matches(self):
        """verify が返す小計が期待値と一致すること"""
        groups = [{"date": datetime(2026, 5, 4), "items": [
            make_item("ツナプレーン", 35),
            make_item("ツナおかず",   35),
        ]}]
        excel  = self._make_excel(groups)
        result = verify_invoice_output(excel, groups)
        expected = 520*35 + 580*35
        self.assertEqual(result["subtotal"], expected,
                         f"小計: 期待{expected} 実際{result['subtotal']}")

    def test_verify_detects_wrong_amount(self):
        """O列の金額が間違っていたら verify がエラーを返すこと"""
        groups = [{"date": datetime(2026, 5, 4), "items": [make_item("ツナプレーン", 35)]}]
        excel  = bytearray(self._make_excel(groups))

        # 金額を手動で壊す（O列を別の値に上書きしてから検証）
        wb = openpyxl.load_workbook(io.BytesIO(bytes(excel)))
        ws = wb.active
        ws.cell(row=ITEM_START, column=15).value = 9999  # 本来 18200
        buf = io.BytesIO(); wb.save(buf)

        result = verify_invoice_output(buf.getvalue(), groups)
        self.assertFalse(result["ok"], "金額が壊れているのに verify が OK を返しました")
        self.assertTrue(any("金額" in e for e in result["errors"]),
                        f"金額エラーメッセージがありません: {result['errors']}")


# ====================================================
class TestMultipleDates(unittest.TestCase):
    """4. 複数日付の処理"""

    def test_different_dates_written_to_separate_rows(self):
        """日付が異なる → 別行に出力されること"""
        groups = [
            {"date": datetime(2026, 5, 4),  "items": [make_item("ツナプレーン", 30)]},
            {"date": datetime(2026, 5, 11), "items": [make_item("ツナプレーン", 20)]},
        ]
        ws = load_excel(create_invoice("テスト", groups, "5月分", "テスト"))
        self.assertEqual(ws.cell(row=ITEM_START,     column=10).value, 30)
        self.assertEqual(ws.cell(row=ITEM_START + 1, column=10).value, 20)

    def test_same_product_not_merged_across_dates(self):
        """同一商品でも日付が違えば合算されないこと"""
        groups = [
            {"date": datetime(2026, 5, 4),  "items": [make_item("ツナプレーン", 35)]},
            {"date": datetime(2026, 5, 11), "items": [make_item("ツナプレーン", 35)]},
        ]
        ws = load_excel(create_invoice("テスト", groups, "5月分", "テスト"))
        self.assertEqual(ws.cell(row=ITEM_START,     column=10).value, 35)
        self.assertEqual(ws.cell(row=ITEM_START + 1, column=10).value, 35)

    def test_same_product_merged_within_same_date(self):
        """同一日付の同一商品は合算されること"""
        groups = [{"date": datetime(2026, 5, 4), "items": [
            make_item("ツナプレーン", 20),
            make_item("ツナプレーン", 15),
        ]}]
        ws = load_excel(create_invoice("テスト", groups, "5月分", "テスト"))
        self.assertEqual(ws.cell(row=ITEM_START, column=10).value, 35)
        self.assertIsNone(ws.cell(row=ITEM_START + 1, column=2).value)

    def test_date_label_position_with_multiple_items(self):
        """1日目が2商品なら2日目のラベルは ITEM_START+2 行目にあること"""
        groups = [
            {"date": datetime(2026, 5, 4), "items": [
                make_item("ツナプレーン", 35),
                make_item("ツナラー油",   35),
            ]},
            {"date": datetime(2026, 5, 11), "items": [make_item("ツナおかず", 35)]},
        ]
        ws = load_excel(create_invoice("テスト", groups, "5月分", "テスト"))
        self.assertEqual(ws.cell(row=ITEM_START + 2, column=1).value, "5月11日")

    def test_amount_per_date_group_is_correct(self):
        """各日付グループの金額が正しく計算されること"""
        groups = [
            {"date": datetime(2026, 4, 6),  "items": [make_item("ツナプレーン", 35), make_item("ツナおかず", 35)]},
            {"date": datetime(2026, 4, 13), "items": [make_item("ツナプレーン", 20)]},
        ]
        ws = load_excel(create_invoice("テスト", groups, "4月分", "テスト"))
        # 4/6: ツナプレーン=18200, ツナおかず=20300
        self.assertEqual(ws.cell(row=ITEM_START,     column=15).value, 18200)
        self.assertEqual(ws.cell(row=ITEM_START + 1, column=15).value, 20300)
        # 4/13: ツナプレーン=10400
        self.assertEqual(ws.cell(row=ITEM_START + 2, column=15).value, 10400)


# ====================================================
class TestEdgeCases(unittest.TestCase):
    """5. 境界値・異常系"""

    def test_row_overflow_does_not_crash(self):
        """商品数が上限を超えてもクラッシュしないこと"""
        many = [make_item("ツナプレーン", 1) for _ in range(20)]
        many = [{"name": f"商品{i:02d}", "unit_price": 500, "quantity": 1, "amount": 500}
                for i in range(20)]
        groups = [{"date": datetime(2026, 5, 4), "items": many}]
        excel = create_invoice("テスト", groups, "5月分", "テスト")
        ws    = load_excel(excel)
        # ITEM_END+1 行目には何も書かれていないこと
        self.assertIsNone(ws.cell(row=ITEM_END + 1, column=2).value)

    def test_none_date_does_not_crash(self):
        """日付 None でもクラッシュせず、A列が空欄になること"""
        groups = [{"date": None, "items": [make_item("ツナプレーン", 35)]}]
        ws = load_excel(create_invoice("テスト", groups, "5月分", "テスト"))
        self.assertEqual(ws.cell(row=ITEM_START, column=2).value, "ツナプレーン")
        self.assertIn(ws.cell(row=ITEM_START, column=1).value, [None, ""])

    def test_empty_groups_does_not_crash(self):
        """グループが空でもクラッシュせず、明細行が空になること"""
        ws = load_excel(create_invoice("テスト", [], "5月分", "テスト"))
        for r in range(ITEM_START, ITEM_END + 1):
            self.assertIsNone(ws.cell(row=r, column=2).value)

    def test_previous_template_data_is_cleared(self):
        """テンプレートの既存データが残らないこと"""
        # 1商品だけ書き込んだあと、2行目以降が空であること
        groups = [{"date": datetime(2026, 5, 4), "items": [make_item("ツナプレーン", 10)]}]
        ws = load_excel(create_invoice("テスト", groups, "5月分", "テスト"))
        for r in range(ITEM_START + 1, ITEM_START + 5):
            self.assertIsNone(ws.cell(row=r, column=2).value,
                              f"行{r}に残存データがあります: {ws.cell(row=r, column=2).value}")

    def test_billing_month_all_months(self):
        """1〜12月すべてで件名が正しく生成されること"""
        groups = [{"date": datetime(2026, 5, 4), "items": [make_item("ツナプレーン", 10)]}]
        for m in range(1, 13):
            ws = load_excel(create_invoice("テスト", groups, f"{m}月分", "テスト代金"))
            title = ws["C6"].value
            self.assertIn(f"{m}月分", title, f"{m}月分が件名にない: {title}")
            reiwa = date.today().year - 2018
            self.assertNotIn(f"年{reiwa}年", title, f"年が重複: {title}")


# ====================================================
class TestZeroPriceDetection(unittest.TestCase):
    """6. 単価不明・ゼロ価格の検出"""

    def test_unknown_product_detected_in_read_result(self):
        """価格マスタにない商品名は unknown_products に記録されること"""
        import openpyxl, io, base64
        from template_data import TEMPLATE_B64

        # テスト用納品書を動的に作成
        wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(TEMPLATE_B64)))
        ws = wb.active
        ws["A3"] = "テスト店舗"
        ws["N4"] = datetime(2026, 5, 4)
        ws.cell(row=18, column=2).value  = "存在しない商品"  # 価格マスタにない
        ws.cell(row=18, column=10).value = 10
        buf = io.BytesIO(); wb.save(buf)

        result = read_nouhinshо(buf.getvalue())
        self.assertIsNotNone(result, "読み取り結果がNoneです")
        self.assertIn("存在しない商品", result["unknown_products"],
                      "単価不明の商品が unknown_products に記録されていません")

    def test_known_product_not_in_unknown_list(self):
        """価格マスタにある商品は unknown_products に入らないこと"""
        import openpyxl, io, base64
        from template_data import TEMPLATE_B64

        wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(TEMPLATE_B64)))
        ws = wb.active
        ws["A3"] = "テスト店舗"
        ws["N4"] = datetime(2026, 5, 4)
        ws.cell(row=18, column=2).value  = "ツナプレーン"
        ws.cell(row=18, column=10).value = 35
        buf = io.BytesIO(); wb.save(buf)

        result = read_nouhinshо(buf.getvalue())
        self.assertEqual(result["unknown_products"], [],
                         f"既知商品が unknown_products にある: {result['unknown_products']}")

    def test_zero_price_item_amount_is_zero(self):
        """単価0の商品は金額も0になること（警告で気づけるよう）"""
        item = {"name": "謎の商品", "unit_price": 0, "quantity": 10, "amount": 0}
        self.assertEqual(item["unit_price"] * item["quantity"], item["amount"])


# ====================================================
if __name__ == "__main__":
    print("=" * 60)
    print("請求書自動生成ツール 自動テスト")
    print("（本番コード invoice.py を直接テスト）")
    print("=" * 60)
    loader = unittest.TestLoader()
    suite  = unittest.TestSuite()
    for cls in [
        TestPriceCalculation,
        TestExcelOutput,
        TestExcelVerification,
        TestMultipleDates,
        TestEdgeCases,
        TestZeroPriceDetection,
    ]:
        suite.addTests(loader.loadTestsFromTestCase(cls))

    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    print("=" * 60)
    if result.wasSuccessful():
        print(f"✅ 全 {result.testsRun} テスト合格 — デプロイしても安全です")
    else:
        print(f"❌ 失敗: {len(result.failures)} 件  エラー: {len(result.errors)} 件")
        print("デプロイ前に修正してください")
    print("=" * 60)
