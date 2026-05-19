"""
請求書生成コアロジック（app.py・テスト共用）
このファイルのみを変更すれば本番とテストの両方に反映される
"""

import io
import base64
import openpyxl
from datetime import date, datetime
from template_data import TEMPLATE_B64

# ---- 定数 ----
PRICE_MASTER = {
    "ツナプレーン": 520,
    "ツナおかず":   580,
    "ツナラー油":   580,
    "ツナ味噌":     580,
    "ツナガリ":     580,
    "ツナカレー":   580,
}
TAX_RATE     = 0.08
ITEM_START   = 18   # 明細開始行
ITEM_END     = 33   # 明細終了行


def load_template():
    return io.BytesIO(base64.b64decode(TEMPLATE_B64))


def read_nouhinshо(file_bytes):
    """
    納品書Excelを読み取り、店舗名・日付・明細を返す。
    単価が不明な商品（価格マスタ未登録）は unknown_products に記録。
    戻り値が None の場合は読み取り失敗（店舗名なし or 明細なし）。
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active

    store_name    = ws["A3"].value or ""
    delivery_date = ws["N4"].value

    # 納品書内の価格マスタ（V6:W11）を読む（本体のPRICE_MASTERより優先）
    price_map = {}
    for row in ws.iter_rows(min_row=6, max_row=11, min_col=22, max_col=23, values_only=True):
        if row[0] and row[1]:
            price_map[str(row[0])] = int(row[1])
    for k, v in PRICE_MASTER.items():
        price_map.setdefault(k, v)

    items           = []
    unknown_products = []  # 単価が0の商品（価格マスタ未登録）

    for r in range(ITEM_START, ITEM_END + 1):  # 18〜33行（最大16商品）
        name = ws.cell(row=r, column=2).value
        qty  = ws.cell(row=r, column=10).value
        if not name or not qty:
            continue
        try:
            qty_int = int(qty)
        except (ValueError, TypeError):
            continue
        if qty_int <= 0:
            continue

        unit_price = price_map.get(str(name), 0)
        if unit_price == 0:
            unknown_products.append(str(name))  # 単価不明として記録

        items.append({
            "name":       str(name),
            "unit_price": unit_price,
            "quantity":   qty_int,
            "amount":     qty_int * unit_price,
        })

    if not store_name.strip() or not items:
        return None

    subtotal = sum(i["amount"] for i in items)
    return {
        "store_name":      store_name.strip(),
        "delivery_date":   delivery_date,
        "items":           items,
        "subtotal":        subtotal,
        "tax":             int(subtotal * TAX_RATE),
        "total":           subtotal + int(subtotal * TAX_RATE),
        "unknown_products": unknown_products,  # [] なら全商品の単価が判明
    }


def create_invoice(store_name, dated_groups, billing_month, billing_subject):
    """
    テンプレートから請求書Excelを生成して bytes で返す。

    dated_groups: [{"date": datetime, "items": [...]}, ...]
    各 item は {"name": str, "unit_price": int, "quantity": int, "amount": int}

    L列（単価）・O列（金額）を数値で直接書き込むことで、
    Excelの VLOOKUP / 数式キャッシュに依存せず合計が確実に出る。
    """
    wb = openpyxl.load_workbook(load_template())
    ws = wb.active

    # Excel を開いたときに数式を再計算（合計 SUM 数式が確実に動く）
    wb.calculation.fullCalcOnLoad = True

    # 宛先・発行日・件名
    ws["A3"] = store_name
    ws["N4"] = date.today()
    reiwa    = date.today().year - 2018
    ws["C6"] = f"令和{reiwa}年{billing_month}{billing_subject}"

    # 明細を全クリア（A: 日付, B: 商品名, J: 数量, K: 単位, L: 単価, O: 金額）
    for r in range(ITEM_START, ITEM_END + 1):
        ws.cell(row=r, column=1).value  = None
        ws.cell(row=r, column=2).value  = None
        ws.cell(row=r, column=10).value = None
        ws.cell(row=r, column=11).value = None
        ws.cell(row=r, column=12).value = None  # L: 単価（VLOOKUP を数値で上書き）
        ws.cell(row=r, column=15).value = None  # O: 金額（数式を数値で上書き）

    # 日付グループごとに書き込む
    current_row = ITEM_START
    for group in dated_groups:
        if current_row > ITEM_END:
            break

        d = group["date"]
        date_label = f"{d.month}月{d.day}日" if isinstance(d, date) else ""

        # 同じ日付内の同一商品を合算
        merged = {}
        for item in group["items"]:
            k = item["name"]
            if k in merged:
                merged[k]["quantity"] += item["quantity"]
                merged[k]["amount"]   += item["amount"]
            else:
                merged[k] = dict(item)

        first_in_group = True
        for item in merged.values():
            if current_row > ITEM_END:
                break
            if first_in_group:
                ws.cell(row=current_row, column=1).value = date_label
                first_in_group = False
            ws.cell(row=current_row, column=2).value  = item["name"]
            ws.cell(row=current_row, column=10).value = item["quantity"]
            ws.cell(row=current_row, column=11).value = "個"
            ws.cell(row=current_row, column=12).value = item["unit_price"]  # L: 単価（数値）
            ws.cell(row=current_row, column=15).value = item["amount"]      # O: 金額（数値）
            current_row += 1

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def verify_invoice_output(excel_bytes, dated_groups):
    """
    生成した Excel を読み直し、書き込んだ値が正しいか検証する。
    問題があれば {"ok": False, "errors": [...]} を返す。
    問題なければ {"ok": True, "errors": [], "subtotal": int} を返す。

    ★ これを呼ぶことで「書いたつもりが書けてなかった」事故を防ぐ。
    """
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    ws = wb.active
    errors = []

    # 期待値を再構築（create_invoice と同じロジック）
    expected_rows = []
    for group in dated_groups:
        d = group["date"]
        date_label = f"{d.month}月{d.day}日" if isinstance(d, date) else ""
        merged = {}
        for item in group["items"]:
            k = item["name"]
            if k in merged:
                merged[k]["quantity"] += item["quantity"]
                merged[k]["amount"]   += item["amount"]
            else:
                merged[k] = dict(item)
        first = True
        for item in merged.values():
            expected_rows.append({
                "date_label": date_label if first else "",
                "name":       item["name"],
                "quantity":   item["quantity"],
                "unit_price": item["unit_price"],
                "amount":     item["amount"],
            })
            first = False

    # Excel の実際の値と比較
    for i, exp in enumerate(expected_rows):
        r = ITEM_START + i
        if r > ITEM_END:
            errors.append(f"行数が上限({ITEM_END - ITEM_START + 1}行)を超えています")
            break

        actual_name  = ws.cell(row=r, column=2).value
        actual_qty   = ws.cell(row=r, column=10).value
        actual_price = ws.cell(row=r, column=12).value
        actual_amt   = ws.cell(row=r, column=15).value

        if actual_name != exp["name"]:
            errors.append(f"行{r} 商品名: 期待「{exp['name']}」実際「{actual_name}」")
        if actual_qty != exp["quantity"]:
            errors.append(f"行{r} 数量: 期待{exp['quantity']} 実際{actual_qty}")
        if actual_price != exp["unit_price"]:
            errors.append(f"行{r} 単価: 期待¥{exp['unit_price']} 実際¥{actual_price}")
        if actual_amt != exp["amount"]:
            errors.append(f"行{r} 金額: 期待¥{exp['amount']} 実際¥{actual_amt}")

    # 期待小計（数値で書いたO列の合計）
    subtotal = sum(
        ws.cell(row=ITEM_START + i, column=15).value or 0
        for i in range(len(expected_rows))
        if ITEM_START + i <= ITEM_END
    )

    return {"ok": len(errors) == 0, "errors": errors, "subtotal": subtotal}
