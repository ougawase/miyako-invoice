"""
請求書生成コアロジック（app.py・テスト共用）
このファイルのみを変更すれば本番とテストの両方に反映される
"""

import io
import os
import openpyxl
from datetime import date, datetime

_TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), "templates")

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
    path = os.path.join(_TEMPLATES_DIR, "invoice_template.xlsx")
    with open(path, "rb") as f:
        return io.BytesIO(f.read())


def load_nouhin_template_bytes():
    """納品書テンプレートを bytes で返す（ダウンロードボタン用）。"""
    path = os.path.join(_TEMPLATES_DIR, "nouhin_template.xlsx")
    with open(path, "rb") as f:
        return f.read()


# ---- 納品書テンプレート用セル位置 ----
# nouhin_template.xlsx の実際のレイアウトに合わせた値
_N_STORE_ROW  = 3   # 店舗名: B3（merged B3:F3）
_N_STORE_COL  = 2
_N_DATE_ROW   = 4   # 納品日: B4
_N_DATE_COL   = 2
_N_ITEM_START = 10  # 明細開始行（9行目はヘッダー行）
_N_COL_NAME   = 2   # B列: 商品名
_N_COL_QTY    = 3   # C列: 数量
_N_COL_PRICE  = 4   # D列: 単価（利用者が手入力）


def create_nouhin_template():
    """
    納品書テンプレートExcelを生成して bytes で返す。

    入力セル（黄色）:
      B3: 店舗名（必須）
      B4: 納品日（必須）
      B7〜B22: 商品名（最大16行、空行で終了）
      C7〜C22: 数量（商品名と同じ行に入力）

    自動計算（入力不要）:
      E列: 単価（右の商品マスタから自動参照）
      F列: 金額（数量×単価）
      F24: 小計 / F25: 消費税 / F26: 合計（税込）

    右側参照（H・I列）: 商品マスタ（商品名と単価の一覧）
    """
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.worksheet.page import PageMargins

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "納品書"

    # ---- スタイル定義 ----
    def fill(color):
        return PatternFill("solid", fgColor=color)

    def border(color="BDBDBD", style="thin"):
        s = Side(style=style, color=color)
        return Border(left=s, right=s, top=s, bottom=s)

    def thick_border():
        t = Side(style="medium", color="555555")
        th = Side(style="thin", color="BDBDBD")
        return Border(left=t, right=t, top=t, bottom=th)

    YELLOW     = fill("FFF176")  # 入力セル（やや濃い黄色）
    NAVY       = fill("1A237E")  # 明細ヘッダー
    LIGHT_GRAY = fill("F5F5F5")  # 奇数行
    WHITE      = fill("FFFFFF")  # 偶数行
    BLUE_GRAY  = fill("ECEFF1")  # 小計エリア
    MASTER_HDR = fill("E8EAF6")  # 商品マスタヘッダー
    MASTER_ROW = fill("F3F4FE")  # 商品マスタ行
    AMBER      = fill("FFF8E1")  # ラベルセル背景

    def hdr_font(size=10, color="FFFFFF"):
        return Font(bold=True, size=size, color=color,
                    name="游ゴシック" if True else "Calibri")

    def body_font(size=10, color="212121", bold=False):
        return Font(bold=bold, size=size, color=color,
                    name="游ゴシック" if True else "Calibri")

    def note_font():
        return Font(italic=True, size=9, color="9E9E9E")

    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=False)
    LEFT   = Alignment(horizontal="left",   vertical="center")
    RIGHT  = Alignment(horizontal="right",  vertical="center")

    thin  = border()
    thick = border(style="medium", color="555555")

    # ---- 列幅 ----
    col_widths = {
        "A": 14,  # ラベル列（店舗名・納品日・No.）
        "B": 28,  # 商品名（入力）
        "C": 8,   # 数量（入力）
        "D": 6,   # 単位
        "E": 11,  # 単価（自動）
        "F": 13,  # 金額（自動）
        "G": 2,   # 区切り
        "H": 22,  # 商品一覧 商品名
        "I": 11,  # 商品一覧 単価
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    # ---- 行高 ----
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 6
    ws.row_dimensions[3].height = 24
    ws.row_dimensions[4].height = 22
    ws.row_dimensions[5].height = 8
    ws.row_dimensions[6].height = 20
    for r in range(_N_ITEM_START, _N_ITEM_END + 1):
        ws.row_dimensions[r].height = 18
    ws.row_dimensions[_N_ITEM_END + 1].height = 6
    ws.row_dimensions[_N_ITEM_END + 2].height = 18
    ws.row_dimensions[_N_ITEM_END + 3].height = 18
    ws.row_dimensions[_N_ITEM_END + 4].height = 20
    ws.row_dimensions[_N_ITEM_END + 6].height = 30

    # ---- タイトル行 ----
    ws.merge_cells("A1:F1")
    ws["A1"].value     = "納　品　書"
    ws["A1"].font      = Font(bold=True, size=18, color="1A237E", name="游ゴシック")
    ws["A1"].alignment = CENTER
    ws["A1"].fill      = fill("FAFAFA")
    ws["A1"].border    = Border(bottom=Side(style="medium", color="1A237E"))

    # ---- 店舗名（B3）----
    ws["A3"].value     = "納品先（店舗名）"
    ws["A3"].font      = body_font(9, bold=True, color="546E7A")
    ws["A3"].alignment = RIGHT
    ws["A3"].fill      = AMBER

    ws.merge_cells("B3:G3")
    ws["B3"].value       = _N_STORE_PLACEHOLDER
    ws["B3"].font        = Font(size=12, color="9E9E9E", italic=True, name="游ゴシック")
    ws["B3"].alignment   = LEFT
    ws["B3"].fill        = YELLOW
    ws["B3"].border      = thick

    ws["A3"].border = Border(
        left=Side(style="medium", color="555555"),
        top=Side(style="medium", color="555555"),
        bottom=Side(style="medium", color="555555"),
    )

    # ---- 納品日（B4）----
    ws["A4"].value     = "納　品　日"
    ws["A4"].font      = body_font(9, bold=True, color="546E7A")
    ws["A4"].alignment = RIGHT
    ws["A4"].fill      = AMBER

    ws["B4"].value       = date.today()
    ws["B4"].number_format = "YYYY年MM月DD日"
    ws["B4"].font        = body_font(11, bold=True)
    ws["B4"].alignment   = LEFT
    ws["B4"].fill        = YELLOW
    ws["B4"].border      = thin

    ws["C4"].value     = "← 日付をクリックして変更してください"
    ws["C4"].font      = note_font()
    ws["C4"].alignment = LEFT
    ws.merge_cells("C4:G4")

    ws["A4"].border = Border(
        left=Side(style="medium", color="555555"),
        bottom=Side(style="medium", color="555555"),
    )

    # ---- 明細ヘッダー（行6）----
    header_cols = {
        1: ("No.", 6),
        2: ("商　品　名\n（黄色セルに入力）", 10),
        3: ("数量\n（個）", 10),
        4: ("単位", 9),
        5: ("単価\n（自動）", 9),
        6: ("金　額\n（自動）", 9),
    }
    for col, (label, size) in header_cols.items():
        c = ws.cell(row=6, column=col)
        c.value     = label
        c.font      = hdr_font(size)
        c.fill      = NAVY
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = border(color="FFFFFF", style="thin")

    # ---- 明細行（7〜22）----
    pm_row_count = len(PRICE_MASTER)
    for i in range(_N_ITEM_END - _N_ITEM_START + 1):
        r       = _N_ITEM_START + i
        row_fill = LIGHT_GRAY if i % 2 == 0 else WHITE

        # No.
        ws.cell(row=r, column=1).value     = i + 1
        ws.cell(row=r, column=1).font      = body_font(9, color="9E9E9E")
        ws.cell(row=r, column=1).alignment = CENTER
        ws.cell(row=r, column=1).fill      = row_fill
        ws.cell(row=r, column=1).border    = thin

        # 商品名（入力）
        c_name = ws.cell(row=r, column=_N_COL_NAME)
        if i == 0:
            c_name.value  = "ツナプレーン"
            c_name.font   = body_font(10, color="1565C0")
        elif i == 1:
            c_name.value  = "ツナおかず"
            c_name.font   = body_font(10, color="1565C0")
        else:
            c_name.font = body_font(10)
        c_name.fill      = YELLOW
        c_name.alignment = LEFT
        c_name.border    = thin

        # 数量（入力）
        c_qty = ws.cell(row=r, column=_N_COL_QTY)
        if i == 0:
            c_qty.value = 3
            c_qty.font  = body_font(10, color="1565C0")
        elif i == 1:
            c_qty.value = 2
            c_qty.font  = body_font(10, color="1565C0")
        else:
            c_qty.font = body_font(10)
        c_qty.fill      = YELLOW
        c_qty.alignment = CENTER
        c_qty.border    = thin

        # 単位（固定）
        c_unit = ws.cell(row=r, column=4)
        c_unit.value     = "個"
        c_unit.font      = body_font(9, color="757575")
        c_unit.alignment = CENTER
        c_unit.fill      = row_fill
        c_unit.border    = thin

        # 単価（VLOOKUP で自動）
        pm_range = f"H${_N_PM_ROW_START}:I${_N_PM_ROW_START + pm_row_count - 1}"
        c_price = ws.cell(row=r, column=5)
        c_price.value       = f'=IFERROR(VLOOKUP(B{r},{pm_range},2,0),"")'
        c_price.number_format = "#,##0"
        c_price.font        = body_font(10, color="424242")
        c_price.alignment   = RIGHT
        c_price.fill        = row_fill
        c_price.border      = thin

        # 金額（自動計算）
        c_amt = ws.cell(row=r, column=6)
        c_amt.value         = f'=IF(OR(B{r}="",C{r}=""),"",C{r}*E{r})'
        c_amt.number_format = "#,##0"
        c_amt.font          = body_font(10, bold=True)
        c_amt.alignment     = RIGHT
        c_amt.fill          = row_fill
        c_amt.border        = thin

    # サンプル行に補足
    ws["B7"].comment = None  # コメントなし（Excelコメント依存はしない）

    # ---- 小計エリア ----
    total_start = _N_ITEM_END + 2

    def write_total_row(row, label, formula, bold=False):
        ws.merge_cells(f"A{row}:E{row}")
        c_lbl = ws.cell(row=row, column=1)
        c_lbl.value     = label
        c_lbl.font      = body_font(10, bold=bold)
        c_lbl.alignment = RIGHT
        c_lbl.fill      = BLUE_GRAY
        c_lbl.border    = thin

        c_val = ws.cell(row=row, column=6)
        c_val.value         = formula
        c_val.number_format = "#,##0"
        c_val.font          = body_font(11 if bold else 10, bold=bold)
        c_val.alignment     = RIGHT
        c_val.fill          = BLUE_GRAY if not bold else fill("E8F5E9")
        c_val.border        = thin

    write_total_row(total_start,     "小計（税抜）",
                    f"=SUM(F{_N_ITEM_START}:F{_N_ITEM_END})")
    write_total_row(total_start + 1, "消費税（8%）",
                    f"=INT(F{total_start}*0.08)")
    write_total_row(total_start + 2, "合　計（税込）",
                    f"=F{total_start}+F{total_start+1}", bold=True)

    # ---- 入力ガイド ----
    guide_row = total_start + 4
    ws.merge_cells(f"A{guide_row}:F{guide_row}")
    ws[f"A{guide_row}"].value = (
        "【入力ガイド】"
        "　①納品先に店舗名を入力　"
        "②商品名は右の「商品一覧」と同じ名前で入力　"
        "③数量を半角数字で入力　"
        "④空行になったら読み込み終了（最大16行）"
    )
    ws[f"A{guide_row}"].font      = note_font()
    ws[f"A{guide_row}"].alignment = LEFT
    ws[f"A{guide_row}"].fill      = fill("FFFDE7")
    ws[f"A{guide_row}"].border    = Border(
        left=Side(style="thin", color="FBC02D"),
        right=Side(style="thin", color="FBC02D"),
        top=Side(style="thin", color="FBC02D"),
        bottom=Side(style="thin", color="FBC02D"),
    )
    ws.row_dimensions[guide_row].height = 28

    # ---- 商品マスタ（H・I列）----
    # タイトル
    ws.merge_cells(f"H{_N_ITEM_START - 2}:I{_N_ITEM_START - 2}")
    c_mt = ws[f"H{_N_ITEM_START - 2}"]
    c_mt.value     = "商 品 一 覧"
    c_mt.font      = hdr_font(9, color="1A237E")
    c_mt.fill      = MASTER_HDR
    c_mt.alignment = CENTER
    c_mt.border    = thin

    # ヘッダー
    for col, label in ((_N_PM_COL_NAME, "商 品 名"), (_N_PM_COL_PRICE, "単価（円）")):
        c = ws.cell(row=_N_ITEM_START - 1, column=col)
        c.value     = label
        c.font      = hdr_font(9)
        c.fill      = NAVY
        c.alignment = CENTER
        c.border    = thin

    # データ
    for i, (name, price) in enumerate(PRICE_MASTER.items()):
        r = _N_PM_ROW_START + i
        cn = ws.cell(row=r, column=_N_PM_COL_NAME)
        cp = ws.cell(row=r, column=_N_PM_COL_PRICE)
        cn.value     = name
        cp.value     = price
        cn.font      = body_font(10)
        cp.font      = body_font(10)
        cn.fill      = cp.fill = MASTER_ROW if i % 2 == 0 else WHITE
        cn.alignment = LEFT
        cp.alignment = RIGHT
        cp.number_format = "#,##0"
        cn.border    = cp.border = thin

    # 注釈
    note_row = _N_PM_ROW_START + pm_row_count + 1
    ws.merge_cells(f"H{note_row}:I{note_row}")
    c_note = ws[f"H{note_row}"]
    c_note.value     = "※ 商品名は正確に入力してください"
    c_note.font      = note_font()
    c_note.alignment = LEFT

    # ---- ページ設定（A4縦・印刷範囲）----
    ws.page_setup.paperSize  = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToPage  = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(left=0.5, right=0.3, top=0.6, bottom=0.6)
    ws.print_area   = f"A1:I{guide_row}"

    # シート保護（入力セルのみ編集可にする）
    from openpyxl.styles import Protection
    # まず全セルをロック状態に
    # 入力セル（黄色）だけ解除
    for r in range(1, guide_row + 2):
        for col in range(1, 10):
            ws.cell(row=r, column=col).protection = Protection(locked=True)

    # 解除: 店舗名・納品日
    ws.cell(row=_N_STORE_ROW, column=_N_STORE_COL).protection = Protection(locked=False)
    ws.cell(row=_N_DATE_ROW,  column=_N_DATE_COL).protection  = Protection(locked=False)

    # 解除: 明細入力列
    for r in range(_N_ITEM_START, _N_ITEM_END + 1):
        ws.cell(row=r, column=_N_COL_NAME).protection = Protection(locked=False)
        ws.cell(row=r, column=_N_COL_QTY).protection  = Protection(locked=False)

    ws.protection.sheet   = True
    ws.protection.password = ""   # パスワードなし（Alt+R→S で解除可）
    ws.protection.enable()

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def read_nouhinshо(file_bytes):
    """
    納品書Excelを読み取り、店舗名・日付・明細を返す。

    読み取り位置:
      B2: 店舗名（必須）
      B3: 納品日（必須）
      B9〜: 商品名, C9〜: 数量, D9〜: 単価（利用者が手入力）
      空行になった時点で読み込み終了

    戻り値: dict（成功）or {"error": str}（フォーマット不正）
    """
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception:
        return {"error": "Excelファイルとして開けませんでした。テンプレートを使用してください。"}

    ws = wb.active

    # ── 店舗名（B2）──
    raw_store  = ws.cell(row=_N_STORE_ROW, column=_N_STORE_COL).value
    store_name = str(raw_store).strip() if raw_store is not None else ""
    if not store_name:
        return {"error": "B3セルに店舗名が入力されていません。"}

    # ── 納品日（B3）──
    delivery_date = ws.cell(row=_N_DATE_ROW, column=_N_DATE_COL).value
    warnings = []
    if delivery_date is None:
        warnings.append("B4セルに納品日が入力されていません。日付を入力することを推奨します。")

    # ── 明細（B・C・D列、行 9〜）──
    items = []
    r = _N_ITEM_START
    while True:
        raw_name  = ws.cell(row=r, column=_N_COL_NAME).value
        raw_qty   = ws.cell(row=r, column=_N_COL_QTY).value
        raw_price = ws.cell(row=r, column=_N_COL_PRICE).value

        # 商品名が空になったら終了
        if raw_name is None or str(raw_name).strip() == "":
            break

        name_str = str(raw_name).strip()

        # 数量チェック
        try:
            qty_int = int(raw_qty)
            if qty_int <= 0:
                raise ValueError
        except (ValueError, TypeError):
            return {"error": f"行{r}の数量「{raw_qty}」が正しくありません。半角数字で入力してください。"}

        # 単価チェック
        try:
            price_int = int(raw_price)
            if price_int < 0:
                raise ValueError
        except (ValueError, TypeError):
            return {"error": f"行{r}の単価「{raw_price}」が正しくありません。半角数字で入力してください。"}

        items.append({
            "name":       name_str,
            "unit_price": price_int,
            "quantity":   qty_int,
            "amount":     qty_int * price_int,
        })
        r += 1

    if not items:
        return {
            "error": (
                "明細が読み取れませんでした。"
                "10行目以降のB列（商品名）・C列（数量）・D列（単価）を入力してください。"
            )
        }

    subtotal = sum(i["amount"] for i in items)
    return {
        "store_name":    store_name,
        "delivery_date": delivery_date,
        "items":         items,
        "subtotal":      subtotal,
        "tax":           int(subtotal * TAX_RATE),
        "total":         subtotal + int(subtotal * TAX_RATE),
        "warnings":      warnings,
    }


def _flatten_items(dated_groups):
    """dated_groups をフラットリストに展開する（マージなし・日付プレフィックスなし）。"""
    rows = []
    for group in dated_groups:
        for item in group["items"]:
            rows.append({
                "name":       item["name"],
                "quantity":   item["quantity"],
                "unit_price": item["unit_price"],
                "amount":     item["amount"],
            })
    return rows


def create_invoice(store_name, dated_groups, billing_month, billing_subject,
                   tax_rate=0.08, payment_due="", payment_terms="", valid_until="",
                   delivery_date=None):
    """
    テンプレートから請求書Excelを生成して bytes で返す。

    A列(No.): 連番 1, 2, 3 ...
    B列(摘要): "M/D 商品名"（日付プレフィックス付き）
    J列: 数量, K列: 単位, L列: 単価, O列: 金額

    tax_rate      : 消費税率（0.08 or 0.10）
    payment_due   : 納期（任意）→ C10 に書き込む
    payment_terms : 支払条件（任意）→ C11
    valid_until   : 有効期限（任意）→ C12
    """
    wb = openpyxl.load_workbook(load_template())
    ws = wb.active
    wb.calculation.fullCalcOnLoad = True

    # ── ヘッダー情報 ──
    ws["A3"] = store_name
    ws["N4"] = date.today()
    reiwa    = date.today().year - 2018
    ws["C6"] = f"令和{reiwa}年{billing_month}{billing_subject}"

    # ── 税率（T6 = 0.08 or 0.10）と消費税ラベル（J35）──
    ws.cell(row=6, column=20).value  = tax_rate                      # T6
    ws.cell(row=35, column=10).value = f"消費税（{int(tax_rate*100)}%）"  # J35

    # ── 任意フィールド ──
    if payment_due:
        ws["C10"] = payment_due
    elif delivery_date is not None:
        # 納期欄に納品日を自動セット（手動入力がない場合のみ）
        if isinstance(delivery_date, (date, datetime)):
            ws["C10"] = f"{delivery_date.month}月{delivery_date.day}日"
        else:
            ws["C10"] = str(delivery_date)
    if payment_terms:
        ws["C11"] = payment_terms
    if valid_until:
        ws["C12"] = valid_until

    # ── 明細を全クリア ──
    for r in range(ITEM_START, ITEM_END + 1):
        ws.cell(row=r, column=1).value  = None   # A: No.
        ws.cell(row=r, column=2).value  = None   # B: 摘要
        ws.cell(row=r, column=10).value = None   # J: 数量
        ws.cell(row=r, column=11).value = None   # K: 単位
        ws.cell(row=r, column=12).value = None   # L: 単価
        ws.cell(row=r, column=15).value = None   # O: 金額

    # ── 明細を書き込む（マージなし・連番）──
    flat_items  = _flatten_items(dated_groups)
    current_row = ITEM_START
    for seq, row in enumerate(flat_items, start=1):
        if current_row > ITEM_END:
            break
        ws.cell(row=current_row, column=1).value  = seq
        ws.cell(row=current_row, column=2).value  = row["name"]
        ws.cell(row=current_row, column=10).value = row["quantity"]
        ws.cell(row=current_row, column=11).value = "個"
        ws.cell(row=current_row, column=12).value = row["unit_price"]
        ws.cell(row=current_row, column=15).value = row["amount"]
        current_row += 1

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def verify_invoice_output(excel_bytes, dated_groups):
    """
    生成した Excel を読み直し、書き込んだ値が正しいか検証する。
    create_invoice と同じ _flatten_items ロジックで期待値を再構築して比較する。
    """
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    ws = wb.active
    errors = []

    expected_rows = _flatten_items(dated_groups)

    for i, exp in enumerate(expected_rows):
        r = ITEM_START + i
        if r > ITEM_END:
            errors.append(f"行数が上限（{ITEM_END - ITEM_START + 1}行）を超えています")
            break

        actual_seq   = ws.cell(row=r, column=1).value
        actual_name  = ws.cell(row=r, column=2).value
        actual_qty   = ws.cell(row=r, column=10).value
        actual_price = ws.cell(row=r, column=12).value
        actual_amt   = ws.cell(row=r, column=15).value

        expected_seq  = i + 1
        expected_name = exp["name"]

        if actual_seq != expected_seq:
            errors.append(f"行{r} No.: 期待{expected_seq} 実際{actual_seq}")
        if actual_name != expected_name:
            errors.append(f"行{r} 摘要: 期待「{expected_name}」実際「{actual_name}」")
        if actual_qty != exp["quantity"]:
            errors.append(f"行{r} 数量: 期待{exp['quantity']} 実際{actual_qty}")
        if actual_price != exp["unit_price"]:
            errors.append(f"行{r} 単価: 期待¥{exp['unit_price']} 実際¥{actual_price}")
        if actual_amt != exp["amount"]:
            errors.append(f"行{r} 金額: 期待¥{exp['amount']} 実際¥{actual_amt}")

    subtotal = sum(
        ws.cell(row=ITEM_START + i, column=15).value or 0
        for i in range(len(expected_rows))
        if ITEM_START + i <= ITEM_END
    )

    return {"ok": len(errors) == 0, "errors": errors, "subtotal": subtotal}
