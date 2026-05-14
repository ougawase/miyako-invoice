import streamlit as st
import openpyxl
import pandas as pd
import io
import zipfile
import base64
from datetime import date, datetime
from template_data import TEMPLATE_B64

def load_template():
    return io.BytesIO(base64.b64decode(TEMPLATE_B64))

# 商品マスタ（テンプレートV6:W11より）
PRICE_MASTER = {
    "ツナプレーン": 520,
    "ツナおかず": 580,
    "ツナラー油": 580,
    "ツナ味噌": 580,
    "ツナガリ": 580,
    "ツナカレー": 580,
}
TAX_RATE = 0.08
ITEM_START_ROW = 18
ITEM_END_ROW = 33

st.set_page_config(page_title="新家 請求書自動生成", page_icon="📄", layout="centered")
st.title("📄 納品書 → 請求書 自動生成")
st.caption("株式会社 新家｜納品書をアップロードするだけで請求書を自動作成")
st.divider()


def read_nouhinshо(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    ws = wb.active

    store_name = ws["A3"].value or ""
    delivery_date = ws["N4"].value

    price_map = {}
    for row in ws.iter_rows(min_row=6, max_row=11, min_col=22, max_col=23, values_only=True):
        if row[0] and row[1]:
            price_map[str(row[0])] = int(row[1])
    for k, v in PRICE_MASTER.items():
        price_map.setdefault(k, v)

    items = []
    for r in range(18, 30):
        name = ws.cell(row=r, column=2).value
        qty  = ws.cell(row=r, column=10).value
        if name and qty and int(qty) > 0:
            unit_price = price_map.get(str(name), 0)
            items.append({
                "name": str(name),
                "unit_price": unit_price,
                "quantity": int(qty),
                "amount": int(qty) * unit_price,
            })

    if not store_name.strip() or not items:
        return None

    subtotal = sum(i["amount"] for i in items)
    return {
        "store_name": store_name.strip(),
        "delivery_date": delivery_date,
        "items": items,
        "subtotal": subtotal,
        "tax": int(subtotal * TAX_RATE),
        "total": subtotal + int(subtotal * TAX_RATE),
    }


def create_invoice(store_name, all_items, billing_month, delivery_dates):
    """お父さんのテンプレートをコピーして数字だけ書き換える"""
    wb = openpyxl.load_workbook(load_template())
    ws = wb.active

    # 宛先・日付・件名
    ws["A3"] = store_name
    ws["N4"] = date.today()

    # 件名（例: 令和8年5月分味噌加工品代金）
    reiwa = date.today().year - 2018
    ws["C6"] = f"令和{reiwa}年{billing_month}味噌加工品代金"

    # 明細を一度クリア（行18〜33のA・B・J列）
    for r in range(ITEM_START_ROW, ITEM_END_ROW + 1):
        ws.cell(row=r, column=1).value = None  # A: 日付
        ws.cell(row=r, column=2).value = None  # B: 商品名
        ws.cell(row=r, column=10).value = None # J: 数量

    # 商品をまとめる（同じ商品は合算）
    merged = {}
    for item in all_items:
        k = item["name"]
        if k in merged:
            merged[k]["quantity"] += item["quantity"]
            merged[k]["amount"] += item["amount"]
        else:
            merged[k] = dict(item)

    # 日付文字列を作る（例: "5月4日"）
    date_label = ""
    if delivery_dates:
        d = delivery_dates[0]
        if isinstance(d, datetime):
            date_label = f"{d.month}月{d.day}日"

    # 明細を書き込む（B列=商品名, J列=数量 → L列とO列の数式が自動計算）
    for i, item in enumerate(merged.values()):
        r = ITEM_START_ROW + i
        if r > ITEM_END_ROW:
            break
        if i == 0:
            ws.cell(row=r, column=1).value = date_label  # A列: 納品日
        ws.cell(row=r, column=2).value = item["name"]    # B列: 商品名
        ws.cell(row=r, column=10).value = item["quantity"] # J列: 数量
        ws.cell(row=r, column=11).value = "個"           # K列: 単位

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---- UI ----

col1, col2 = st.columns(2)
with col1:
    reiwa_now = date.today().year - 2018
    billing_month = st.text_input("請求月（件名に使用）", value=f"{reiwa_now}年{date.today().month}月分")
with col2:
    st.write("")
    st.caption("例: 8年5月分　→　件名「令和8年5月分味噌加工品代金」")

uploaded_files = st.file_uploader(
    "納品書Excelをアップロード（複数可）",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"{len(uploaded_files)}枚の納品書を読み込みました")

st.divider()

if st.button("請求書を自動生成する", type="primary", use_container_width=True):
    if not uploaded_files:
        st.warning("納品書ファイルをアップロードしてください")
        st.stop()

    store_data = {}
    errors = []

    progress = st.progress(0, text="読み取り中...")
    for i, f in enumerate(uploaded_files):
        progress.progress((i + 1) / len(uploaded_files), text=f"読み取り: {f.name}")
        result = read_nouhinshо(f.read())
        if result is None:
            errors.append(f.name)
            continue
        store = result["store_name"]
        if store not in store_data:
            store_data[store] = {"items": [], "dates": []}
        store_data[store]["items"].extend(result["items"])
        if result["delivery_date"]:
            store_data[store]["dates"].append(result["delivery_date"])
    progress.empty()

    if errors:
        st.warning(f"読み取れなかったファイル: {', '.join(errors)}")
    if not store_data:
        st.error("データを抽出できませんでした")
        st.stop()

    st.success(f"{len(store_data)}店舗分の請求書を生成しました")
    st.subheader("内容確認")
    st.caption("金額を必ず確認してからダウンロードしてください")

    all_ok = True
    for store_name, data in store_data.items():
        merged = {}
        for item in data["items"]:
            k = item["name"]
            if k in merged:
                merged[k]["quantity"] += item["quantity"]
                merged[k]["amount"]   += item["amount"]
            else:
                merged[k] = dict(item)
        merged_items = list(merged.values())

        subtotal = sum(i["amount"] for i in merged_items)
        tax      = int(subtotal * TAX_RATE)
        total    = subtotal + tax

        # 単価×数量の検証
        check_errors = [
            f"{i['name']}: 計算値¥{i['unit_price']*i['quantity']:,} ≠ ¥{i['amount']:,}"
            for i in merged_items if i["unit_price"] * i["quantity"] != i["amount"]
        ]

        with st.expander(
            f"{'✅' if not check_errors else '⚠️'} {store_name}　合計: ¥{total:,}（税込）",
            expanded=True
        ):
            if check_errors:
                st.error("金額の不一致: " + " / ".join(check_errors))
                all_ok = False

            df = pd.DataFrame(merged_items)[["name", "unit_price", "quantity", "amount"]]
            df.columns = ["商品名", "単価", "数量", "金額"]
            df["単価"] = df["単価"].map(lambda x: f"¥{x:,}")
            df["金額"] = df["金額"].map(lambda x: f"¥{x:,}")
            st.dataframe(df, use_container_width=True, hide_index=True)

            c1, c2, c3 = st.columns(3)
            c1.metric("小計（税抜）", f"¥{subtotal:,}")
            c2.metric("消費税（8%）", f"¥{tax:,}")
            c3.metric("合計（税込）", f"¥{total:,}")

        store_data[store]["merged_items"] = merged_items

    st.divider()

    if not all_ok:
        st.error("金額の不一致があります。確認してください。")
    else:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w") as zf:
            for store_name, data in store_data.items():
                excel_bytes = create_invoice(
                    store_name,
                    data["merged_items"],
                    billing_month,
                    data["dates"]
                )
                safe = store_name.replace("/", "_").replace(" ", "_").replace("　", "_")
                zf.writestr(f"令和{date.today().year-2018}年{billing_month}請求書（{safe}）.xlsx", excel_bytes)

        zip_buf.seek(0)
        st.download_button(
            label=f"全{len(store_data)}店舗の請求書をZIPでダウンロード",
            data=zip_buf,
            file_name=f"請求書一括_{billing_month}.zip",
            mime="application/zip",
            use_container_width=True,
            type="primary",
        )
