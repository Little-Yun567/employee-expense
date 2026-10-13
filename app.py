
import re
from io import BytesIO
from datetime import date
from uuid import uuid4

import pandas as pd
import streamlit as st


# =========================
# 1. 頁面設定
# =========================
st.set_page_config(
    page_title="員工報銷小助手",
    page_icon="🧾",
    layout="wide",
)

st.title("🧾 員工費用報銷小助手")
st.caption("油資與停車發票一次填寫，完成後可下載 Excel")


# =========================
# 2. 初始化發票資料
# =========================
if "invoices" not in st.session_state:
    st.session_state.invoices = [
        {
            "id": str(uuid4()),
            "category": "⛽ 油資",
            "invoice_date": date.today(),
            "invoice_no": "",
            "seller_id": "",
            "net_amount": 0.0,
            "tax_amount": 0.0,
        }
    ]

# 確保每張發票都有獨立 ID
for invoice in st.session_state.invoices:
    if "id" not in invoice:
        invoice["id"] = str(uuid4())

    # 相容舊版發票資料
    if "net_amount" not in invoice:
        invoice["net_amount"] = float(
            invoice.pop("amount", 0.0)
        )

    if "tax_amount" not in invoice:
        invoice["tax_amount"] = 0.0

    if "invoice_date" not in invoice:
        invoice["invoice_date"] = date.today()

    if "invoice_no" not in invoice:
        invoice["invoice_no"] = ""

    if "seller_id" not in invoice:
        invoice["seller_id"] = ""

    if "category" not in invoice:
        invoice["category"] = "⛽ 油資"


# =========================
# 3. 基本資料
# =========================
st.subheader("👤 基本資料")

employee_name = st.text_input(
    "員工姓名",
    placeholder="請輸入姓名",
)

report_date = st.date_input(
    "報銷日期",
    value=date.today(),
)

st.divider()


# =========================
# 4. 發票明細
# =========================
st.subheader("🧾 發票明細")
st.info("一張發票填一列，可新增多張發票。")

for i, invoice in enumerate(st.session_state.invoices):
    invoice_id = invoice["id"]

    st.markdown(f"### 第 {i + 1} 張發票")

    col1, col2, col3 = st.columns([1, 1, 1])

    with col1:
        category = st.selectbox(
            "費用類別",
            ["⛽ 油資", "🅿️ 停車"],
            index=(
                0 if invoice["category"] == "⛽ 油資" else 1
            ),
            key=f"category_{invoice_id}",
        )

    with col2:
        invoice_date = st.date_input(
            "發票日期",
            value=invoice["invoice_date"],
            key=f"date_{invoice_id}",
        )

    with col3:
        invoice_no = st.text_input(
            "發票號碼",
            value=invoice["invoice_no"],
            placeholder="例如 AB12345678",
            key=f"invoice_no_{invoice_id}",
        )

    col4, col5, col6 = st.columns([1, 1, 1])

    with col4:
        seller_id = st.text_input(
            "賣方統編",
            value=invoice["seller_id"],
            placeholder="8 碼統一編號",
            key=f"seller_id_{invoice_id}",
        )

    with col5:
        net_amount = st.number_input(
            "未稅金額",
            min_value=0.0,
            value=float(invoice["net_amount"]),
            step=1.0,
            key=f"net_amount_{invoice_id}",
        )

    with col6:
        tax_amount = st.number_input(
            "稅額",
            min_value=0.0,
            value=float(invoice["tax_amount"]),
            step=1.0,
            key=f"tax_amount_{invoice_id}",
        )

    # 更新發票資料
    invoice["category"] = category
    invoice["invoice_date"] = invoice_date
    invoice["invoice_no"] = invoice_no
    invoice["seller_id"] = seller_id
    invoice["net_amount"] = net_amount
    invoice["tax_amount"] = tax_amount

    # 計算含稅總額
    total_amount = net_amount + tax_amount

    st.metric(
        f"第 {i + 1} 張發票含稅總額",
        f"${total_amount:,.0f}",
    )

    # 刪除發票
    if len(st.session_state.invoices) > 1:
        if st.button(
            "🗑️ 刪除這張發票",
            key=f"delete_{invoice_id}",
        ):
            st.session_state.invoices.pop(i)
            st.rerun()

    st.divider()


# =========================
# 5. 新增發票
# =========================
if st.button(
    "➕ 新增一張發票",
    use_container_width=True,
):
    st.session_state.invoices.append(
        {
            "id": str(uuid4()),
            "category": "⛽ 油資",
            "invoice_date": date.today(),
            "invoice_no": "",
            "seller_id": "",
            "net_amount": 0.0,
            "tax_amount": 0.0,
        }
    )
    st.rerun()


# =========================
# 6. 報銷統計
# =========================
st.subheader("💰 報銷統計")

net_total = sum(
    x["net_amount"]
    for x in st.session_state.invoices
)

tax_total = sum(
    x["tax_amount"]
    for x in st.session_state.invoices
)

oil_total = sum(
    x["net_amount"] + x["tax_amount"]
    for x in st.session_state.invoices
    if x["category"] == "⛽ 油資"
)

parking_total = sum(
    x["net_amount"] + x["tax_amount"]
    for x in st.session_state.invoices
    if x["category"] == "🅿️ 停車"
)

grand_total = net_total + tax_total

col1, col2, col3 = st.columns(3)

with col1:
    st.metric(
        "未稅金額合計",
        f"${net_total:,.0f}",
    )

with col2:
    st.metric(
        "稅額合計",
        f"${tax_total:,.0f}",
    )

with col3:
    st.metric(
        "💰 報銷總額（含稅）",
        f"${grand_total:,.0f}",
    )

col4, col5 = st.columns(2)

with col4:
    st.metric(
        "⛽ 油資合計（含稅）",
        f"${oil_total:,.0f}",
    )

with col5:
    st.metric(
        "🅿️ 停車合計（含稅）",
        f"${parking_total:,.0f}",
    )

st.divider()


# =========================
# 7. 建立 Excel
# =========================
def create_excel():
    data = []

    for invoice in st.session_state.invoices:
        net = invoice["net_amount"]
        tax = invoice["tax_amount"]

        data.append(
            {
                "員工姓名": employee_name.strip(),
                "報銷日期": report_date,
                "費用類別": invoice["category"],
                "發票日期": invoice["invoice_date"],
                "發票號碼": invoice["invoice_no"].strip(),
                "賣方統編": invoice["seller_id"].strip(),
                "未稅金額": net,
                "稅額": tax,
                "含稅總額": net + tax,
            }
        )

    columns = [
        "員工姓名",
        "報銷日期",
        "費用類別",
        "發票日期",
        "發票號碼",
        "賣方統編",
        "未稅金額",
        "稅額",
        "含稅總額",
    ]

    df = pd.DataFrame(data, columns=columns)

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:
        # 全部報銷明細
        df.to_excel(
            writer,
            index=False,
            sheet_name="報銷明細",
        )

        # 油資明細
        df[df["費用類別"] == "⛽ 油資"].to_excel(
            writer,
            index=False,
            sheet_name="油資",
        )

        # 停車明細
        df[df["費用類別"] == "🅿️ 停車"].to_excel(
            writer,
            index=False,
            sheet_name="停車",
        )

        # 調整 Excel 欄寬
        for worksheet in writer.sheets.values():
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions

            for column_cells in worksheet.columns:
                max_length = max(
                    len(str(cell.value or ""))
                    for cell in column_cells
                )
                column_letter = column_cells[0].column_letter

                worksheet.column_dimensions[
                    column_letter
                ].width = min(max_length + 3, 30)

            # 金額欄位使用數字格式
            headers = {
                cell.value: cell.column
                for cell in worksheet[1]
            }

            for header in ["未稅金額", "稅額", "含稅總額"]:
                if header in headers:
                    col_index = headers[header]

                    for row in range(2, worksheet.max_row + 1):
                        worksheet.cell(
                            row=row,
                            column=col_index,
                        ).number_format = '#,##0.00'

    output.seek(0)
    return output.getvalue()


# =========================
# 8. 送出及下載
# =========================
st.subheader("📊 完成報銷")

if st.button(
    "📥 產生 Excel 報銷檔",
    type="primary",
    use_container_width=True,
):
    errors = []

    if not employee_name.strip():
        errors.append("請先輸入員工姓名。")

    if not st.session_state.invoices:
        errors.append("至少需要一張發票。")

    for i, invoice in enumerate(st.session_state.invoices):
        row_number = i + 1

        if not invoice["invoice_no"].strip():
            errors.append(
                f"第 {row_number} 張發票尚未填寫發票號碼。"
            )

        seller_id = invoice["seller_id"].strip()

        if not re.fullmatch(r"\d{8}", seller_id):
            errors.append(
                f"第 {row_number} 張發票的賣方統編"
                "必須是 8 位數字。"
            )

        if (
            invoice["net_amount"] + invoice["tax_amount"]
        ) <= 0:
            errors.append(
                f"第 {row_number} 張發票含稅總額必須大於 0。"
            )

        if invoice["net_amount"] < 0 or invoice["tax_amount"] < 0:
            errors.append(
                f"第 {row_number} 張發票金額不可為負數。"
            )

    if errors:
        for error in errors:
            st.error(f"❗ {error}")

    else:
        try:
            excel_file = create_excel()

            safe_name = re.sub(
                r'[\\/:*?"<>|]',
                "_",
                employee_name.strip(),
            )

            filename = (
                f"{safe_name}_報銷_"
                f"{report_date.strftime('%Y%m%d')}.xlsx"
            )

            st.success("🎉 報銷資料整理完成！")

            st.download_button(
                label="⬇️ 下載 Excel 報銷檔",
                data=excel_file,
                file_name=filename,
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"Excel 產生失敗：{e}")

