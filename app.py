
import json
from datetime import date
from io import BytesIO

import pandas as pd
import requests
import streamlit as st


# ==================================================
# 1. 頁面設定
# ==================================================

st.set_page_config(
    page_title="員工報銷小助手",
    page_icon="🧾",
    layout="wide",
)

st.title("🧾 員工費用報銷小助手")
st.caption("油資與停車發票一次填寫，完成後可儲存至 Google 試算表並下載 Excel")


# ==================================================
# 2. 基本設定
# ==================================================

TAX_RATE = 0.05

CATEGORIES = ["⛽ 油資", "🅿️ 停車"]


def calculate_tax(net_amount):
    """依未稅金額計算 5% 稅額，四捨五入至整數。"""
    return int(float(net_amount or 0) * TAX_RATE + 0.5)


# ==================================================
# 3. Google 試算表儲存
# ==================================================

def save_to_google_sheets(employee_name, report_date, invoices):
    """將本次報銷資料傳送至 Google Apps Script。"""

    settings = st.secrets["google_sheets"]

    web_app_url = settings["web_app_url"]
    api_token = settings["api_token"]

    payload = {
        "token": api_token,
        "employee_name": employee_name.strip(),
        "report_date": report_date.isoformat(),
        "invoices": [
            {
                "category": invoice["category"],
                "invoice_date": invoice["invoice_date"].isoformat(),
                "invoice_no": invoice["invoice_no"].strip(),
                "seller_id": invoice["seller_id"].strip(),
                "pretax_amount": float(invoice["net_amount"]),
                "tax_amount": int(invoice["tax_amount"]),
            }
            for invoice in invoices
        ],
    }

    response = requests.post(
        web_app_url,
        data=json.dumps(payload, ensure_ascii=False),
        headers={"Content-Type": "application/json"},
        timeout=30,
    )

    response.raise_for_status()

    try:
        result = response.json()
    except ValueError as exc:
        raise RuntimeError(
            "Google 試算表回傳的內容不是 JSON，請檢查 Apps Script 部署設定。"
        ) from exc

    if not result.get("success"):
        raise RuntimeError(
            result.get("message", "Google 試算表儲存失敗。")
        )

    return result


# ==================================================
# 4. Excel 匯出
# ==================================================

def create_excel(employee_name, report_date, invoices):
    """建立含有報銷明細、油資、停車三個工作表的 Excel。"""

    rows = []

    for invoice in invoices:
        net_amount = float(invoice["net_amount"])
        tax_amount = int(invoice["tax_amount"])

        rows.append({
            "員工姓名": employee_name.strip(),
            "報銷日期": report_date.strftime("%Y-%m-%d"),
            "費用類別": invoice["category"],
            "發票日期": invoice["invoice_date"].strftime("%Y-%m-%d"),
            "發票號碼": invoice["invoice_no"].strip(),
            "賣方統編": invoice["seller_id"].strip(),
            "未稅金額": net_amount,
            "稅額": tax_amount,
            "含稅總額": net_amount + tax_amount,
        })

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

    detail_df = pd.DataFrame(rows, columns=columns)

    oil_df = detail_df[
        detail_df["費用類別"] == "⛽ 油資"
    ].copy()

    parking_df = detail_df[
        detail_df["費用類別"] == "🅿️ 停車"
    ].copy()

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        detail_df.to_excel(
            writer,
            sheet_name="報銷明細",
            index=False,
        )

        oil_df.to_excel(
            writer,
            sheet_name="油資",
            index=False,
        )

        parking_df.to_excel(
            writer,
            sheet_name="停車",
            index=False,
        )

        # 設定欄寬及標題列
        for sheet_name in ["報銷明細", "油資", "停車"]:
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions

            for cells in worksheet.columns:
                max_length = 0
                column_letter = cells[0].column_letter

                for cell in cells:
                    value = "" if cell.value is None else str(cell.value)
                    max_length = max(max_length, len(value))

                worksheet.column_dimensions[column_letter].width = min(
                    max(max_length + 2, 12),
                    24,
                )

            # 金額欄位格式
            header_map = {
                cell.value: cell.column
                for cell in worksheet[1]
            }

            for column_name in ["未稅金額", "稅額", "含稅總額"]:
                if column_name in header_map:
                    col_num = header_map[column_name]

                    for row_num in range(2, worksheet.max_row + 1):
                        worksheet.cell(
                            row=row_num,
                            column=col_num,
                        ).number_format = "#,##0.00"

    output.seek(0)
    return output.getvalue()


# ==================================================
# 5. 初始化發票資料
# ==================================================

def new_invoice():
    return {
        "category": "⛽ 油資",
        "invoice_date": date.today(),
        "invoice_no": "",
        "seller_id": "",
        "net_amount": 0.0,
        "tax_amount": 0,
    }


if "invoices" not in st.session_state:
    st.session_state.invoices = [new_invoice()]

if "excel_file" not in st.session_state:
    st.session_state.excel_file = None

if "excel_filename" not in st.session_state:
    st.session_state.excel_filename = None


def update_tax_from_net(index):
    """未稅金額改變時，自動更新稅額。"""
    net_key = f"net_amount_{index}"
    tax_key = f"tax_amount_{index}"

    net_amount = st.session_state.get(net_key, 0.0) or 0.0

    st.session_state[tax_key] = calculate_tax(net_amount)

    st.session_state.invoices[index]["net_amount"] = float(net_amount)
    st.session_state.invoices[index]["tax_amount"] = int(
        st.session_state[tax_key]
    )


def sync_invoice(index):
    """將各發票欄位的最新輸入同步到 session_state。"""
    invoice = st.session_state.invoices[index]

    invoice["category"] = st.session_state[f"category_{index}"]
    invoice["invoice_date"] = st.session_state[f"invoice_date_{index}"]
    invoice["invoice_no"] = st.session_state[f"invoice_no_{index}"]
    invoice["seller_id"] = st.session_state[f"seller_id_{index}"]
    invoice["net_amount"] = float(
        st.session_state.get(f"net_amount_{index}", 0.0) or 0.0
    )
    invoice["tax_amount"] = int(
        st.session_state.get(f"tax_amount_{index}", 0) or 0
    )


# ==================================================
# 6. 基本資料
# ==================================================

st.subheader("👤 基本資料")

col1, col2 = st.columns(2)

with col1:
    employee_name = st.text_input(
        "員工姓名",
        placeholder="請輸入姓名",
        key="employee_name",
    )

with col2:
    report_date = st.date_input(
        "報銷日期",
        value=date.today(),
        key="report_date",
    )


# ==================================================
# 7. 發票明細
# ==================================================

st.divider()
st.subheader("🧾 發票明細")

for i, invoice in enumerate(st.session_state.invoices):
    with st.container(border=True):
        st.markdown(f"**第 {i + 1} 張發票**")

        col1, col2, col3 = st.columns(3)

        with col1:
            st.selectbox(
                "費用類別",
                options=CATEGORIES,
                index=CATEGORIES.index(invoice["category"]),
                key=f"category_{i}",
                on_change=sync_invoice,
                args=(i,),
            )

        with col2:
            st.date_input(
                "發票日期",
                value=invoice["invoice_date"],
                key=f"invoice_date_{i}",
                on_change=sync_invoice,
                args=(i,),
            )

        with col3:
            st.text_input(
                "發票號碼",
                value=invoice["invoice_no"],
                key=f"invoice_no_{i}",
                placeholder="例如：AB12345678",
                on_change=sync_invoice,
                args=(i,),
            )

        col4, col5, col6 = st.columns(3)

        with col4:
            st.text_input(
                "賣方統編（8 碼）",
                value=invoice["seller_id"],
                key=f"seller_id_{i}",
                max_chars=8,
                on_change=sync_invoice,
                args=(i,),
            )

        with col5:
            st.number_input(
                "未稅金額",
                min_value=0.0,
                value=float(invoice["net_amount"]),
                step=1.0,
                key=f"net_amount_{i}",
                on_change=update_tax_from_net,
                args=(i,),
            )

        with col6:
            st.number_input(
                "稅額（可手動調整）",
                min_value=0,
                value=int(invoice["tax_amount"]),
                step=1,
                key=f"tax_amount_{i}",
                on_change=sync_invoice,
                args=(i,),
                help="未稅金額變動時會自動計算 5% 稅額；若有四捨五入差異，可直接修改。",
            )

        net_amount = float(invoice["net_amount"])
        tax_amount = int(invoice["tax_amount"])

        st.caption(
            f"含稅總額：{net_amount + tax_amount:,.2f} 元"
        )

        if len(st.session_state.invoices) > 1:
            if st.button(
                f"🗑️ 刪除第 {i + 1} 張發票",
                key=f"delete_{i}",
            ):
                st.session_state.invoices.pop(i)

                # 清除本頁產生的舊下載檔
                st.session_state.excel_file = None
                st.session_state.excel_filename = None

                st.rerun()


if st.button("➕ 新增一張發票", use_container_width=True):
    # 先同步目前畫面上的資料
    for i in range(len(st.session_state.invoices)):
        sync_invoice(i)

    st.session_state.invoices.append(new_invoice())
    st.session_state.excel_file = None
    st.session_state.excel_filename = None
    st.rerun()


# ==================================================
# 8. 報稅統計
# ==================================================

# 先同步欄位，確保統計使用最新資料
for i in range(len(st.session_state.invoices)):
    sync_invoice(i)

invoices = st.session_state.invoices

parking_tax_total = sum(
    int(inv["tax_amount"])
    for inv in invoices
    if inv["category"] == "🅿️ 停車"
)

oil_pretax_total = sum(
    float(inv["net_amount"])
    for inv in invoices
    if inv["category"] == "⛽ 油資"
)

tax_total = sum(
    int(inv["tax_amount"])
    for inv in invoices
)

report_total = sum(
    float(inv["net_amount"]) + int(inv["tax_amount"])
    for inv in invoices
)

st.divider()
st.subheader("📊 報稅統計")

col1, col2 = st.columns(2)

with col1:
    st.metric("停車位稅總計", f"{parking_tax_total:,.0f} 元")
    st.metric("油資未稅總計", f"{oil_pretax_total:,.2f} 元")

with col2:
    st.metric("稅額總計", f"{tax_total:,.0f} 元")
    st.metric("報稅總計", f"{report_total:,.2f} 元")


# ==================================================
# 9. 驗證、儲存及下載
# ==================================================

st.divider()
st.subheader("📤 送出報銷資料")

st.caption(
    "送出後會將本次發票明細儲存到 Google 試算表，"
    "並產生可下載的 Excel 檔案。"
)

if st.button(
    "💾 儲存資料並產生 Excel",
    type="primary",
    use_container_width=True,
):
    # 同步所有發票的最新資料
    for i in range(len(st.session_state.invoices)):
        sync_invoice(i)

    invoices = st.session_state.invoices
    errors = []

    if not employee_name.strip():
        errors.append("請輸入員工姓名。")

    if not invoices:
        errors.append("至少需要一張發票。")

    for i, invoice in enumerate(invoices, start=1):
        if not invoice["invoice_no"].strip():
            errors.append(f"第 {i} 張發票：請輸入發票號碼。")

        seller_id = invoice["seller_id"].strip()

        if not re.fullmatch(r"\d{8}", seller_id):
            errors.append(f"第 {i} 張發票：賣方統編必須為 8 位數字。")

        if float(invoice["net_amount"]) <= 0:
            errors.append(f"第 {i} 張發票：未稅金額必須大於 0。")

        if int(invoice["tax_amount"]) < 0:
            errors.append(f"第 {i} 張發票：稅額不能小於 0。")

    if errors:
        for error in errors:
            st.error(error)

    else:
        try:
            with st.spinner("正在儲存至 Google 試算表……"):
                save_to_google_sheets(
                    employee_name=employee_name,
                    report_date=report_date,
                    invoices=invoices,
                )

            excel_bytes = create_excel(
                employee_name=employee_name,
                report_date=report_date,
                invoices=invoices,
            )

            filename = (
                f"報銷_{employee_name.strip()}_"
                f"{report_date.strftime('%Y%m%d')}.xlsx"
            )

            st.session_state.excel_file = excel_bytes
            st.session_state.excel_filename = filename

            st.success("Google 試算表儲存成功！現在可以下載 Excel。")

        except Exception as e:
            st.error(f"儲存失敗：{e}")
            st.info(
                "請確認 Streamlit Secrets 的網址與密鑰正確，"
                "以及 Apps Script 已部署為可存取的網路應用程式。"
            )


if (
    st.session_state.excel_file is not None
    and st.session_state.excel_filename is not None
):
    st.download_button(
        label="⬇️ 下載 Excel 報銷檔",
        data=st.session_state.excel_file,
        file_name=st.session_state.excel_filename,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True,
    )

