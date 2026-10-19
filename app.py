import re
from io import BytesIO
from datetime import date
from uuid import uuid4

import pandas as pd
import streamlit as st

st.set_page_config(page_title="員工報銷小助手", page_icon="🧾", layout="wide")
st.title("🧾 員工費用報銷小助手")
st.caption("可依序輸入多位員工，最後一次下載合併後的 Excel 檔")

TAX_RATE = 0.05


def calculate_tax(net_amount):
    return int(float(net_amount) * TAX_RATE + 0.5)


def new_invoice():
    return {
        "id": str(uuid4()), "category": "⛽ 油資", "invoice_date": date.today(),
        "invoice_no": "", "seller_id": "", "net_amount": 0.0, "tax_amount": 0,
    }


if "all_expense_rows" not in st.session_state:
    st.session_state.all_expense_rows = []
if "invoices" not in st.session_state:
    st.session_state.invoices = [new_invoice()]
if "employee_name" not in st.session_state:
    st.session_state.employee_name = ""
if "report_date" not in st.session_state:
    st.session_state.report_date = date.today()

# 在新一輪畫面開始時重設本位員工表單
if st.session_state.get("reset_employee_form", False):
    st.session_state.employee_name = ""
    st.session_state.report_date = date.today()
    st.session_state.invoices = [new_invoice()]
    st.session_state.reset_employee_form = False

# 相容先前版本的 session 資料
for invoice in st.session_state.invoices:
    invoice.setdefault("id", str(uuid4()))
    if "net_amount" not in invoice:
        invoice["net_amount"] = float(invoice.pop("amount", 0.0))
    invoice.setdefault("invoice_date", date.today())
    invoice.setdefault("invoice_no", "")
    invoice.setdefault("seller_id", "")
    invoice.setdefault("category", "⛽ 油資")
    if "tax_amount" not in invoice:
        invoice["tax_amount"] = calculate_tax(invoice.get("net_amount", 0.0))


def update_tax_from_net(invoice_id):
    for item in st.session_state.invoices:
        if item["id"] == invoice_id:
            net_key = f"net_amount_{invoice_id}"
            tax_key = f"tax_amount_{invoice_id}"
            net = float(st.session_state.get(net_key, item["net_amount"]))
            item["net_amount"] = net
            item["tax_amount"] = calculate_tax(net)
            st.session_state[tax_key] = item["tax_amount"]
            break


def validate_current_employee():
    errors = []
    if not st.session_state.employee_name.strip():
        errors.append("請先輸入員工姓名。")
    if not st.session_state.invoices:
        errors.append("至少需要一張發票。")
    for i, invoice in enumerate(st.session_state.invoices):
        n = i + 1
        if not invoice["invoice_no"].strip():
            errors.append(f"第 {n} 張發票尚未填寫發票號碼。")
        if not re.fullmatch(r"\d{8}", invoice["seller_id"].strip()):
            errors.append(f"第 {n} 張發票的賣方統編必須是 8 位數字。")
        if float(invoice["net_amount"]) <= 0:
            errors.append(f"第 {n} 張發票未稅金額必須大於 0。")
    return errors


def current_employee_rows():
    rows = []
    name = st.session_state.employee_name.strip()
    report_day = st.session_state.report_date
    for invoice in st.session_state.invoices:
        net = float(invoice["net_amount"])
        tax = int(invoice["tax_amount"])
        rows.append({
            "員工姓名": name, "報銷日期": report_day,
            "費用類別": invoice["category"], "發票日期": invoice["invoice_date"],
            "發票號碼": invoice["invoice_no"].strip(), "賣方統編": invoice["seller_id"].strip(),
            "未稅金額": net, "稅額": tax, "含稅總額": net + tax,
        })
    return rows


def create_excel(rows):
    columns = ["員工姓名", "報銷日期", "費用類別", "發票日期", "發票號碼", "賣方統編", "未稅金額", "稅額", "含稅總額"]
    df = pd.DataFrame(rows, columns=columns)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="報銷明細")
        df[df["費用類別"] == "⛽ 油資"].to_excel(writer, index=False, sheet_name="油資")
        df[df["費用類別"] == "🅿️ 停車"].to_excel(writer, index=False, sheet_name="停車")

        if not df.empty:
            summary_rows = []
            for (name, report_day), group in df.groupby(["員工姓名", "報銷日期"], dropna=False):
                summary_rows.append({
                    "員工姓名": name,
                    "報銷日期": report_day,
                    "停車未稅總計": group.loc[group["費用類別"] == "🅿️ 停車", "未稅金額"].sum(),
                    "油資未稅總計": group.loc[group["費用類別"] == "⛽ 油資", "未稅金額"].sum(),
                    "稅額總計": group["稅額"].sum(),
                    "報銷總計含稅": group["含稅總額"].sum(),
                })
            summary = pd.DataFrame(summary_rows)
        else:
            summary = pd.DataFrame(columns=["員工姓名", "報銷日期", "停車未稅總計", "油資未稅總計", "稅額總計", "報銷總計含稅"])
        summary.to_excel(writer, index=False, sheet_name="員工統計")

        money_headers = ["未稅金額", "稅額", "含稅總額", "停車未稅總計", "油資未稅總計", "稅額總計", "報銷總計含稅"]
        for worksheet in writer.sheets.values():
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
            for cells in worksheet.columns:
                max_length = max(len(str(cell.value or "")) for cell in cells)
                letter = cells[0].column_letter
                worksheet.column_dimensions[letter].width = min(max_length + 3, 30)
            headers = {cell.value: cell.column for cell in worksheet[1]}
            for header in money_headers:
                if header in headers:
                    col_index = headers[header]
                    for row_idx in range(2, worksheet.max_row + 1):
                        worksheet.cell(row=row_idx, column=col_index).number_format = '#,##0.00'
    output.seek(0)
    return output.getvalue()


st.subheader("👤 本次員工基本資料")
st.text_input("員工姓名", placeholder="請輸入姓名", key="employee_name")
st.date_input("報銷日期", key="report_date")
st.divider()

st.subheader("🧾 本次員工發票明細")
st.info("每位員工填寫完成後，按「加入總表並新增下一位員工」。系統會保存本位員工資料並清空表單；全部輸入完後，再一次下載所有員工的 Excel。未稅金額會先依 5% 預填稅額，稅額仍可手動修改。")

for i, invoice in enumerate(st.session_state.invoices):
    invoice_id = invoice["id"]
    st.markdown(f"### 第 {i + 1} 張發票")
    col1, col2, col3 = st.columns(3)
    with col1:
        category = st.selectbox("費用類別", ["⛽ 油資", "🅿️ 停車"], index=0 if invoice["category"] == "⛽ 油資" else 1, key=f"category_{invoice_id}")
    with col2:
        invoice_date = st.date_input("發票日期", value=invoice["invoice_date"], key=f"date_{invoice_id}")
    with col3:
        invoice_no = st.text_input("發票號碼", value=invoice["invoice_no"], placeholder="例如 AB12345678", key=f"invoice_no_{invoice_id}")
    col4, col5, col6 = st.columns(3)
    with col4:
        seller_id = st.text_input("賣方統編", value=invoice["seller_id"], placeholder="8 碼統一編號", key=f"seller_id_{invoice_id}")
    with col5:
        net_amount = st.number_input("未稅金額", min_value=0.0, value=float(invoice["net_amount"]), step=1.0, key=f"net_amount_{invoice_id}", on_change=update_tax_from_net, args=(invoice_id,))
    with col6:
        tax_amount = st.number_input("稅額（可手動修改）", min_value=0, value=int(invoice.get("tax_amount", calculate_tax(net_amount))), step=1, key=f"tax_amount_{invoice_id}")
    invoice["category"] = category
    invoice["invoice_date"] = invoice_date
    invoice["invoice_no"] = invoice_no
    invoice["seller_id"] = seller_id
    invoice["net_amount"] = net_amount
    invoice["tax_amount"] = tax_amount
    if len(st.session_state.invoices) > 1 and st.button("🗑️ 刪除這張發票", key=f"delete_{invoice_id}"):
        st.session_state.invoices.pop(i)
        st.rerun()
    st.divider()

if st.button("➕ 新增一張發票", use_container_width=True):
    st.session_state.invoices.append(new_invoice())
    st.rerun()

st.subheader("💰 本位員工報銷統計")
parking_net_total = sum(float(x["net_amount"]) for x in st.session_state.invoices if x["category"] == "🅿️ 停車")
oil_net_total = sum(float(x["net_amount"]) for x in st.session_state.invoices if x["category"] == "⛽ 油資")
tax_total = sum(int(x["tax_amount"]) for x in st.session_state.invoices)
grand_total = sum(float(x["net_amount"]) + int(x["tax_amount"]) for x in st.session_state.invoices)
c1, c2 = st.columns(2)
c1.metric("🅿️ 停車未稅總計", f"${parking_net_total:,.0f}")
c2.metric("⛽ 油資未稅總計", f"${oil_net_total:,.0f}")
c3, c4 = st.columns(2)
c3.metric("稅額總計", f"${tax_total:,.0f}")
c4.metric("💰 報銷總計（含稅）", f"${grand_total:,.0f}")
st.divider()

st.subheader("📚 多位員工合併報銷")
if st.button("➕ 加入總表並新增下一位員工", type="primary", use_container_width=True):
    errors = validate_current_employee()
    if errors:
        for error in errors:
            st.error(f"❗ {error}")
    else:
        st.session_state.all_expense_rows.extend(current_employee_rows())
        st.session_state.reset_employee_form = True
        st.rerun()

saved_rows = st.session_state.all_expense_rows
if saved_rows:
    saved_employees = {row["員工姓名"] for row in saved_rows}
    st.success(f"已加入總表：{len(saved_employees)} 位員工、{len(saved_rows)} 張發票。")
    combined_df = pd.DataFrame(saved_rows)
    st.dataframe(combined_df[["員工姓名", "報銷日期", "費用類別", "發票號碼", "未稅金額", "稅額", "含稅總額"]], use_container_width=True, hide_index=True)
    try:
        excel_file = create_excel(saved_rows)
        st.download_button(label="📥 下載所有員工合併 Excel", data=excel_file, file_name=f"全體員工報銷總表_{date.today().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    except Exception as e:
        st.error(f"Excel 產生失敗：{e}")
    if st.button("🗑️ 清空已加入的總表資料", type="secondary"):
        st.session_state.all_expense_rows = []
        st.rerun()
else:
    st.caption("目前總表還沒有資料。先填完一位員工，再按「加入總表並新增下一位員工」。")

st.caption("注意：合併資料目前保存在這個瀏覽器工作階段中；若重新整理、關閉頁面或 app 重新啟動，請先下載 Excel 備份。此版本不會自動寫入公司共用雲端檔案。")


