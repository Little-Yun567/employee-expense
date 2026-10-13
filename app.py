import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date

st.set_page_config(
page_title="員工報銷小助手",
page_icon="🧾",
layout="wide"
)

st.title("🧾 員工費用報銷小助手")
st.caption("油資與停車發票一次填寫，完成後可下載 Excel")

# -------------------------
# 初始化發票資料
# -------------------------
if "invoices" not in st.session_state:
    st.session_state.invoices = [
        {
            "category": "⛽ 油資",
            "invoice_date": date.today(),
            "invoice_no": "",
            "seller_id": "",
            "amount": 0.0
        }
    ]
"category": "⛽ 油資",
"invoice_date": date.today(),
"invoice_no": "",
"seller_id": "",
"amount": 0.0
}
]

# -------------------------
# 基本資料
# -------------------------
st.subheader("👤 基本資料")

employee_name = st.text_input(
"員工姓名",
placeholder="請輸入姓名"
)

report_date = st.date_input(
"報銷日期",
value=date.today()
)

st.divider()

# -------------------------
# 發票明細
# -------------------------
st.subheader("🧾 發票明細")

st.info("一張發票填一列，需要幾張就按「＋ 新增發票」即可。")

for i, invoice in enumerate(st.session_state.invoices):

st.markdown(f"### 第 {i + 1} 張發票")

col1, col2, col3, col4, col5 = st.columns(
[1.2, 1.2, 1.8, 1.8, 1.2]
)

with col1:
category = st.selectbox(
"費用類別",
["⛽ 油資", "🅿️ 停車"],
index=0 if invoice["category"] == "⛽ 油資" else 1,
key=f"category_{i}"
)

with col2:
invoice_date = st.date_input(
"發票日期",
value=invoice["invoice_date"],
key=f"date_{i}"
)

with col3:
invoice_no = st.text_input(
"發票號碼",
value=invoice["invoice_no"],
placeholder="例如 AB12345678",
key=f"invoice_no_{i}"
)

with col4:
seller_id = st.text_input(
"賣方統編",
value=invoice["seller_id"],
placeholder="8碼統一編號",
key=f"seller_id_{i}"
)

with col5:
amount = st.number_input(
"金額",
min_value=0.0,
value=float(invoice["amount"]),
step=1.0,
key=f"amount_{i}"
)

# 更新資料
invoice["category"] = category
invoice["invoice_date"] = invoice_date
invoice["invoice_no"] = invoice_no
invoice["seller_id"] = seller_id
invoice["amount"] = amount

# 刪除按鈕
if len(st.session_state.invoices) > 1:
if st.button("🗑️ 刪除這張", key=f"delete_{i}"):
st.session_state.invoices.pop(i)
st.rerun()

st.divider()

# -------------------------
# 新增發票
# -------------------------
if st.button("➕ 新增一張發票", use_container_width=True):

st.session_state.invoices.append(
{
"category": "⛽ 油資",
"invoice_date": date.today(),
"invoice_no": "",
"seller_id": "",
"amount": 0.0
}
)

st.rerun()

# -------------------------
# 統計
# -------------------------
st.subheader("💰 報銷統計")

oil_total = sum(
x["amount"]
for x in st.session_state.invoices
if x["category"] == "⛽ 油資"
)

parking_total = sum(
x["amount"]
for x in st.session_state.invoices
if x["category"] == "🅿️ 停車"
)

grand_total = oil_total + parking_total

col1, col2, col3 = st.columns(3)

with col1:
st.metric("⛽ 油資合計", f"${oil_total:,.0f}")

with col2:
st.metric("🅿️ 停車合計", f"${parking_total:,.0f}")

with col3:
st.metric("💰 報銷總額", f"${grand_total:,.0f}")

st.divider()

# -------------------------
# Excel 產生功能
# -------------------------
def create_excel():

data = []

for invoice in st.session_state.invoices:

data.append({
"員工姓名": employee_name,
"報銷日期": report_date,
"費用類別": invoice["category"],
"發票日期": invoice["invoice_date"],
"發票號碼": invoice["invoice_no"],
"賣方統編": invoice["seller_id"],
"金額": invoice["amount"]
})

df = pd.DataFrame(data)

output = BytesIO()

with pd.ExcelWriter(
output,
engine="openpyxl"
) as writer:

df.to_excel(
writer,
index=False,
sheet_name="報銷明細"
)

# 油資
oil_df = df[df["費用類別"] == "⛽ 油資"]
oil_df.to_excel(
writer,
index=False,
sheet_name="油資"
)

# 停車
parking_df = df[df["費用類別"] == "🅿️ 停車"]
parking_df.to_excel(
writer,
index=False,
sheet_name="停車"
)

output.seek(0)

return output


# -------------------------
# 送出
# -------------------------
st.subheader("📊 完成報銷")

if st.button(
"📥 產生 Excel 報銷檔",
type="primary",
use_container_width=True
):

if not employee_name.strip():
st.error("❗ 請先輸入員工姓名")

else:

# 檢查發票資料
error = False

for i, invoice in enumerate(st.session_state.invoices):

if not invoice["invoice_no"].strip():
st.error(f"❗ 第 {i + 1} 張發票尚未填寫發票號碼")
error = True

if not invoice["seller_id"].strip():
st.error(f"❗ 第 {i + 1} 張發票尚未填寫賣方統編")
error = True

if invoice["amount"] <= 0:
st.error(f"❗ 第 {i + 1} 張發票金額必須大於 0")
error = True

if not error:

excel_file = create_excel()

filename = (
f"{employee_name}_報銷_"
f"{report_date.strftime('%Y%m%d')}.xlsx"
)

st.success("🎉 報銷資料整理完成！")

st.download_button(
label="⬇️ 下載 Excel",
data=excel_file,
file_name=filename,
mime=(
"application/vnd.openxmlformats-officedocument."
"spreadsheetml.sheet"
),
use_container_width=True
)
