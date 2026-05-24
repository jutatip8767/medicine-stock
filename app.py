import streamlit as st
import pandas as pd
from pathlib import Path
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="ตัดยอดยา", layout="wide")

st.title("โปรแกรมตัดยอดยาในสถานพยาบาล")

DATA_FILE = Path("medicine_stock.csv")


# =========================
# สร้างตารางเปล่า
# =========================
def create_template():
    return pd.DataFrame(
        columns=[
            "วันที่รับยา",
            "ชื่อยา",
            "ยอดรับเข้า",
            "ยอดใช้ไป",
            "ราคาต่อหน่วย",
        ]
    )


# =========================
# โหลดข้อมูล
# =========================
if DATA_FILE.exists():
    df = pd.read_csv(DATA_FILE)
else:
    df = create_template()


# =========================
# ตารางกรอกข้อมูล
# =========================
st.subheader("กรอกข้อมูลยา")

edited_df = st.data_editor(
    df,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "วันที่รับยา": st.column_config.TextColumn(
            "วันที่รับยา (เช่น 24/5/69)",
            help="พิมพ์วันที่ให้สม่ำเสมอ"
        ),

        "ชื่อยา": st.column_config.TextColumn(
            "ชื่อยา"
        ),

        "ยอดรับเข้า": st.column_config.NumberColumn(
            "ยอดรับเข้า",
            min_value=0,
            step=1
        ),

        "ยอดใช้ไป": st.column_config.NumberColumn(
            "ยอดใช้ไป",
            min_value=0,
            step=1
        ),

        "ราคาต่อหน่วย": st.column_config.TextColumn(
            "ราคาต่อหน่วย",
            help="ใส่ทศนิยมได้ทุกแบบ เช่น 0.05 หรือ 1.237"
        ),
    }
)


# =========================
# แปลงข้อมูลตัวเลข
# =========================
edited_df["ยอดรับเข้า"] = pd.to_numeric(
    edited_df["ยอดรับเข้า"],
    errors="coerce"
).fillna(0)

edited_df["ยอดใช้ไป"] = pd.to_numeric(
    edited_df["ยอดใช้ไป"],
    errors="coerce"
).fillna(0)

edited_df["ราคาต่อหน่วย"] = pd.to_numeric(
    edited_df["ราคาต่อหน่วย"],
    errors="coerce"
).fillna(0)


# =========================
# สรุปยอดคงเหลือปัจจุบัน
# =========================
summary = (
    edited_df.groupby("ชื่อยา", as_index=False)
    .agg({
        "ยอดรับเข้า": "sum",
        "ยอดใช้ไป": "sum",
        "ราคาต่อหน่วย": "last"
    })
)

summary["ยอดคงเหลือ"] = (
    summary["ยอดรับเข้า"] - summary["ยอดใช้ไป"]
)

summary["ยอดคงเหลือ"] = summary["ยอดคงเหลือ"].clip(lower=0)

summary["มูลค่ายาคงเหลือ"] = (
    summary["ยอดคงเหลือ"] * summary["ราคาต่อหน่วย"]
)

summary = summary.rename(columns={
    "ยอดรับเข้า": "ยอดรับสะสม",
    "ยอดใช้ไป": "ยอดใช้สะสม"
})


# =========================
# รวมยอดทั้งหมด
# =========================
total_qty = float(summary["ยอดคงเหลือ"].sum())

total_value = float(summary["มูลค่ายาคงเหลือ"].sum())


# =========================
# แสดงผลสรุป
# =========================
st.divider()

st.header("สรุปยอดคงเหลือปัจจุบัน")

col1, col2 = st.columns(2)

with col1:
    st.metric(
        "ยอดคงเหลือรวม (หน่วย)",
        f"{total_qty:,.0f}"
    )

with col2:
    st.metric(
        "มูลค่าคงเหลือรวม (บาท)",
        f"{total_value:,.2f}"
    )


# =========================
# ตารางสรุป
# =========================
st.subheader("ตารางสรุปผลปัจจุบัน")

show_summary = summary[
    [
        "ชื่อยา",
        "ยอดรับสะสม",
        "ยอดใช้สะสม",
        "ยอดคงเหลือ",
        "ราคาต่อหน่วย",
        "มูลค่ายาคงเหลือ"
    ]
]

st.dataframe(
    show_summary,
    use_container_width=True
)


# =========================
# ประวัติทั้งหมด
# =========================
st.subheader("ประวัติการตัดยอดทั้งหมด")

history = edited_df.copy()

history["ยอดคงเหลือ"] = (
    history["ยอดรับเข้า"] - history["ยอดใช้ไป"]
)

history["ยอดคงเหลือ"] = history["ยอดคงเหลือ"].clip(lower=0)

history["มูลค่ายาคงเหลือ"] = (
    history["ยอดคงเหลือ"] * history["ราคาต่อหน่วย"]
)

st.dataframe(
    history,
    use_container_width=True
)


# =========================
# ปุ่มบันทึก CSV
# =========================
if st.button("บันทึกข้อมูลลงไฟล์ CSV"):
    edited_df.to_csv(DATA_FILE, index=False)
    st.success("บันทึกข้อมูลเรียบร้อยแล้ว")


# =========================
# ดาวน์โหลด CSV
# =========================
csv_data = edited_df.to_csv(index=False).encode("utf-8-sig")

st.download_button(
    label="ดาวน์โหลด CSV",
    data=csv_data,
    file_name="medicine_stock.csv",
    mime="text/csv"
)


# =========================
# Export Word
# =========================
def export_word_report(summary_df):

    doc = Document()

    title = doc.add_heading(
        "รายงานตัดยอดยาในสถานพยาบาล",
        level=1
    )

    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    p.add_run(
        f"ยอดคงเหลือรวม (หน่วย): {total_qty:,.0f}"
    ).bold = True

    p.add_run(
        f"\nมูลค่าคงเหลือรวม (บาท): {total_value:,.2f}"
    ).bold = True

    table = doc.add_table(
        rows=1,
        cols=len(summary_df.columns)
    )

    table.style = "Table Grid"

    hdr_cells = table.rows[0].cells

    for i, col in enumerate(summary_df.columns):
        hdr_cells[i].text = col

    for _, row in summary_df.iterrows():

        row_cells = table.add_row().cells

        for i, value in enumerate(row):
            row_cells[i].text = str(value)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer


word_buffer = export_word_report(show_summary)

st.download_button(
    label="ดาวน์โหลด Word (.docx)",
    data=word_buffer,
    file_name="medicine_stock_report.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)