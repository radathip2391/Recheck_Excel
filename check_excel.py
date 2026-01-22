import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

# --- 1. ตั้งค่าหน้าตาเว็บและพื้นหลังสีส้มอ่อน ---
st.set_page_config(page_title="Employee Data Validator Pro", layout="wide")

st.markdown("""
    <style>
    .stApp {
        background-color: #FFF5EE; /* สีส้มอ่อน Seashell */
    }
    .main-header {
        background-color: #FF8C00;
        color: white;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<div class="main-header"><h1>🎯 ตรวจเช็คไฟล์นำเข้าข้อมูลพนักงาน</h1></div>', unsafe_allow_html=True)
st.write("🟠 **สีส้ม**: ค่าว่าง (ต้องกรอก) | 🔴 **สีแดง**: ข้อมูลผิด (ไม่ตรงระบบ / เลขบัตรไม่ครบ / วันที่ผิดฟอร์แมต)")

# 2. นิยามตำแหน่งคอลัมน์ (Index) และกฎการตรวจ
ORANGE_INDICES = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15, 16, 17, 20, 21, 23, 24, 25, 39, 40, 64, 65]
MAP_COLS = ["คำนำหน้า", "เพศ", "ระดับ", "ตำแหน่ง", "บริษัท", "สายงาน", "ฝ่าย", "แผนก", "สถานะพนักงาน", "ประเภทการจ้างงาน"]
DATE_COLS_IDX = [1, 25] # วันเริ่มงาน และ วันเกิด

uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel ของคุณ", type=["xlsx"])

if uploaded_file:
    try:
        # อ่านข้อมูลทั้ง 2 ชีทเก็บไว้
        df_emp = pd.read_excel(uploaded_file, sheet_name="พนักงาน")
        df_ref = pd.read_excel(uploaded_file, sheet_name="รายละเอียด (ห้ามแก้ไข)")
        
        # เตรียมฐานข้อมูลอ้างอิง
        ref_data = {}
        for col in MAP_COLS:
            if col in df_ref.columns:
                ref_data[col] = df_ref[col].dropna().astype(str).str.strip().unique().tolist()

        error_details = []

        # 3. เริ่มตรวจสอบข้อมูลในชีท "พนักงาน"
        for row_idx, row in df_emp.iterrows():
            for col_idx in ORANGE_INDICES:
                if col_idx < len(df_emp.columns):
                    val = row.iloc[col_idx]
                    col_name = df_emp.columns[col_idx]
                    val_str = str(val).strip() if pd.notna(val) else ""
                    
                    reason = ""
                    color = ""

                    # --- เงื่อนไข 1: ตรวจค่าว่าง -> มาร์คสีส้ม ---
                    if val_str == "":
                        reason = "⚠️ ห้ามว่าง: กรุณากรอกข้อมูล"
                        color = '#FFCC99' # ส้มอ่อน
                    
                    else:
                        # --- เงื่อนไข 2: ตรวจฟอร์แมตวันที่ -> มาร์คสีแดง ---
                        if col_idx in DATE_COLS_IDX:
                            is_date_valid = False
                            if isinstance(val, datetime):
                                is_date_valid = True
                            else:
                                for fmt in ('%d/%m/%Y', '%d-%m-%Y'):
                                    try:
                                        datetime.strptime(val_str, fmt)
                                        is_date_valid = True
                                        break
                                    except ValueError:
                                        continue
                            if not is_date_valid:
                                reason = "❌ วันที่ผิดฟอร์แมต: โปรดใช้ วัน/เดือน/ปี (เช่น 25/12/2023)"
                                color = '#FFC7CE' # แดง

                        # --- เงื่อนไข 3: เลขบัตรประชาชน 13 หลัก -> มาร์คสีแดง ---
                        elif col_name == "เลขบัตรประชาชน":
                            clean_id = re.sub(r'\D', '', val_str)
                            if len(clean_id) != 13:
                                reason = f"❌ รูปแบบผิด: เลขบัตรต้องครบ 13 หลัก (ปัจจุบัน {len(clean_id)} หลัก)"
                                color = '#FFC7CE' # แดง
                        
                        # --- เงื่อนไข 4: ข้อมูลไม่ตรงฐานข้อมูล -> มาร์คสีแดง ---
                        elif col_name in ref_data:
                            if val_str not in ref_data[col_name]:
                                reason = "❌ ข้อมูลไม่ตรงระบบ: โปรดเลือกค่าจากชีท 'รายละเอียด(ห้ามแก้ไข)'"
                                color = '#FFC7CE' # แดง

                    if reason:
                        error_details.append({
                            "row": row_idx + 1, "col": col_idx, 
                            "reason": reason, "color": color, "col_name": col_name
                        })

        # 4. สร้างไฟล์ Excel (เขียน 2 ชีทลงไป)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # เขียนชีทพนักงาน
            df_emp.to_excel(writer, index=False, sheet_name='พนักงาน')
            # เขียนชีทรายละเอียด (ห้ามแก้ไข) กลับไปด้วย
            df_ref.to_excel(writer, index=False, sheet_name='รายละเอียด (ห้ามแก้ไข)')
            
            workbook  = writer.book
            worksheet = writer.sheets['พนักงาน']
            
            # มาร์คจุดผิดเฉพาะในชีทพนักงาน
            for err in error_details:
                fmt = workbook.add_format({'bg_color': err['color'], 'border': 1})
                orig_val = df_emp.iloc[err['row']-1, err['col']]
                worksheet.write(err['row'], err['col'], orig_val if pd.notna(orig_val) else "", fmt)
                worksheet.write_comment(err['row'], err['col'], err['reason'], {'x_scale': 2.5, 'y_scale': 1})

        # 5. ส่วนแสดงผล
        if error_details:
            st.error(f"🚩 พบจุดที่ต้องแก้ไขทั้งหมด {len(error_details)} รายการ")
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์เพื่อทำการแก้ไข",
                data=output.getvalue(),
                file_name="Check_Result_Marked.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            summary_df = pd.DataFrame([{"แถว": e['row']+1, "คอลัมน์": e['col_name'], "สาเหตุ": e['reason']} for e in error_details])
            st.dataframe(summary_df, use_container_width=True)
        else:
            st.balloons()
            st.success("🎉 ถูกต้องทั้งหมด! ข้อมูลครบถ้วน ตรงระบบ และฟอร์แมตถูกต้อง")

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: โปรดเช็คชื่อชีท 'พนักงาน' และ 'รายละเอียด (ห้ามแก้ไข)' ในไฟล์ของคุณ")
else:
    st.info("💡 อัปโหลดไฟล์ Excel เพื่อเริ่มการตรวจสอบ")