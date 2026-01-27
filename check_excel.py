import streamlit as st
import pandas as pd
import io
import re
import gc # <--- สำหรับจัดการคืน Memory
from datetime import datetime

# --- 1. การตั้งค่าหน้าเว็บ ---
st.set_page_config(page_title="Employee Data Validator Pro", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #FFF5EE; }
    .main-header {
        background-color: #FF8C00; color: white; padding: 20px;
        border-radius: 10px; text-align: center; margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<div class="main-header"><h1>🎯 ตรวจเช็คไฟล์นำเข้าข้อมูลพนักงาน</h1></div>', unsafe_allow_html=True)
st.write("🟠 **สีส้ม**: ค่าว่าง | 🔴 **สีแดง**: ข้อมูลผิด (ระบบจะแก้ฟอร์แมตวันที่/ตัวเลข และบังคับเป็น Text ให้อัตโนมัติ)")

# --- การตั้งค่าคอลัมน์ ---
ORANGE_INDICES = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15, 16, 17, 20, 21, 23, 24, 25, 39, 40, 41, 64, 65]
MAP_COLS = ["คำนำหน้า", "เพศ", "ระดับ", "ตำแหน่ง", "บริษัท", "สายงาน", "ฝ่าย", "แผนก", "สถานะพนักงาน", "ประเภทการจ้างงาน"]
DATE_COLS_IDX = [1, 25]

def smart_date_parser(val):
    if isinstance(val, datetime): return val, True
    val_str = str(val).strip()
    if not val_str or val_str.lower() == 'nan': return None, False
    
    clean_date = re.sub(r'[.\- ]', '/', val_str)
    formats = ['%d/%m/%Y', '%m/%d/%Y', '%Y/%m/%d', '%d/%m/%y', '%y/%m/%d']
    
    for fmt in formats:
        try:
            dt_obj = datetime.strptime(clean_date, fmt)
            if dt_obj.year > 2500: dt_obj = dt_obj.replace(year=dt_obj.year - 543)
            return dt_obj, True
        except ValueError: continue
    return val, False

def process_excel_data(uploaded_file):
    # ป้องกัน Memory Leak โดยการเคลียร์ pointer
    uploaded_file.seek(0)
    
    # อ่านไฟล์แบบระบุ dtype=object เพื่อความแม่นยำของข้อความ
    df_emp = pd.read_excel(uploaded_file, sheet_name="พนักงาน", dtype=object)
    uploaded_file.seek(0)
    df_ref = pd.read_excel(uploaded_file, sheet_name="รายละเอียด (ห้ามแก้ไข)", dtype=object)
    
    ref_data = {col: set(df_ref[col].dropna().astype(str).str.strip().unique()) 
                for col in MAP_COLS if col in df_ref.columns}
    
    error_details = []

    # วนลูปเช็คและแก้ไขข้อมูล
    for r_idx in range(len(df_emp)):
        for c_idx in ORANGE_INDICES:
            if c_idx >= len(df_emp.columns): continue
            
            val = df_emp.iloc[r_idx, c_idx]
            col_name = df_emp.columns[c_idx]
            val_str = str(val).strip() if pd.notna(val) else ""
            
            reason, color = "", ""

            if val_str == "" or val_str.lower() == 'nan':
                reason, color = "⚠️ ห้ามว่าง: กรุณากรอกข้อมูล", '#FFCC99'
            else:
                if c_idx in DATE_COLS_IDX:
                    dt_obj, success = smart_date_parser(val)
                    if success:
                        df_emp.iloc[r_idx, c_idx] = dt_obj
                    else:
                        reason, color = "❌ วันที่ผิดฟอร์แมต", '#FFC7CE'
                
                elif col_name in ["เลขบัตรประชาชน", "เลขประกันสังคม"]:
                    clean_id = re.sub(r'\D', '', val_str)
                    df_emp.iloc[r_idx, c_idx] = str(clean_id) # บังคับ Text
                    if len(clean_id) != 13:
                        reason, color = f"❌ {col_name} ต้องมี 13 หลัก", '#FFC7CE'
                
                elif col_name == "เลขบัญชีธนาคาร":
                    clean_acc = re.sub(r'\D', '', val_str)
                    df_emp.iloc[r_idx, c_idx] = str(clean_acc) # บังคับ Text
                    if len(clean_acc) != 10:
                        reason, color = "❌ เลขบัญชีต้องมี 10 หลัก", '#FFC7CE'
                        
                elif col_name in ref_data:
                    if val_str not in ref_data[col_name]:
                        reason, color = "❌ ข้อมูลไม่ตรงระบบอ้างอิง", '#FFC7CE'

            if reason:
                error_details.append({"row": r_idx + 1, "col": c_idx, "reason": reason, "color": color, "col_name": col_name})

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='dd/mm/yyyy') as writer:
        df_emp.to_excel(writer, index=False, sheet_name='พนักงาน')
        df_ref.to_excel(writer, index=False, sheet_name='รายละเอียด (ห้ามแก้ไข)')
        
        ws = writer.sheets['พนักงาน']
        workbook = writer.book
        
        # บังคับทั้งแผ่นงานเป็น Text และกำหนดรูปแบบสี
        text_fmt = workbook.add_format({'num_format': '@'})
        ws.set_column('A:ZZ', None, text_fmt)
        
        fmt_orange = workbook.add_format({'bg_color': '#FFCC99', 'border': 1, 'num_format': '@'})
        fmt_red = workbook.add_format({'bg_color': '#FFC7CE', 'border': 1, 'num_format': '@'})

        for err in error_details:
            f = fmt_orange if err['color'] == '#FFCC99' else fmt_red
            current_val = df_emp.iloc[err['row']-1, err['col']]
            ws.write(err['row'], err['col'], str(current_val) if pd.notna(current_val) else "", f)
            ws.write_comment(err['row'], err['col'], err['reason'], {'x_scale': 2.5})

    processed_data = output.getvalue()
    output.close()
    
    # เคลียร์ตัวแปรใหญ่เพื่อลด Memory
    del df_emp, df_ref
    return error_details, processed_data

# --- UI ส่วนล่าง ---
uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel ของคุณ", type=["xlsx"])

if uploaded_file:
    try:
        error_details, final_data = process_excel_data(uploaded_file)
        if error_details:
            st.error(f"🚩 พบจุดที่ต้องแก้ไข {len(error_details)} รายการ")
            st.download_button("📥 ดาวน์โหลดไฟล์ที่แก้ไขแล้ว", data=final_data, 
                               file_name=f"Recheck_{datetime.now().strftime('%H%M%S')}.xlsx", use_container_width=True)
            st.dataframe(pd.DataFrame([{"แถว": e['row']+1, "คอลัมน์": e['col_name'], "สาเหตุ": e['reason']} for e in error_details]), use_container_width=True)
        else:
            st.balloons()
            st.success("🎉 ข้อมูลถูกต้องทั้งหมดและถูกจัดฟอร์แมตเป็น Text เรียบร้อย!")
            st.download_button("📥 ดาวน์โหลดไฟล์ Clean Data", data=final_data, file_name="Clean_Data.xlsx", use_container_width=True)
        
        # บังคับเก็บกวาด Memory ทันทีที่จบกระบวนการ
        gc.collect()
        
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")
else:
    st.info("💡 อัปโหลดไฟล์ Excel เพื่อเริ่มการตรวจสอบ")