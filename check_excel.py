import streamlit as st
import pandas as pd
import io
import re
import gc
from datetime import datetime

# --- 1. การตั้งค่าหน้าเว็บ ---
st.set_page_config(page_title="Employee Data Validator Pro", layout="wide")

@st.cache_data
def load_master_db_from_csv():
    file_path = "DataBaseจังหวัด.csv"
    for enc in ['utf-8-sig', 'tis-620']:
        try:
            db = pd.read_csv(file_path, dtype=str, encoding=enc, header=None)
            # โครงสร้างไฟล์: 0=ไปรษณีย์, 1=ตำบล, 7=อำเภอ, 10=จังหวัด (อ้างอิงตามโค้ดล่าสุดของคุณ)
            db_clean = pd.DataFrame({
                'zipcode': db[0],
                'subdistrict': db[1],
                'district': db[7],
                'province': db[10]
            }).apply(lambda x: x.str.strip())
            return db_clean
        except:
            continue
    return None

MASTER_DB = load_master_db_from_csv()

st.markdown("""
    <style>
    .stApp { background-color: #FFF5EE; }
    .main-header {
        background-color: #FF8C00; color: white; padding: 20px;
        border-radius: 10px; text-align: center; margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<div class="main-header"><h1>🎯 ตรวจเช็คข้อมูลพนักงาน </h1></div>', unsafe_allow_html=True)
st.write("🟠 **สีส้ม**: ค่าว่าง | 🔴 **สีแดง**: ข้อมูลผิด (ระบบจะแก้ฟอร์แมตวันที่/ตัวเลข และบังคับเป็น Text ให้อัตโนมัติ)")

# --- 2. การตั้งค่าคอลัมน์ ---
ORANGE_INDICES = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15, 16, 17, 20, 21, 23, 24, 25, 39, 40, 41, 64, 65]
DATE_COLS_IDX = [1, 25]

def smart_date_parser(val):
    if isinstance(val, datetime): return val, True
    val_str = str(val).strip()
    if not val_str or val_str.lower() == 'nan': return None, False
    clean_date = re.sub(r'[.\- ]', '/', val_str)
    for fmt in ['%d/%m/%Y', '%m/%d/%Y', '%Y/%m/%d', '%d/%m/%y', '%y/%m/%d']:
        try:
            dt_obj = datetime.strptime(clean_date, fmt)
            if dt_obj.year > 2500: dt_obj = dt_obj.replace(year=dt_obj.year - 543)
            return dt_obj, True
        except: continue
    return val, False

# --- 3. ฟังก์ชันหลักในการประมวลผล ---
def process_excel_data(uploaded_file):
    uploaded_file.seek(0)
    # อ่านไฟล์แบบระบุเฉพาะชีทที่ต้องการเพื่อประหยัด RAM
    try:
        df_emp = pd.read_excel(uploaded_file, sheet_name="พนักงาน", dtype=object).fillna("")
    except Exception as e:
        st.error(f"ไม่พบชีท 'พนักงาน' หรือไฟล์มีปัญหา: {e}")
        return None, None

    df_emp.columns = [str(c).replace('\n', ' ').strip() for c in df_emp.columns]
    error_details = []

    def find_idx(keywords):
        for i, col in enumerate(df_emp.columns):
            if all(k in col for k in keywords): return i
        return None

    # นิยามตำแหน่งคอลัมน์ที่อยู่
    addr_sets = [
        {"type": "ทะเบียนบ้าน", 
         "p": find_idx(["จังหวัด", "ทะเบียนบ้าน"]), 
         "d": find_idx(["อำเภอ", "ทะเบียนบ้าน"]), 
         "s": find_idx(["ตำบล", "ทะเบียนบ้าน"]), 
         "z": find_idx(["รหัสไปรษณีย์", "ทะเบียนบ้าน"])},
        {"type": "ติดต่อได้", 
         "p": find_idx(["จังหวัด", "ติดต่อได้"]), 
         "d": find_idx(["อำเภอ", "ติดต่อได้"]), 
         "s": find_idx(["ตำบล", "ติดต่อได้"]), 
         "z": find_idx(["รหัสไปรษณีย์", "ติดต่อได้"])}
    ]

    # วนลูปตรวจสอบ
    for r_idx in range(len(df_emp)):
        for c_idx in range(len(df_emp.columns)):
            val = df_emp.iloc[r_idx, c_idx]
            col_name = df_emp.columns[c_idx]
            val_str = str(val).strip()
            if val_str.lower() == 'nan': val_str = ""
            
            reason, color = "", ""

            # 3.1 เช็คค่าว่าง (Orange)
            if c_idx in ORANGE_INDICES and val_str == "":
                reason, color = "⚠️ ห้ามว่าง 'กรุณากรอกข้อมูล'", '#FFCC99'
            
            # 3.2 เช็คข้อมูลทั่วไป (ID/Date)
            elif val_str != "":
                if c_idx in DATE_COLS_IDX:
                    dt_obj, success = smart_date_parser(val)
                    if not success: reason, color = "❌ วันที่ผิดฟอร์แมต", '#FFC7CE'
                elif any(k in col_name for k in ["เลขบัตรประชาชน", "เลขประกันสังคม"]):
                    clean_id = re.sub(r'\D', '', val_str)
                    if len(clean_id) != 13: reason, color = "❌ ต้องมี 13 หลัก", '#FFC7CE'

            # 3.3 ระบบตรวจสอบที่อยู่แบบแยกเช็ค
            if MASTER_DB is not None:
                for ad in addr_sets:
                    if c_idx in [ad['p'], ad['d'], ad['s'], ad['z']]:
                        p_v = str(df_emp.iloc[r_idx, ad['p']]).strip()
                        d_v = str(df_emp.iloc[r_idx, ad['d']]).strip()
                        s_v = str(df_emp.iloc[r_idx, ad['s']]).strip()
                        z_v = str(df_emp.iloc[r_idx, ad['z']]).strip()

                        if c_idx == ad['p'] and p_v != "" and p_v not in MASTER_DB['province'].values:
                            reason, color = f"❌ ไม่พบจังหวัด {p_v}", '#FFC7CE'
                        
                        if c_idx == ad['d'] and d_v != "" and p_v != "":
                            match_d = MASTER_DB[(MASTER_DB['province'] == p_v) & (MASTER_DB['district'] == d_v)]
                            if match_d.empty:
                                reason, color = f"❌ อ.{d_v} ไม่ได้อยู่ใน {p_v}", '#FFC7CE'
                        
                        if c_idx == ad['s'] and s_v != "" and d_v != "" and p_v != "":
                            match_s = MASTER_DB[(MASTER_DB['province'] == p_v) & (MASTER_DB['district'] == d_v) & (MASTER_DB['subdistrict'] == s_v)]
                            if match_s.empty:
                                reason, color = f"❌ ต.{s_v} ข้อมูลไม่สัมพันธ์กับ อ./จ.", '#FFC7CE'
                        
                        if c_idx == ad['z'] and z_v != "" and s_v != "" and d_v != "" and p_v != "":
                            match_z = MASTER_DB[(MASTER_DB['province'] == p_v) & (MASTER_DB['district'] == d_v) & (MASTER_DB['subdistrict'] == s_v) & (MASTER_DB['zipcode'] == z_v)]
                            if match_z.empty:
                                reason, color = "❌ รหัสไปรษณีย์ไม่ถูกต้องตามพื้นที่", '#FFC7CE'

            if reason:
                error_details.append({"row": r_idx + 1, "col": c_idx, "reason": reason, "color": color, "col_name": col_name})

    # --- 4. สร้างไฟล์ Output (กู้คืน Memory) ---
    output = io.BytesIO()
    try:
        uploaded_file.seek(0)
        original_workbook = pd.read_excel(uploaded_file, sheet_name=None, dtype=object)
    except:
        original_workbook = {}

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_emp.replace(["nan", "NaN"], "", inplace=True)
        df_emp.to_excel(writer, index=False, sheet_name='พนักงาน')
        
        for sheet_name, df_content in original_workbook.items():
            if sheet_name != 'พนักงาน':
                df_content.replace(["nan", "NaN"], "", inplace=True)
                df_content.to_excel(writer, index=False, sheet_name=sheet_name)
        
        ws = writer.sheets['พนักงาน']
        workbook = writer.book
        text_fmt = workbook.add_format({'num_format': '@'})
        ws.set_column('A:ZZ', None, text_fmt)
        
        fmt_orange = workbook.add_format({'bg_color': '#FFCC99', 'border': 1})
        fmt_red = workbook.add_format({'bg_color': '#FFC7CE', 'border': 1})

        for err in error_details:
            f = fmt_orange if err['color'] == '#FFCC99' else fmt_red
            curr_val = df_emp.iloc[err['row']-1, err['col']]
            ws.write(err['row'], err['col'], str(curr_val) if str(curr_val).lower() != 'nan' else "", f)
            ws.write_comment(err['row'], err['col'], err['reason'])

    # เคลียร์ตัวแปรหนักๆ ออกจาก RAM
    del original_workbook
    del df_emp
    gc.collect() 
    
    return error_details, output.getvalue()

# --- 5. UI ---
uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel", type=["xlsx"])
if uploaded_file:
    try:
        errs, final_data = process_excel_data(uploaded_file)
        if final_data:
            if errs:
                st.error(f"🚩 พบจุดที่ต้องแก้ไข {len(errs)} รายการ")
                st.download_button("📥 ดาวน์โหลดไฟล์เพื่อทำการแก้ไข", data=final_data, file_name=f"Checked_Data.xlsx", use_container_width=True)
                st.dataframe(pd.DataFrame([{"แถว": e['row']+1, "คอลัมน์": e['col_name'], "สาเหตุ": e['reason']} for e in errs]))
            else:
                st.success("🎉 ข้อมูลถูกต้องทั้งหมด")
                st.download_button("📥 ดาวน์โหลดไฟล์ (ข้อมูลถูกต้อง)", data=final_data, file_name=f"Verified_Data.xlsx", use_container_width=True)
            
            # ล้างแรมหลังจบกระบวนการ
            del final_data
            gc.collect()
            
    except Exception as e:
        st.error(f"❌ Error: {e}")
else:
    st.info("💡 กรุณาวางไฟล์ เพื่อตรวจสอบข้อมูล")