import streamlit as st
import pandas as pd
import io
import re
import gc
from datetime import datetime

# --- 1. การตั้งค่าหน้าเว็บ ---
st.set_page_config(page_title="Employee Data Validator Pro", layout="wide")

# ฟังก์ชันโหลด Database จากไฟล์ CSV ของคุณ
@st.cache_data
def load_master_db_from_csv():
    file_path = "DataDaseจังหวัด.csv"
    try:
        # อ่านไฟล์ CSV แบบไม่มี Header เพราะไฟล์ของคุณเริ่มด้วยข้อมูลเลย
        db = pd.read_csv(file_path, dtype=str, encoding='utf-8-sig', header=None)
        
        # เลือกคอลัมน์ตามโครงสร้างไฟล์ CSV ที่อัปโหลดมา
        # 0=ไปรษณีย์, 3=ตำบล, 4=อำเภอ, 5=จังหวัด
        db_clean = pd.DataFrame({
            'zipcode': db[0],
            'subdistrict': db[3],
            'district': db[4],
            'province': db[5]
        })
        return db_clean.apply(lambda x: x.str.strip())
    except Exception as e:
        # ถ้าเปิดด้วย utf-8 ไม่ได้ ให้ลอง tis-620 (ภาษาไทยแบบเก่า)
        try:
            db = pd.read_csv(file_path, dtype=str, encoding='tis-620', header=None)
            db_clean = pd.DataFrame({'zipcode': db[0], 'subdistrict': db[3], 'district': db[4], 'province': db[5]})
            return db_clean.apply(lambda x: x.str.strip())
        except:
            st.error(f"❌ ไม่สามารถโหลด Database ได้: กรุณาวางไฟล์ '{file_path}' ไว้ที่เดียวกับโค้ด")
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

# --- 2. การตั้งค่าคอลัมน์และฟังก์ชันช่วย (โค้ดเดิม) ---
ORANGE_INDICES = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15, 16, 17, 20, 21, 23, 24, 25, 39, 40, 41, 64, 65]
MAP_COLS_BASIC = ["คำนำหน้า", "เพศ", "ระดับ", "ตำแหน่ง", "บริษัท", "สายงาน", "ฝ่าย", "แผนก", "สถานะพนักงาน", "ประเภทการจ้างงาน"]
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

# --- 3. ฟังก์ชันหลักในการประมวลผลไฟล์ ---
def process_excel_data(uploaded_file):
    uploaded_file.seek(0)
    df_emp = pd.read_excel(uploaded_file, sheet_name="พนักงาน", dtype=object)
    
    # อ่านชีทรายละเอียดเพื่อใช้เช็คค่าพื้นฐาน
    try:
        uploaded_file.seek(0)
        df_ref = pd.read_excel(uploaded_file, sheet_name="รายละเอียด (ห้ามแก้ไข)", dtype=object)
        ref_data_basic = {col: set(df_ref[col].dropna().astype(str).str.strip().unique()) 
                          for col in MAP_COLS_BASIC if col in df_ref.columns}
    except:
        df_ref = pd.DataFrame()
        ref_data_basic = {}

    df_emp.columns = [str(c).replace('\n', ' ').strip() for c in df_emp.columns]
    error_details = []

    def find_col_idx(df, keywords):
        for i, col in enumerate(df.columns):
            if all(k in col for k in keywords): return i
        return None

    # หา Index AY-BH
    ay_idx = find_col_idx(df_emp, ["ที่อยู่", "ทะเบียนบ้าน"])
    az_idx = find_col_idx(df_emp, ["จังหวัด", "ทะเบียนบ้าน"])
    ba_idx = find_col_idx(df_emp, ["อำเภอ", "ทะเบียนบ้าน"])
    bb_idx = find_col_idx(df_emp, ["ตำบล", "ทะเบียนบ้าน"])
    bc_idx = find_col_idx(df_emp, ["รหัสไปรษณีย์", "ทะเบียนบ้าน"])
    
    bd_idx = find_col_idx(df_emp, ["ที่อยู่", "ติดต่อได้"])
    be_idx = find_col_idx(df_emp, ["จังหวัด", "ติดต่อได้"])
    bf_idx = find_col_idx(df_emp, ["อำเภอ", "ติดต่อได้"])
    bg_idx = find_col_idx(df_emp, ["ตำบล", "ติดต่อได้"])
    bh_idx = find_col_idx(df_emp, ["รหัสไปรษณีย์", "ติดต่อได้"])

    for r_idx in range(len(df_emp)):
        for c_idx in range(len(df_emp.columns)):
            val = df_emp.iloc[r_idx, c_idx]
            col_name = df_emp.columns[c_idx]
            val_str = str(val).strip() if pd.notna(val) else ""
            reason, color = "", ""

            # 3.1 เช็คค่าว่าง (Orange)
            if c_idx in ORANGE_INDICES and (val_str == "" or val_str.lower() == 'nan'):
                reason, color = "⚠️ ห้ามว่าง", '#FFCC99'
            
            # 3.2 เช็คข้อมูลผิดพลาด (Red)
            elif val_str != "" and val_str.lower() != 'nan':
                if c_idx in DATE_COLS_IDX:
                    dt_obj, success = smart_date_parser(val)
                    if success: df_emp.iloc[r_idx, c_idx] = dt_obj
                    else: reason, color = "❌ วันที่ผิดฟอร์แมต", '#FFC7CE'
                
                elif any(k in col_name for k in ["เลขบัตรประชาชน", "เลขประกันสังคม"]):
                    clean_id = re.sub(r'\D', '', val_str)
                    df_emp.iloc[r_idx, c_idx] = clean_id
                    if len(clean_id) != 13: reason, color = "❌ ต้องมี 13 หลัก", '#FFC7CE'

                # --- 3.3 การตรวจสอบที่อยู่ขัดแย้ง (เชื่อม CSV MASTER_DB) ---
                is_addr_col = c_idx in [ay_idx, az_idx, ba_idx, bb_idx, bc_idx, bd_idx, be_idx, bf_idx, bg_idx, bh_idx]
                if is_addr_col and MASTER_DB is not None:
                    is_reg = c_idx in [ay_idx, az_idx, ba_idx, bb_idx, bc_idx]
                    p_i, d_i, s_i, z_i = (az_idx, ba_idx, bb_idx, bc_idx) if is_reg else (be_idx, bf_idx, bg_idx, bh_idx)
                    
                    if all(idx is not None for idx in [p_i, d_i, s_i, z_i]):
                        p_v = str(df_emp.iloc[r_idx, p_i]).strip()
                        d_v = str(df_emp.iloc[r_idx, d_i]).strip()
                        s_v = str(df_emp.iloc[r_idx, s_i]).strip()
                        z_v = str(df_emp.iloc[r_idx, z_i]).strip()

                        if p_v != "" and p_v != 'nan':
                            db_match = MASTER_DB[MASTER_DB['province'] == p_v]
                            if db_match.empty:
                                if c_idx == p_i: reason, color = f"❌ ไม่พบจังหวัด {p_v}", '#FFC7CE'
                            else:
                                if c_idx == d_i and d_v != "" and d_v not in db_match['district'].values:
                                    reason, color = f"❌ อ.{d_v} ไม่ได้อยู่ใน {p_v}", '#FFC7CE'
                                if c_idx == s_i and s_v != "" and s_v not in db_match['subdistrict'].values:
                                    reason, color = f"❌ ต.{s_v} ไม่ได้อยู่ใน {p_v}", '#FFC7CE'
                                if c_idx == z_i and z_v != "" and z_v not in db_match['zipcode'].values:
                                    reason, color = "❌ รหัสไปรษณีย์ไม่ตรงพื้นที่", '#FFC7CE'

            if reason:
                error_details.append({"row": r_idx + 1, "col": c_idx, "reason": reason, "color": color, "col_name": col_name})

    # --- 4. การสร้างไฟล์ Output สำหรับดาวน์โหลด ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_emp.to_excel(writer, index=False, sheet_name='พนักงาน')
        if not df_ref.empty:
            df_ref.to_excel(writer, index=False, sheet_name='รายละเอียด (ห้ามแก้ไข)')
            
        ws = writer.sheets['พนักงาน']
        workbook = writer.book
        text_fmt = workbook.add_format({'num_format': '@'})
        ws.set_column('A:ZZ', None, text_fmt)
        
        fmt_orange = workbook.add_format({'bg_color': '#FFCC99', 'border': 1})
        fmt_red = workbook.add_format({'bg_color': '#FFC7CE', 'border': 1})

        for err in error_details:
            f = fmt_orange if err['color'] == '#FFCC99' else fmt_red
            ws.write(err['row'], err['col'], str(df_emp.iloc[err['row']-1, err['col']]), f)
            ws.write_comment(err['row'], err['col'], err['reason'])

    return error_details, output.getvalue()

# --- 5. ส่วน UI ---
uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel เพื่อเริ่มตรวจสอบ", type=["xlsx"])
if uploaded_file:
    try:
        error_details, final_data = process_excel_data(uploaded_file)
        if error_details:
            st.error(f"🚩 พบจุดที่ต้องแก้ไข {len(error_details)} รายการ")
            st.download_button("📥 ดาวน์โหลดไฟล์เพื่อทำการแก้ไข", data=final_data, file_name=f"Recheck_{datetime.now().strftime('%H%M%S')}.xlsx", use_container_width=True)
            st.dataframe(pd.DataFrame([{"แถว": e['row']+1, "คอลัมน์": e['col_name'], "สาเหตุ": e['reason']} for e in error_details]))
        else:
            st.success("🎉 ข้อมูลถูกต้องทั้งหมด")
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
else:
    st.info("💡 กรุณาวางไฟล์ เพื่อตรวจสอบข้อมูล")