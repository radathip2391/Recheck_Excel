import streamlit as st
import pandas as pd
import io
import re
import gc
from datetime import datetime

# --- 1. การตั้งค่าหน้าเว็บ (Premium Theme) ---
st.set_page_config(
    page_title="Data Validator Pro", 
    page_icon="🛡️",
    layout="wide"
)

# Custom CSS สำหรับดีไซน์หรูหราแบบ Modern SaaS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&family=Sarabun:wght@300;400;600&display=swap');

    .dev-credit {
        position: absolute;
        top: 20px;
        right: 30px;
        font-size: 0.85rem;
        color: #94a3b8;
        font-family: 'Plus Jakarta Sans', sans-serif;
        font-weight: 500;
        opacity: 0.7;
        z-index: 100;
    }
    
    .stApp {
        background: #f8faff;
        background-image: 
            radial-gradient(at 0% 0%, rgba(168, 85, 247, 0.15) 0px, transparent 50%), 
            radial-gradient(at 100% 0%, rgba(59, 130, 246, 0.15) 0px, transparent 50%),
            radial-gradient(at 50% 100%, rgba(236, 72, 153, 0.1) 0px, transparent 50%);
        font-family: 'Plus Jakarta Sans', 'Sarabun', sans-serif;
    }

    .hero-section {
        background: rgba(255, 255, 255, 0.3);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border: 1px solid rgba(255, 255, 255, 0.6);
        border-radius: 32px;
        padding: 4rem 2rem;
        box-shadow: 0 20px 50px rgba(168, 85, 247, 0.1);
        position: relative;
        overflow: hidden;
    }

    .hero-section::before {
        content: "";
        position: absolute;
        top: 0; left: 0; right: 0; height: 4px;
        background: linear-gradient(90deg, #818cf8, #c084fc, #f472b6);
    }

    .hero-title {
        background: linear-gradient(135deg, #6366f1 0%, #d946ef 50%, #3b82f6 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 3.5rem;
        font-weight: 800;
        filter: drop-shadow(0 2px 10px rgba(217, 70, 239, 0.2));
    }

    .status-card {
        background: rgba(255, 255, 255, 0.8);
        border-radius: 24px;
        border: 1px solid rgba(168, 85, 247, 0.1);
        padding: 1.5rem;
        box-shadow: 0 10px 25px rgba(168, 85, 247, 0.05);
        transition: all 0.4s ease;
    }

    div.stDownloadButton > button {
        background: linear-gradient(135deg, #818cf8 0%, #f472b6 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 20px !important;
        padding: 1rem !important;
        font-weight: 700 !important;
        box-shadow: 0 10px 25px rgba(244, 114, 182, 0.4) !important;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. Logic การคำนวณ ---

@st.cache_data
def load_master_db_from_csv():
    file_path = "DataBaseจังหวัด.csv"
    for enc in ['utf-8-sig', 'tis-620']:
        try:
            db = pd.read_csv(file_path, dtype=str, encoding=enc, header=None)
            db_clean = pd.DataFrame({
                'zipcode': db[0], 'subdistrict': db[1], 'district': db[7], 'province': db[10]
            }).apply(lambda x: x.str.strip())
            return db_clean
        except: continue
    return None

MASTER_DB = load_master_db_from_csv()
ORANGE_INDICES = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15, 16, 17, 20, 21, 23, 24, 25, 39, 40, 41, 64, 65]
DATE_COLS_IDX = [1, 25]
MAP_COLS = ["คำนำหน้า", "เพศ", "ระดับ", "ตำแหน่ง", "บริษัท", "สายงาน", "ฝ่าย", "แผนก", "สถานะพนักงาน", "ประเภทการจ้างงาน"]

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

def process_excel_data(uploaded_file):
    uploaded_file.seek(0)
    try:
        # อ่านข้อมูลเป็น string เพื่อป้องกัน format เพี้ยน
        df_emp = pd.read_excel(uploaded_file, sheet_name="พนักงาน", dtype=str).fillna("")
        df_ref = pd.read_excel(uploaded_file, sheet_name="รายละเอียด (ห้ามแก้ไข)", dtype=str)
        ref_data = {col: set(df_ref[col].dropna().astype(str).str.strip().unique()) 
                   for col in MAP_COLS if col in df_ref.columns}
    except Exception as e:
        return f"ไม่พบชีท 'พนักงาน': {e}", None

    df_emp.columns = [str(c).replace('\n', ' ').strip() for c in df_emp.columns]
    error_details = []

    def find_idx(keywords):
        for i, col in enumerate(df_emp.columns):
            if all(k in col for k in keywords): return i
        return None

    addr_sets = [
        {"type": "ทะเบียนบ้าน", "p": find_idx(["จังหวัด", "ทะเบียนบ้าน"]), "d": find_idx(["อำเภอ", "ทะเบียนบ้าน"]), "s": find_idx(["ตำบล", "ทะเบียนบ้าน"]), "z": find_idx(["รหัสไปรษณีย์", "ทะเบียนบ้าน"])},
        {"type": "ติดต่อได้", "p": find_idx(["จังหวัด", "ติดต่อได้"]), "d": find_idx(["อำเภอ", "ติดต่อได้"]), "s": find_idx(["ตำบล", "ติดต่อได้"]), "z": find_idx(["รหัสไปรษณีย์", "ติดต่อได้"])}
    ]

    for r_idx in range(len(df_emp)):
        for c_idx in range(len(df_emp.columns)):
            val = str(df_emp.iloc[r_idx, c_idx]).strip()
            col_name = df_emp.columns[c_idx]
            if val.lower() == 'nan': val = ""
            reason, color = "", ""

            # --- [🔥 จุดที่เพิ่ม: ลบเครื่องหมายขีด (-) และคลีนตัวเลข] ---
            target_cols = ["เลขบัตรประชาชน", "เลขประกันสังคม", "เลขบัญชีธนาคาร", "โทรศัพท์", "เบอร์โทร"]
            if any(k in col_name for k in target_cols) and val != "":
                clean_id = re.sub(r'\D', '', val) # ลบทุกอย่างที่ไม่ใช่ตัวเลข
                df_emp.iloc[r_idx, c_idx] = clean_id # เขียนค่าที่ลบขีดแล้วกลับลงไป
                
                # ตรวจสอบหลังลบขีดแล้ว
                if any(k in col_name for k in ["เลขบัตรประชาชน", "เลขประกันสังคม"]):
                    if len(clean_id) != 13: reason, color = "❌ ไม่ครบ 13 หลัก", '#FFC7CE'
                elif col_name == "เลขบัญชีธนาคาร":
                    if len(clean_id) != 10: reason, color = "❌ ไม่ครบ 10 หลัก", '#FFC7CE'

            elif val != "":
                if c_idx in DATE_COLS_IDX:
                    dt_obj, success = smart_date_parser(val)
                    if not success: reason, color = "❌ วันที่ผิดฟอร์แมต", '#FFC7CE'
                    else:
                        df_emp.iloc[r_idx, c_idx] = dt_obj.strftime('%d/%m/%Y') # บังคับ format วันที่
                elif col_name in ref_data:
                    if val not in ref_data[col_name]: reason, color = "❌ ไม่ตรงฐานข้อมูลกลาง", '#FFC7CE'
            
            if c_idx in ORANGE_INDICES and val == "":
                reason, color = "⚠️ ห้ามว่าง", '#FFCC99'

            if MASTER_DB is not None:
                for ad in addr_sets:
                    if ad['p'] is None: continue 
                    if c_idx in [ad['p'], ad['d'], ad['s'], ad['z']]:
                        p_v, d_v, s_v, z_v = [str(df_emp.iloc[r_idx, ad[k]]).strip() for k in ['p', 'd', 's', 'z']]
                        if c_idx == ad['p'] and p_v != "" and p_v not in MASTER_DB['province'].values:
                            reason, color = f"❌ ไม่พบจังหวัด {p_v}", '#FFC7CE'
                        elif c_idx == ad['d'] and d_v != "" and p_v != "":
                            if MASTER_DB[(MASTER_DB['province'] == p_v) & (MASTER_DB['district'] == d_v)].empty:
                                reason, color = f"❌ อ.{d_v} ไม่อยู่ใน {p_v}", '#FFC7CE'
                        elif c_idx == ad['s'] and s_v != "" and d_v != "" and p_v != "":
                            if MASTER_DB[(MASTER_DB['province'] == p_v) & (MASTER_DB['district'] == d_v) & (MASTER_DB['subdistrict'] == s_v)].empty:
                                reason, color = f"❌ ต.{s_v} ข้อมูลไม่สัมพันธ์", '#FFC7CE'
                        elif c_idx == ad['z'] and z_v != "" and s_v != "" and d_v != "" and p_v != "":
                            if MASTER_DB[(MASTER_DB['province'] == p_v) & (MASTER_DB['district'] == d_v) & (MASTER_DB['subdistrict'] == s_v) & (MASTER_DB['zipcode'] == z_v)].empty:
                                reason, color = "❌ รหัสไปรษณีย์ผิด", '#FFC7CE'
            
            if reason:
                error_details.append({"row": r_idx + 1, "col": c_idx, "reason": reason, "color": color, "col_name": col_name})

    # --- ส่วนการเขียนไฟล์ Output ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_emp.to_excel(writer, index=False, sheet_name='พนักงาน')
        # เขียนชีทอื่นคืน
        uploaded_file.seek(0)
        orig_dict = pd.read_excel(uploaded_file, sheet_name=None, dtype=str)
        for sn, content in orig_dict.items():
            if sn != 'พนักงาน': content.to_excel(writer, index=False, sheet_name=sn)
        
        ws = writer.sheets['พนักงาน']
        workbook = writer.book
        text_fmt = workbook.add_format({'num_format': '@'}) # บังคับให้เลข 0 ตัวหน้าไม่หาย
        ws.set_column('A:ZZ', None, text_fmt)
        
        fmt_orange = workbook.add_format({'bg_color': '#FFCC99', 'border': 1, 'num_format': '@'})
        fmt_red = workbook.add_format({'bg_color': '#FFC7CE', 'border': 1, 'num_format': '@'})

        for err in error_details:
            f = fmt_orange if err['color'] == '#FFCC99' else fmt_red
            curr_val = str(df_emp.iloc[err['row']-1, err['col']])
            ws.write(err['row'], err['col'], curr_val, f)
            ws.write_comment(err['row'], err['col'], err['reason'])

    gc.collect()
    return error_details, output.getvalue()

# --- 3. UI หน้าจอหลัก ---

st.markdown("""
    <div class="hero-section">
        <div class="dev-credit">Developed by Intern X SC</div>
        <div class="hero-title">Employee Data Validator</div>
        <p style="color: #94a3b8; font-size: 1.1rem;">อัปโหลดและตรวจสอบความถูกต้องของข้อมูลพนักงาน</p>
    </div>
    """, unsafe_allow_html=True)

top_col1, top_col2 = st.columns([1.5, 1], gap="large")

with top_col1:
    st.markdown("### 📤 ขั้นตอนที่ 1: อัปโหลดไฟล์")
    uploaded_file = st.file_uploader("ลากไฟล์ Excel มาวางที่นี่", type=["xlsx"])

with top_col2:
    st.markdown("### 🏷️ คำอธิบายสี")
    st.markdown("""
        <div class="status-card">
            <span class="indicator-tag" style="background-color: #FFCC99; color: #92400e;">🟠 สีส้ม</span> ช่องว่างที่จำเป็นต้องระบุ<br>
            <div style="margin-top: 8px;"></div>
            <span class="indicator-tag" style="background-color: #FFC7CE; color: #991b1b;">🔴 สีแดง</span> ข้อมูลผิดรูปแบบ / ไม่ตรงฐานข้อมูล
        </div>
    """, unsafe_allow_html=True)

if uploaded_file:
    with st.status("🛠️ กำลังวิเคราะห์และลบเครื่องหมายขีด...", expanded=True) as status:
        errs, final_data = process_excel_data(uploaded_file)
        status.update(label="✅ ตรวจสอบและคลีนข้อมูลเรียบร้อยแล้ว", state="complete", expanded=False)

    st.markdown("---")
    
    if final_data:
        if isinstance(errs, list) and len(errs) > 0:
            res_left, res_right = st.columns([1.2, 0.8])
            with res_left:
                st.markdown(f"### 📋 รายงานผลการตรวจสอบ")
                st.warning(f"พบรายการที่ต้องแก้ไขทั้งหมด **{len(errs)}** แห่ง กรุณาตรวจสอบในไฟล์ Excel")
                err_df = pd.DataFrame([{"ลำดับ": i+1, "แถว": e['row']+1, "ชื่อคอลัมน์": e['col_name'], "รายละเอียด": e['reason']} for i, e in enumerate(errs)])
                st.dataframe(err_df, use_container_width=True, hide_index=True)
            with res_right:
                st.markdown("### 📥 ดาวน์โหลด")
                st.write("ระบบทำการไฮไลต์สีและใส่คำแนะนำ (Comment) ไว้ในไฟล์ให้เรียบร้อยแล้ว")
                st.download_button(label="ดาวน์โหลดไฟล์ที่แก้ไขแล้ว", data=final_data, file_name=f"Cleaned_EmployeeData_{datetime.now().strftime('%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            if isinstance(errs, str): st.error(errs)
            else:
                st.balloons()
                st.success("🎉 ข้อมูลถูกต้องและลบขีดเรียบร้อยแล้ว!")
                st.download_button(label="ดาวน์โหลดไฟล์ที่ตรวจสอบแล้ว", data=final_data, file_name="Verified_EmployeeData.xlsx", use_container_width=True)
else:
    st.info("👋 กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มการตรวจสอบข้อมูล")