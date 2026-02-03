import streamlit as st
import time
import random

# 1. ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Our Little Journey 🎀", page_icon="🧸")

# 2. ฟังก์ชันหัวใจลอยตอนจบ (ให้เยอะขึ้นและหวานขึ้น)
def heart_animation():
    st.markdown("""
        <style>
        @keyframes float {
            0% { transform: translateY(0) rotate(0deg); opacity: 1; }
            100% { transform: translateY(-100vh) rotate(360deg); opacity: 0; }
        }
        .heart-particle {
            position: fixed;
            bottom: -10px;
            color: #ffb3c1;
            font-size: 20px;
            user-select: none;
            pointer-events: none;
            z-index: 9999;
            animation: float 4s linear forwards;
        }
        </style>
    """, unsafe_allow_html=True)
    
    hearts_html = ""
    for i in range(40):
        left = random.randint(0, 100)
        delay = random.uniform(0, 3)
        duration = random.uniform(4, 7)
        size = random.randint(20, 40)
        hearts_html += f'<div class="heart-particle" style="left:{left}vw; animation-delay:{delay}s; animation-duration:{duration}s; font-size:{size}px;">🌸</div>'
        hearts_html += f'<div class="heart-particle" style="left:{left}vw; animation-delay:{delay+0.5}s; animation-duration:{duration}s; font-size:{size-5}px;">💖</div>'
    
    st.markdown(hearts_html, unsafe_allow_html=True)

# 3. CSS ธีมชมพูคลีนๆ แบบสาวหวาน
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Kanit:wght@300&display=swap');
    
    .stApp { background-color: #fff0f3; }
    
    html, body, [class*="css"] { font-family: 'Kanit', sans-serif; }

    .question-box {
        background: rgba(255, 255, 255, 0.8);
        padding: 30px; 
        border-radius: 30px;
        box-shadow: 0 10px 30px rgba(255, 179, 193, 0.3);
        margin-bottom: 25px; 
        text-align: center; 
        border: 2px solid #ffcad4;
    }

    .stButton>button {
        border-radius: 50px; 
        border: none;
        background: linear-gradient(to right, #ff85a1, #ffb3c1);
        color: white;
        font-size: 18px; 
        padding: 10px 20px;
        transition: 0.5s; 
        box-shadow: 0 5px 15px rgba(255, 133, 161, 0.4);
    }
    
    .stButton>button:hover {
        transform: translateY(-3px);
        box-shadow: 0 8px 20px rgba(255, 133, 161, 0.6);
        color: white;
    }

    h1 { color: #ff4d6d !important; text-align: center; font-size: 2.5rem; }
    h3 { color: #800f2f !important; }
    </style>
    """, unsafe_allow_html=True)

# 4. ระบบจัดการด่าน
if 'step' not in st.session_state: st.session_state.step = 1
def next_step(): st.session_state.step += 1

# --- เนื้อหาเกม ---

st.markdown("<h1>Entering Valentine's Day 🧸</h1>", unsafe_allow_html=True)
st.write("---")

if st.session_state.step == 1:
    st.markdown("<div class='question-box'><h3>ก่อนจะเริ่มภารกิจ...<br>มิทราบว่า คุณพลรัตน์ พร้อมรึยังคะ? 😊</h3></div>", unsafe_allow_html=True)
    if st.button("พร้อมแล้ววว🎉"):
        next_step()
        st.rerun()

elif st.session_state.step == 2:
    st.markdown("<div class='question-box'><h3>ด่านที่ 1: วันพักผ่อน ☁️</h3><p>ถ้าเรามีวันว่างหนึ่งวัน คุณอยากชวนเค้าไปทำอะไรมากที่สุด?</p></div>", unsafe_allow_html=True)
    q1 = st.radio("", [
        "ไปเดินสวนสาธารณะตอนเย็นๆ ลมเย็นๆ", 
        "นั่งดู Netflix แล้วสั่งของอร่อยมาจอยกัน", 
        "ออกไปตะลุยคาเฟ่ หาที่ถ่ายรูปสวยๆ", 
        "แค่เดินห้างโง่ๆ ไปด้วยกันก็โอเคแล้ว"
    ])
    if st.button("บันทึกคำตอบ ✨"):
        next_step()
        st.rerun()

elif st.session_state.step == 3:
    st.markdown("<div class='question-box'><h3>ด่านที่ 2: ความในใจ 💌</h3><p>มีใจให้กันบ้างรอ๊ะยัง?</p></div>", unsafe_allow_html=True)
    q3 = st.radio("", [
        "มีใจให้ตั้งนานแล้วครับ ไม่รู้ตัวเหรอ?", 
        "ก็เริ่มจะใจสั่นๆ เวลาเห็นแจ้งเตือนคุณแล้วนะ", 
        "ขอเก็บไว้ลุ้นต่ออีกนิด แต่คะแนนนำโด่งเลยล่ะ",
        "มากกว่ามีใจ... คืออยากมีคุณอยู่ข้างๆ แล้วครับ"
    ])
    if st.button("ไปด่านสุดท้ายกัน 🧸"):
        next_step()
        st.rerun()

elif st.session_state.step == 4:
    st.markdown("<div class='question-box'><h3>ด่านสุดท้าย: ไหนบอกซื้ 💋</h3><p>วาเลนไทน์ปีนี้เป็นคนคุยไปก่อน... แล้ววาเลนไทน์ปีหน้า 'สถานะ' ของเราควรเป็นอะไรดีคะ?</p></div>", unsafe_allow_html=True)
    q2 = st.radio("", [
        "ก็ต้องเป็น 'คนของกันและกัน' สิครับ", 
        "แฟนครับ... พิมพ์รอไว้ล่วงหน้าเลยได้มั้ย?", 
        "เลื่อนขั้นมาเป็นคนดูแลหัวใจกันแบบยาวๆ เลยนะ",
        "ทุกข้อที่กล่าวมาครับ อยู่ที่ว่าคุณจะตกลงเมื่อไหร่"
    ])
    if st.button("รับของขวัญวาเลนไทน์ 🎁"):
        next_step()
        st.rerun()

elif st.session_state.step == 5:
    heart_animation()
    st.markdown("<div class='question-box'><h2>Happy Valentine's Day ✨</h2></div>", unsafe_allow_html=True)
    st.markdown(f"""
    <div style='text-align: center; color: #590d22;'>
        <p style='font-size: 1.2rem;'>ขอบคุณที่ยังไม่เบื่อความดื้อของเค้านะ</p>
        <p>ไม่รู้ว่าจะเจอคนดื้อมาเยอะแค่ไหน แต่คนดื้อคนนี้กะจะอยู่เป็นความปวดหัวให้ไปอีกนาน</p>
        <p>วาเลนไทน์ปีนี้ไม่ต้องมีอะไรหวือหวา แค่เทอยังอยู่ตรงนี้ก็น่ารักที่สุดแล้ว</p>
        <h2 style='color: #ff4d6d;'>Happy Valentine’s Day! ทำตัวน่ารักแบบนี้ไปนานๆล่ะ ❤️</h2>
    </div>
    """, unsafe_allow_html=True)
    
    if st.button("เล่นอีกรอบ 🔄"):
        st.session_state.step = 1
        st.rerun()

st.write("---")
st.caption("By Radathip Peansawanglap 🌸")