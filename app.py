import streamlit as st
import pandas as pd
from datetime import datetime
import os

DATA_FILE = "responses.csv"
LOGO_PATH = "assets/msc_logo.png"

if "submitted" not in st.session_state:
    st.session_state["submitted"] = False

st.set_page_config(
    page_title="MSC Latvia – Employee Wellbeing Survey",
    page_icon=None,
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Archivo:wght@400;900&display=swap');

    /* Global Body */
    html, body, [class*="st-"], .stApp {
        font-family: 'Archivo', sans-serif !important;
        color: #8C7F72;
    }

    /* TITLE */
    h1 {
        font-family: 'Archivo', sans-serif !important;
        font-weight: 900 !important;
        color: #8C7F72 !important;
        text-align: center;
        text-transform: uppercase;
        letter-spacing: 1px;
        font-size: 2.2rem !important;
    }

    .subtitle {
        text-align: center;
        color: #8C7F72;
        font-size: 0.85rem;
        text-transform: uppercase;
        letter-spacing: 3px;
        margin-bottom: 50px;
    }

    /* QUESTION BOX (MSC Grey) */
    .question-box {
        background-color: #8C7F72;
        padding: 22px 30px;
        border-radius: 0px;
        margin-top: 40px;
        margin-bottom: 5px;
    }

    .question-text {
        color: white !important;
        font-weight: 900 !important;
        font-size: 1.1rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin: 0;
    }

    .q-number {
        color: #F8DE8D;
        font-weight: 900;
        margin-right: 15px;
    }

    /* --- SLIDER STYLING --- */

    /* Pilnībā paslēpt visus ciparus iekš slidera (tooltip + 1/10) */
    .stSlider span {
        display: none !important;
        opacity: 0 !important;
        color: transparent !important;
    }

    /* Konteiners */
    .stSlider [data-baseweb="slider"] {
        margin-top: 10px;
        margin-bottom: 10px;
        background: transparent !important;
    }

    /* Viena plāna līnija – MSC pelēkā */
    .stSlider [data-baseweb="slider"] > div > div > div {
        background: #8C7F72 !important;      /* MSC Warm Grey */
        height: 40px !important;
        border-radius: 20px;
        margin: 20px #8C7F72;
    }

    .stSlider [data-baseweb="slider"] > div > div > div > div:hover {
        box-shadow: #F8DE8D80 0 0 0 0.2rem;      /* MSC Yellow */
    }

    .stSlider [data-baseweb="slider"] > div > div > div > div {
        background: #F8DE8D !important;      /* MSC Yellow */
        width: 20px;
        height: 20px;
    }

    .stSlider [data-baseweb="slider"] > div > div > div > div > div {
        background: transparent !important;      /* MSC Warm Grey */
        margin: -10px;
        display: inline;
        font-weight: 900 !important;
        font-size: 1.1rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        color: #8C7F72 !important;
    }

    .stElementContainer {
        width: 100%;
    }

    p, div{
        font-family: 'Archivo', sans-serif !important;
        font-weight: 900 !important;
    }

    /* Bumbiņa – MSC Yellow */
    .stSlider [data-baseweb="thumb"] {
        background-color: #F8DE8D !important; /* MSC Yellow */
        border: 0 !important;
        height: 20px !important;
        width: 20px !important;
        border-radius: 50% !important;
        box-shadow: none !important;
    }

    /* Apakšējais cipars – vairs neizmantojam, bet klase paliek, ja nu vajag nākotnē */
    .value-label {
        display: none !important;
    }

    /* SUBMIT BUTTON */
    .stButton {
        display: flex;
        justify-content: center;
        margin-top: 60px;
        margin-bottom: 80px;
    }

    .stButton > button {
        background-color: transparent !important;
        color: #8C7F72 !important;
        font-family: 'Archivo', sans-serif !important;
        font-weight: 900 !important;
        border: 2px solid #F8DE8D !important;
        border-radius: 0px !important;
        padding: 12px 70px !important;
        text-transform: uppercase;
        letter-spacing: 2px;
        transition: 0.3s;
    }

    .stButton > button:hover {
        background-color: #F8DE8D !important;
        border-color: #F8DE8D !important;
    }

    /* THANK YOU FULL-SCREEN FEEL */
    .thank-you-area {
        text-align: center;
        margin-top: 80px;
        padding: 80px 20px;
        border: 2px solid #F8DE8D;
    }

    .thank-you-title {
        color: #8C7F72 !important;
        font-family: 'Archivo', sans-serif !important;
        font-weight: 900 !important;
        font-size: 2.2rem;
        text-transform: uppercase;
        margin-bottom: 10px;
    }

    /* Error message styling */
    .required-error {
        color: #b00020;
        font-weight: 900;
        text-align: center;
        margin-top: 10px;
    }
</style>
""", unsafe_allow_html=True)

# 4. LOGO
col1, col2, col3 = st.columns([1, 1, 1])
with col2:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=160)
    else:
        st.error("Logo missing")

st.title("Wellbeing Survey")
st.markdown('<p class="subtitle">MSC Latvia Internal Feedback</p>', unsafe_allow_html=True)

if st.session_state["submitted"]:
    st.markdown("""
        <div class="thank-you-area">
            <h2 class="thank-you-title">Thank you for your answers!</h2>
            <p style="letter-spacing: 3px; text-transform: uppercase; font-size: 0.9rem; font-weight:900;">
                Have a great day.
            </p>
        </div>
    """, unsafe_allow_html=True)
    st.stop()

def load_data():
    if os.path.exists(DATA_FILE):
        return pd.read_csv(DATA_FILE)
    return pd.DataFrame()

def save_response(row: dict):
    df = load_data()
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    df.to_csv(DATA_FILE, index=False)

def refined_question(number, text, key):
    st.markdown(f"""
        <div class="question-box">
            <p class="question-text"><span class="q-number">{number:02d}</span> {text}</p>
        </div>
    """, unsafe_allow_html=True)

   
    st.markdown(
        "<div style='display:flex; justify-content:space-between; font-size:0.75rem; "
        "color:#8C7F72; font-weight:900; letter-spacing:1px; margin-top:10px;'>"
        "<span>STRONGLY DISAGREE</span>"
        "<span>STRONGLY AGREE</span>"
        "</div>",
        unsafe_allow_html=True
    )

    val = st.slider(text, 1, 10, 5, key=key, label_visibility="collapsed")

    return val

dept_placeholder = "Select department..."
dept = st.selectbox(
    "Department Selection",
    [
        dept_placeholder,
        "Administration", "Customer Invoicing", "Finance & Accounting",
        "Commercial Reporting & BI", "Information Technology", "OVA",
        "Documentation, Pricing & Legal"
    ],
    index=0
)

st.write("---")

q1 = refined_question(1, "I am satisfied with my current workload.", "q1")
q2 = refined_question(2, "After a workday, I have enough mental energy for the rest of my day.", "q2")
q3 = refined_question(3, "I am able to disconnect from work outside working hours.", "q3")
q4 = refined_question(4, "I feel comfortable sharing my opinions within my team.", "q4")
q5 = refined_question(5, "My direct manager supports my wellbeing.", "q5")
q6 = refined_question(6, "There is effective collaboration within my department.", "q6")
q7 = refined_question(7, "I feel motivated in my daily work.", "q7")
q8 = refined_question(8, "My work has a positive impact on my mental wellbeing.", "q8")
q9 = refined_question(9, "I feel physically well during working hours.", "q9")
q10 = refined_question(10, "I can maintain a healthy balance between work and personal life.", "q10")
q11 = refined_question(11, "My working schedule is flexible enough for my personal needs.", "q11")
q12 = refined_question(12, "I am satisfied with the remote or hybrid work options available to me.", "q12")
q13 = refined_question(13, "I see clear opportunities for professional growth at MSC Latvia.", "q13")
q14 = refined_question(14, "I feel valued and recognized for my contributions.", "q14")
q15 = refined_question(15, "Overall, I am satisfied working at MSC Latvia.", "q15")

error_msg = st.empty()

if st.button("Submit Survey"):
    # jāpārbauda, vai izvēlēta nodaļa
    if dept == dept_placeholder:
        error_msg.markdown(
            '<p class="required-error">Please select your department before submitting.</p>',
            unsafe_allow_html=True
        )
    else:
        error_msg.empty()
        row = {
            "timestamp": datetime.now().isoformat(), "department": dept,
            "q1": q1, "q2": q2, "q3": q3, "q4": q4, "q5": q5,
            "q6": q6, "q7": q7, "q8": q8, "q9": q9, "q10": q10,
            "q11": q11, "q12": q12, "q13": q13, "q14": q14, "q15": q15
        }
        save_response(row)
        st.session_state["submitted"] = True
        st.rerun()
