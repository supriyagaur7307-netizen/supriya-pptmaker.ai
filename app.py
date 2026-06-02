import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import fitz
import io

st.set_page_config(page_title="Supriya ppt maker .ai", layout="wide")

# CSS for Stylish Header and Dark UI
st.markdown("""
    <style>
    .stApp { background-color: #050505; }
    .stylish-title { 
        font-family: 'Segoe UI', sans-serif; 
        font-size: 50px; 
        font-weight: 900; 
        background: linear-gradient(90deg, #38bdf8, #818cf8, #c084fc);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        margin-bottom: 30px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="stylish-title">🚀 Supriya ppt maker .ai</h1>', unsafe_allow_html=True)

# SIDEBAR DESIGN STUDIO
st.sidebar.header("🎨 Design Studio")
theme_bg = st.sidebar.color_picker("Pick Background Color", "#0f172a")
theme_accent = st.sidebar.color_picker("Pick Accent Color", "#38bdf8")
font_choice = st.sidebar.selectbox("Choose Font:", ["Arial", "Calibri", "Verdana", "Georgia", "Trebuchet MS"])

col1, col2 = st.columns([1, 2])

with col1:
    uploaded_file = st.file_uploader("Upload PDF:", type=["pdf"])

with col2:
    if uploaded_file:
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        prs = Presentation()
        
        for i, page in enumerate(doc):
            text = page.get_text()
            with st.expander(f"⚙️ Slide {i+1} Editor", expanded=(i==0)):
                title = st.text_input(f"Title {i+1}", value=f"Topic {i+1}", key=f"t{i}")
                content = st.text_area(f"Content {i+1}", value=text[:300], key=f"c{i}")
            
            # Slide Creation
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            fill = slide.background.fill
            fill.solid()
            # Convert hex to RGB for PPT
            r, g, b = tuple(int(theme_bg.lstrip('#')[j:j+2], 16) for j in (0, 2, 4))
            fill.fore_color.rgb = RGBColor(r, g, b)
            
            # Title
            t_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(1))
            p = t_box.text_frame.paragraphs[0]
            p.text = title
            p.font.name = font_choice
            p.font.size = Pt(40)
            
        if st.button("Generate Final PPT ⚡"):
            st.balloons()
            buf = io.BytesIO()
            prs.save(buf)
            buf.seek(0)
            st.download_button("📥 Download Premium PPT", buf, "Supriya_Premium.pptx")
