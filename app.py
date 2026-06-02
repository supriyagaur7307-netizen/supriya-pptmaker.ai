import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import fitz  # PyMuPDF
import io
import requests

# Wide Page Configuration
st.set_page_config(page_title="Supriya ppt maker .ai", page_icon="🚀", layout="wide")

st.markdown("""
    <style>
    body { background-color: #0b0f19; color: #f1f5f9; }
    .main-header { font-size: 42px; font-weight: 800; background: linear-gradient(135deg, #6366f1, #a855f7, #ec4899); -webkit-background-clip: text; -webkit-text-fill-color: transparent; text-align: center; padding: 10px; font-family: 'Helvetica Neue', Arial, sans-serif; }
    .sub-header { font-size: 15px; color: #94a3b8; text-align: center; margin-bottom: 30px; }
    .gamma-card { background: #111827; border-radius: 16px; padding: 24px; border: 1px solid #1f2937; margin-bottom: 20px; box-shadow: 0 20px 25px -5px rgba(0,0,0,0.5); }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<div class="main-header">🚀 Supriya ppt maker .ai</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Premium Presentation Engine | Clean Layouts, Proper Fonts & High-Res Images</div>', unsafe_allow_html=True)

# SIDEBAR DESIGN STUDIO
st.sidebar.markdown("### 🎨 Executive Styles")
template_choice = st.sidebar.selectbox(
    "Choose Theme Matrix:",
    ["Premium Midnight Navy", "Physics Wallah Elite", "Minimalist Clean Light"]
)

# Color Scheme Mapping
if template_choice == "Premium Midnight Navy":
    bg_hex, title_hex, body_hex = "#0f172a", "#38bdf8", "#cbd5e1"
elif template_choice == "Physics Wallah Elite":
    bg_hex, title_hex, body_hex = "#000000", "#facc15", "#f8fafc"
else:
    bg_hex, title_hex, body_hex = "#ffffff", "#1e293b", "#475569"

def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return tuple(int(hex_str[i:i+2], 16) for i in (0, 2, 4))

col1, col2 = st.columns([1, 2])

with col1:
    st.markdown('<div class="gamma-card">', unsafe_allow_html=True)
    st.subheader("📁 Upload Module")
    uploaded_file = st.file_uploader("Drop your PDF notes here:", type=["pdf"])
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    if uploaded_file:
        pdf_bytes = uploaded_file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        
        r_bg, g_bg, b_bg = hex_to_rgb(bg_hex)
        
        for index, page in enumerate(doc):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(r_bg, g_bg, b_bg)
            
            # Simple Text Insertion
            title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11), Inches(1))
            title_box.text_frame.text = "Presentation Slide"
            
        ppt_buffer = io.BytesIO()
        prs.save(ppt_buffer)
        ppt_buffer.seek(0)
        
        st.success("File Processed!")
        st.download_button("📥 Download Presentation", ppt_buffer, "Supriya_Presentation.pptx")
