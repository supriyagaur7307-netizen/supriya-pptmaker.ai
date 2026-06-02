import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import fitz  # PyMuPDF
import io

# Page Config
st.set_page_config(page_title="Supriya ppt maker .ai", page_icon="🚀", layout="wide")

# CSS Styling (Gamma-like UI)
st.markdown("""
    <style>
    .main-header { font-size: 42px; font-weight: 800; text-align: center; color: #38bdf8; }
    .gamma-card { background: #111827; border-radius: 16px; padding: 24px; border: 1px solid #1f2937; }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<div class="main-header">🚀 Supriya ppt maker .ai</div>', unsafe_allow_html=True)

col1, col2 = st.columns([1, 2])

with col1:
    st.markdown('<div class="gamma-card">', unsafe_allow_html=True)
    st.subheader("📁 Upload Module")
    uploaded_file = st.file_uploader("PDF upload karein:", type=["pdf"])
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    if uploaded_file:
        pdf_bytes = uploaded_file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        
        # Slides loop with Text Editor
        for index, page in enumerate(doc):
            text_content = page.get_text()
            
            # Text Editor Expander
            with st.expander(f"⚙️ Slide Layout {index + 1} Editor", expanded=(index==0)):
                edited_title = st.text_input(f"Slide {index+1} Title:", value=f"Topic {index+1}", key=f"t_{index}")
                edited_body = st.text_area(f"Slide {index+1} Content:", value=text_content[:200], key=f"b_{index}")
            
            # Slide Creation
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(15, 23, 42) # Midnight Navy
            
            # Adding Text to Slide
            title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11), Inches(1))
            title_box.text_frame.text = edited_title
            
        ppt_buffer = io.BytesIO()
        prs.save(ppt_buffer)
        ppt_buffer.seek(0)
        
        if st.button("Generate & Download"):
            st.balloons() # Confetti effect
            st.success("🎉 Gamma-Style Layout Compiled Perfectly!")
            st.download_button("📥 Download Presentation", ppt_buffer, "Supriya_Presentation.pptx")
