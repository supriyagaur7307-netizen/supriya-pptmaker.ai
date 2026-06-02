import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import fitz
import io
import requests

st.set_page_config(page_title="Supriya Presentation Studio AI", layout="wide")

# CSS for Studio Feel
st.markdown("""<style>.stApp {background: #050505; color: #e2e8f0;}</style>""", unsafe_allow_html=True)

# Customization Sidebar
st.sidebar.header("🎨 Studio Design Panel")
theme_color = st.sidebar.color_picker("Pick Brand Color", "#38bdf8")
font_choice = st.sidebar.selectbox("Font Style", ["Arial", "Calibri", "Helvetica", "Verdana"])

def get_dynamic_image(title):
    # Context-aware image search
    url = f"https://source.unsplash.com/800x600/?{title.replace(' ', '+')}"
    try:
        res = requests.get(url, timeout=5)
        return io.BytesIO(res.content) if res.status_code == 200 else None
    except: return None

def create_slide(prs, title, points, color_hex, font):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Background - Premium Dark
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(10, 10, 10)
    
    # Title - Dynamic
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12), Inches(1))
    p = title_box.text_frame.paragraphs[0]
    p.text, p.font.size, p.font.bold = title, Pt(44), True
    p.font.color.rgb = RGBColor.from_string(color_hex.lstrip('#'))
    
    # Body - Professional List
    body_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(6), Inches(4))
    for pt in points:
        p = body_box.text_frame.add_paragraph()
        p.text, p.font.size, p.font.name = f"• {pt}", Pt(24), font
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.space_after = Pt(20)
        
    # Image - Unique for each slide
    img_data = get_dynamic_image(title)
    if img_data:
        slide.shapes.add_picture(img_data, Inches(7.2), Inches(1.8), Inches(5.5), Inches(4))

# Execution
uploaded_file = st.file_uploader("Upload PDF Notes:", type=["pdf"])
if uploaded_file:
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    
    for page in doc:
        txt = page.get_text().split('\n')
        title, pts = (txt[0], txt[1:7]) if txt else ("Concept", [])
        create_slide(prs, title, pts, theme_color, font_choice)
        
    buf = io.BytesIO()
    prs.save(buf)
    st.download_button("⚡ Download Final Premium PPT", buf.getvalue(), "Supriya_Pro_Presentation.pptx")
