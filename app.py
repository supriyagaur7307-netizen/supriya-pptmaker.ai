import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import fitz  # PyMuPDF
import io
import requests

# Page Config
st.set_page_config(page_title="Supriya PPT Maker Pro", layout="wide")

# Theme Profiles
THEMES = {
    "Midnight Navy": {"bg": "#0f172a", "title": "#38bdf8", "body": "#cbd5e1"},
    "Corporate Elite": {"bg": "#ffffff", "title": "#1e293b", "body": "#475569"},
    "Modern Dark": {"bg": "#121212", "title": "#facc15", "body": "#e2e8f0"}
}

def hex_to_rgb(hex_str):
    h = hex_str.lstrip('#')
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def get_smart_image(keyword):
    # Dynamic Search Query
    query = keyword.replace(" ", "+")
    url = f"https://source.unsplash.com/800x600/?{query}"
    try:
        res = requests.get(url, timeout=5)
        if res.status_code == 200:
            return io.BytesIO(res.content)
    except:
        return None
    return None

st.title("🚀 Supriya PPT Maker Pro")
theme_name = st.sidebar.selectbox("Select Theme:", list(THEMES.keys()))
theme = THEMES[theme_name]

uploaded_file = st.file_uploader("Upload PDF:", type=["pdf"])

if uploaded_file:
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

    for i, page in enumerate(doc):
        text = page.get_text().split('\n')
        title = text[0] if text else "Slide " + str(i+1)
        points = [l for l in text[1:6] if len(l) > 10]
        
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        # Background
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*hex_to_rgb(theme["bg"]))
        
        # Title
        t_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(1))
        p = t_box.text_frame.paragraphs[0]
        p.text, p.font.size, p.font.color.rgb = title, Pt(40), RGBColor(*hex_to_rgb(theme["title"]))
        
        # Body & Image Grid
        body = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(6), Inches(4))
        for pt in points:
            p = body.text_frame.add_paragraph()
            p.text, p.font.size, p.font.color.rgb = f"• {pt}", Pt(20), RGBColor(*hex_to_rgb(theme["body"]))
            
        img = get_smart_image(title)
        if img:
            slide.shapes.add_picture(img, Inches(7), Inches(2), Inches(5.5), Inches(3.5))

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    st.download_button("Download Premium PPT", buf, "Presentation.pptx")
