import streamlit as st
from groq import Groq
from pypdf import PdfReader
import fitz  # PyMuPDF
import tempfile
import os
import re
from datetime import date
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="PPT Médico Pro 60+", page_icon="🏥", layout="wide")

if "GROQ_API_KEY" not in st.secrets:
    st.error("Falta GROQ_API_KEY en Secrets")
    st.stop()

client = Groq(api_key=st.secrets["GROQ_API_KEY"])

# Estilos
PRIMARY = RGBColor(0, 33, 71)
ACCENT = RGBColor(0, 122, 255)
WHITE = RGBColor(255, 255, 255)
SLIDE_W, SLIDE_H = Inches(13.33), Inches(7.5)

# --- ESTADO DE SESIÓN ---
if "step" not in st.session_state: st.session_state.step = "config"
if "slides_data" not in st.session_state: st.session_state.slides_data = []
if "pdf_text" not in st.session_state: st.session_state.pdf_text = ""
if "flow_steps" not in st.session_state: st.session_state.flow_steps = []
if "fig_paths" not in st.session_state: st.session_state.fig_paths = []

# --- FUNCIONES ---
def get_batch(text, num, block_idx):
    start = block_idx * 8000
    segment = text[start : start + 18000]
    prompt = f"Actúa como oncólogo experto. Genera {num} diapositivas técnicas. FORMATO: ---SLIDE--- TÍTULO: [título] PUNTOS: - [punto 1] - [punto 2]. Texto: {segment}"
    chat = client.chat.completions.create(
        messages=[{"role": "user", "content": prompt}],
        model="llama-3.1-8b-instant",
        temperature=0.3
    )
    return chat.choices[0].message.content

def get_flow(text):
    prompt = f"Crea un algoritmo clínico de 5 pasos basado en este texto médico. Solo lista numerada: 1. ... 2. ... Texto: {text[:15000]}"
    chat = client.chat.completions.create(
        messages=[{"role": "user", "content": prompt}],
        model="llama-3.1-8b-instant",
        temperature=0.2
    )
    return [l.strip() for l in chat.choices[0].message.content.split("\n") if l.strip()[:1].isdigit()]

def render_fig(pdf_bytes, page_idx):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[page_idx]
    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    pix.save(tmp.name)
    return tmp.name

# --- INTERFAZ ---
st.title("🏥 Generador Médico Profesional (Modo Extenso)")

if st.session_state.step == "config":
    with st.sidebar:
        num_total = st.slider("Número de diapositivas", 10, 80, 60)
        num_figs = st.slider("Número de figuras/imágenes", 0, 10, 3)
        include_flow = st.checkbox("Incluir Algoritmo Clínico", True)
    
    uploaded_file = st.file_uploader("Sube tu PDF médico", type="pdf")
    
    if uploaded_file and st.button("🚀 Iniciar Procesamiento"):
        reader = PdfReader(uploaded_file)
        text = ""
        for page in reader.pages: text += (page.extract_text() or "") + "\n"
        st.session_state.pdf_text = text
        st.session_state.num_total = num_total
        st.session_state.num_figs = num_figs
        st.session_state.include_flow = include_flow
        st.session_state.pdf_bytes = uploaded_file.getvalue()
        st.session_state.doc_name = uploaded_file.name
        st.session_state.step = "generating"
        st.rerun()

elif st.session_state.step == "generating":
    total = st.session_state.num_total
    current_count = len(st.session_state.slides_data)
    
    if current_count < total:
        st.info(f"Progreso: {current_count} / {total} diapositivas")
        batch_size = 15
        if st.button(f"▶️ Generar Siguiente Bloque (15 slides)"):
            with st.spinner("Generando contenido técnico..."):
                raw = get_batch(st.session_state.pdf_text, batch_size, current_count // batch_size)
                new = [s for s in raw.split("---SLIDE---") if "TÍTULO:" in s]
                st.session_state.slides_data.extend(new)
                st.rerun()
    else:
        st.success("✅ Contenido generado. Procesando elementos visuales...")
        if st.session_state.include_flow and not st.session_state.flow_steps:
            st.session_state.flow_steps = get_flow(st.session_state.pdf_text)
        
        if st.session_state.num_figs > 0 and not st.session_state.fig_paths:
            # Renderizamos las últimas páginas (donde suelen estar las figuras)
            doc = fitz.open(stream=st.session_state.pdf_bytes, filetype="pdf")
            for i in range(max(0, len(doc)-st.session_state.num_figs), len(doc)):
                st.session_state.fig_paths.append(render_fig(st.session_state.pdf_bytes, i))
        
        st.session_state.step = "finalize"
        st.rerun()

elif st.session_state.step == "finalize":
    if st.button("🎨 Crear y Descargar PowerPoint"):
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        
        # Portada
        s = prs.slides.add_slide(prs.slide_layouts[6])
        s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H).fill.solid().fore_color.rgb = PRIMARY
        
        # Slides de contenido
        for i, s_data in enumerate(st.session_state.slides_data[:st.session_state.num_total]):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            # Barra y Título
            slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.2)).fill.solid().fore_color.rgb = PRIMARY
            lines = s_data.strip().split("\n")
            title = lines[0].replace("TÍTULO:", "").strip()
            txt = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.5), Inches(1))
            txt.text_frame.paragraphs[0].text = title.upper()
            txt.text_frame.paragraphs[0].font.size = Pt(26)
            txt.text_frame.paragraphs[0].font.color.rgb = PRIMARY
            
            # Puntos
            body = slide.shapes.add_textbox(Inches(1), Inches(1.8), Inches(11), Inches(5))
            for l in lines:
                if l.strip().startswith("-"):
                    p = body.text_frame.add_paragraph()
                    p.text = "• " + l.strip().lstrip("-").strip()
                    p.font.size = Pt(18)
            
        # Algoritmo
        if st.session_state.flow_steps:
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            for i, step in enumerate(st.session_state.flow_steps):
                box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(1.5 + i*1.1), Inches(10), Inches(0.8))
                box.fill.solid().fore_color.rgb = ACCENT
                box.text_frame.text = step
        
        # Figuras
        for path in st.session_state.fig_paths:
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.shapes.add_picture(path, Inches(1), Inches(1), width=Inches(11))

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
            prs.save(tmp.name)
            with open(tmp.name, "rb") as f:
                st.download_button("📥 DESCARGAR PPT COMPLETO", f, file_name="Presentacion_Medica_Pro.pptx")
    
    if st.button("🔄 Reiniciar"):
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.rerun()
