import streamlit as st
from groq import Groq
from pypdf import PdfReader
import fitz  # PyMuPDF
import tempfile
import os
from datetime import date
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="PPT Médico Ultra", page_icon="🏥")

if "GROQ_API_KEY" not in st.secrets:
    st.error("Falta GROQ_API_KEY")
    st.stop()

client = Groq(api_key=st.secrets["GROQ_API_KEY"])

# Estilos
PRIMARY = RGBColor(0, 33, 71)
ACCENT = RGBColor(0, 122, 255)
SLIDE_W, SLIDE_H = Inches(13.33), Inches(7.5)

# --- ESTADO DE SESIÓN ---
if "step" not in st.session_state: st.session_state.step = "config"
if "slides_data" not in st.session_state: st.session_state.slides_data = []
if "pdf_text" not in st.session_state: st.session_state.pdf_text = ""
if "fig_paths" not in st.session_state: st.session_state.fig_paths = []

# --- FUNCIÓN ULTRA RÁPIDA ---
def get_fast_batch(text, num, block_idx):
    # Usamos fragmentos más pequeños para velocidad máxima
    start = block_idx * 4000
    segment = text[start : start + 8000]
    prompt = f"Eres oncólogo. Genera {num} diapositivas. Formato: ---SLIDE--- TÍTULO: [título] PUNTOS: - [punto 1] - [punto 2]. Texto: {segment}"
    
    chat = client.chat.completions.create(
        messages=[{"role": "user", "content": prompt}],
        model="llama-3.1-8b-instant", # El modelo más rápido
        temperature=0.2,
        max_tokens=1500 # Respuesta corta = respuesta rápida
    )
    return chat.choices[0].message.content

def render_fig(pdf_bytes, page_idx):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pix = doc[page_idx].get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    pix.save(tmp.name)
    return tmp.name

# --- INTERFAZ ---
st.title("🏥 Generador Médico Ultra-Rápido")

if st.session_state.step == "config":
    num_total = st.number_input("¿Cuántas diapositivas quieres? (Ej: 60)", 10, 100, 60)
    num_figs = st.slider("Imágenes del PDF", 0, 10, 3)
    uploaded_file = st.file_uploader("Sube tu PDF", type="pdf")
    
    if uploaded_file and st.button("🚀 Iniciar"):
        reader = PdfReader(uploaded_file)
        text = ""
        for page in reader.pages: text += (page.extract_text() or "") + "\n"
        st.session_state.pdf_text = text
        st.session_state.num_total = num_total
        st.session_state.num_figs = num_figs
        st.session_state.pdf_bytes = uploaded_file.getvalue()
        st.session_state.step = "generating"
        st.rerun()

elif st.session_state.step == "generating":
    current = len(st.session_state.slides_data)
    total = st.session_state.num_total
    
    st.progress(current / total)
    st.write(f"Progreso: {current} / {total} diapositivas")

    if current < total:
        # Generamos de 5 en 5 para que sea instantáneo
        if st.button(f"▶️ Generar Siguiente Bloque (+5 slides)"):
            with st.spinner("Generando..."):
                raw = get_fast_batch(st.session_state.pdf_text, 5, current // 5)
                new = [s for s in raw.split("---SLIDE---") if "TÍTULO:" in s]
                st.session_state.slides_data.extend(new)
                st.rerun()
    else:
        st.success("✅ Contenido listo.")
        if st.button("🎨 Finalizar y Descargar"):
            # Renderizar figuras rápido
            if st.session_state.num_figs > 0:
                doc = fitz.open(stream=st.session_state.pdf_bytes, filetype="pdf")
                for i in range(max(0, len(doc)-st.session_state.num_figs), len(doc)):
                    st.session_state.fig_paths.append(render_fig(st.session_state.pdf_bytes, i))
            
            # Crear PPT
            prs = Presentation()
            prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
            for s_data in st.session_state.slides_data[:total]:
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                lines = s_data.strip().split("\n")
                title_text = lines[0].replace("TÍTULO:", "").strip()
                
                # Título
                txt = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.5), Inches(1))
                txt.text_frame.paragraphs[0].text = title_text.upper()
                txt.text_frame.paragraphs[0].font.size = Pt(26)
                txt.text_frame.paragraphs[0].font.color.rgb = PRIMARY
                
                # Puntos
                body = slide.shapes.add_textbox(Inches(1), Inches(1.8), Inches(11), Inches(5))
                for l in lines:
                    if l.strip().startswith("-"):
                        p = body.text_frame.add_paragraph()
                        p.text = "• " + l.strip().lstrip("-").strip()
                        p.font.size = Pt(18)
            
            # Añadir Figuras
            for path in st.session_state.fig_paths:
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                slide.shapes.add_picture(path, Inches(1), Inches(1), width=Inches(11))

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
                prs.save(tmp.name)
                with open(tmp.name, "rb") as f:
                    st.download_button("📥 DESCARGAR PPT", f, file_name="Presentacion_Medica.pptx")

    if st.button("🔄 Reiniciar"):
        st.session_state.step = "config"
        st.session_state.slides_data = []
        st.rerun()
