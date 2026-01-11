import streamlit as st
from groq import Groq
from pypdf import PdfReader
import tempfile
import os
from datetime import date
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="PPT Médico 60+ Estable", page_icon="🏥")

if "GROQ_API_KEY" not in st.secrets:
    st.error("Falta GROQ_API_KEY en Secrets")
    st.stop()

client = Groq(api_key=st.secrets["GROQ_API_KEY"])

# Colores Médicos
PRIMARY = RGBColor(0, 33, 71)
WHITE = RGBColor(255, 255, 255)
SLIDE_W, SLIDE_H = Inches(13.33), Inches(7.5)

# --- ESTADO DE SESIÓN (Para evitar el 404) ---
if "current_step" not in st.session_state:
    st.session_state.current_step = 0
if "all_slides_data" not in st.session_state:
    st.session_state.all_slides_data = []
if "pdf_text" not in st.session_state:
    st.session_state.pdf_text = ""

# --- FUNCIONES ---
def get_batch(text, num, block_num):
    start_char = block_num * 6000
    segment = text[start_char : start_char + 15000]
    prompt = f"Actúa como oncólogo. Genera {num} diapositivas técnicas. Formato: ---SLIDE--- TÍTULO: [título] PUNTOS: - [punto 1] - [punto 2]. Texto: {segment}"
    chat = client.chat.completions.create(
        messages=[{"role": "user", "content": prompt}],
        model="llama-3.1-8b-instant",
        temperature=0.3
    )
    return chat.choices[0].message.content

def create_pptx(data_list, filename):
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    
    # Portada
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    rect.fill.solid()
    rect.fill.fore_color.rgb = PRIMARY
    
    # Contenido
    for i, s_raw in enumerate(data_list):
        if "TÍTULO:" in s_raw:
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            # Barra azul
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.2))
            bar.fill.solid()
            bar.fill.fore_color.rgb = PRIMARY
            
            lines = s_raw.strip().split("\n")
            title_text = lines[0].replace("TÍTULO:", "").strip()
            
            # Título
            txt = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.5), Inches(1))
            p = txt.text_frame.paragraphs[0]
            p.text = title_text.upper()
            p.font.size = Pt(26)
            p.font.bold = True
            p.font.color.rgb = PRIMARY
            
            # Puntos
            body = slide.shapes.add_textbox(Inches(1), Inches(1.8), Inches(11), Inches(5))
            tf = body.text_frame
            for l in lines:
                if l.strip().startswith("-"):
                    p = tf.add_paragraph()
                    p.text = "• " + l.strip().lstrip("-").strip()
                    p.font.size = Pt(18)
                    p.space_after = Pt(10)
    
    path = f"{filename}.pptx"
    prs.save(path)
    return path

# --- INTERFAZ ---
st.title("🏥 Generador Médico de 60 Slides")
st.info("Para evitar desconexiones, generaremos la presentación en 4 bloques de 15 diapositivas.")

uploaded_file = st.file_uploader("Sube tu PDF", type="pdf")

if uploaded_file:
    if not st.session_state.pdf_text:
        reader = PdfReader(uploaded_file)
        text = ""
        for page in reader.pages:
            text += (page.extract_text() or "") + "\n"
        st.session_state.pdf_text = text

    # Lógica de bloques
    total_blocks = 4
    current = st.session_state.current_step

    if current < total_blocks:
        if st.button(f"🚀 Generar Bloque {current + 1} de {total_blocks} (15 slides)"):
            with st.spinner(f"Generando bloque {current + 1}..."):
                raw = get_batch(st.session_state.pdf_text, 15, current)
                new_slides = [s for s in raw.split("---SLIDE---") if "TÍTULO:" in s]
                st.session_state.all_slides_data.extend(new_slides)
                st.session_state.current_step += 1
                st.rerun()
    else:
        st.success(f"✅ ¡Listo! {len(st.session_state.all_slides_data)} diapositivas generadas.")
        if st.button("🎨 Finalizar y Preparar Descarga"):
            pptx_path = create_pptx(st.session_state.all_slides_data, "Presentacion_Medica_60")
            with open(pptx_path, "rb") as f:
                st.download_button("📥 DESCARGAR POWERPOINT COMPLETO", f, file_name="Clase_Medica_60.pptx")
        
        if st.button("🔄 Empezar de nuevo"):
            st.session_state.current_step = 0
            st.session_state.all_slides_data = []
            st.rerun()

st.write(f"Progreso: {len(st.session_state.all_slides_data)} / 60 diapositivas")
