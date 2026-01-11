import streamlit as st
from groq import Groq
from pypdf import PdfReader
import tempfile
import os
import re
from datetime import date
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# Configuración de página
st.set_page_config(page_title="PPT Médico 60+", page_icon="🏥", layout="wide")

# Estilos Médicos (Oxford Blue & Clinical White)
PRIMARY = RGBColor(0, 33, 71)
ACCENT = RGBColor(0, 122, 255)
TEXT_DARK = RGBColor(30, 30, 30)
WHITE = RGBColor(255, 255, 255)
SLIDE_W, SLIDE_H = Inches(13.33), Inches(7.5)

if "GROQ_API_KEY" not in st.secrets:
    st.error("❌ Configura GROQ_API_KEY en Secrets")
    st.stop()

client = Groq(api_key=st.secrets["GROQ_API_KEY"])

# --- FUNCIONES DE DISEÑO ---
def apply_medical_style(slide, title_text, slide_num, total):
    # Barra superior fina
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.15))
    bar.fill.solid()
    bar.fill.fore_color.rgb = PRIMARY
    bar.line.fill.background()
    
    # Título
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.5), Inches(1))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title_text.upper()
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    
    # Numeración
    num_box = slide.shapes.add_textbox(Inches(12.3), Inches(6.8), Inches(1), Inches(0.5))
    p_num = num_box.text_frame.paragraphs[0]
    p_num.text = f"{slide_num}/{total}"
    p_num.font.size = Pt(12)
    p_num.font.color.rgb = RGBColor(150, 150, 150)
    p_num.alignment = PP_ALIGN.RIGHT

def add_bullets(slide, bullets):
    body = slide.shapes.add_textbox(Inches(1), Inches(1.6), Inches(11.3), Inches(5))
    tf = body.text_frame
    tf.word_wrap = True
    for b in bullets[:7]: # Máximo 7 puntos por slide para legibilidad
        p = tf.add_paragraph()
        p.text = f"• {b}"
        p.font.size = Pt(18)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(10)

# --- MOTOR DE GENERACIÓN ---
def get_slides_batch(text, num_slides, batch_idx):
    # Dividimos el texto para que la IA se enfoque en partes diferentes del PDF
    start = batch_idx * 5000
    segment = text[start : start + 15000]
    
    prompt = f"""
    Eres un Especialista en Oncología y Educación Médica. 
    Genera EXACTAMENTE {num_slides} diapositivas de contenido técnico.
    
    FORMATO:
    ---SLIDE---
    TÍTULO: [Título clínico]
    PUNTOS:
    - [Dato relevante 1]
    - [Dato relevante 2]
    - [Dato relevante 3]
    
    Usa terminología médica avanzada. No resumas demasiado, mantén el rigor.
    TEXTO DE REFERENCIA:
    {segment}
    """
    
    chat = client.chat.completions.create(
        messages=[{"role": "user", "content": prompt}],
        model="llama-3.1-8b-instant",
        temperature=0.3
    )
    return chat.choices[0].message.content

# --- INTERFAZ ---
st.title("🏥 Generador de Clases Médicas (60 Diapositivas)")
st.markdown("Este sistema genera presentaciones extensas dividiendo el trabajo en bloques para evitar errores de conexión.")

num_total = st.slider("Número total de diapositivas deseadas", 20, 80, 60)
uploaded_file = st.file_uploader("Sube el PDF (Guía, Paper o Texto)", type="pdf")

if uploaded_file:
    if st.button(f"🚀 Iniciar Generación de {num_total} Slides"):
        try:
            # 1. Leer PDF
            reader = PdfReader(uploaded_file)
            full_text = ""
            for page in reader.pages:
                full_text += (page.extract_text() or "") + "\n"
            
            # 2. Calcular Bloques (Ej: 60 slides = 4 bloques de 15)
            batch_size = 15
            num_batches = (num_total // batch_size) + (1 if num_total % batch_size != 0 else 0)
            
            all_raw_content = ""
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # 3. Generación por bloques (Evita el 404)
            for i in range(num_batches):
                status_text.text(f"⏳ Generando bloque {i+1} de {num_batches}...")
                slides_to_gen = batch_size if (i < num_batches - 1) else (num_total - (i * batch_size))
                all_raw_content += get_slides_batch(full_text, slides_to_gen, i)
                progress_bar.progress((i + 1) / num_batches)
            
            # 4. Crear PPTX
            status_text.text("🎨 Diseñando presentación final...")
            prs = Presentation()
            prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
            
            # Portada
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
            rect.fill.solid()
            rect.fill.fore_color.rgb = PRIMARY
            title = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(11.3), Inches(2))
            p = title.text_frame.paragraphs[0]
            p.text = uploaded_file.name.replace(".pdf", "").upper()
            p.font.size = Pt(40)
            p.font.bold = True
            p.font.color.rgb = WHITE
            p.alignment = PP_ALIGN.CENTER

            # Procesar Slides
            raw_slides = all_raw_content.split("---SLIDE---")
            valid_slides = 0
            
            for s_data in raw_slides:
                if "TÍTULO:" in s_data and "PUNTOS:" in s_data:
                    valid_slides += 1
                    lines = s_data.strip().split("\n")
                    title_text = lines[0].replace("TÍTULO:", "").strip()
                    
                    bullets = []
                    for l in lines:
                        if l.strip().startswith("-"):
                            bullets.append(l.strip().lstrip("-").strip())
                    
                    slide = prs.slides.add_slide(prs.slide_layouts[6])
                    apply_medical_style(slide, title_text, valid_slides, num_total)
                    add_bullets(slide, bullets)
                    
                    if valid_slides >= num_total: break

            # 5. Descarga
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
                prs.save(tmp.name)
                with open(tmp.name, "rb") as f:
                    st.success(f"✅ ¡Éxito! Se generaron {valid_slides} diapositivas médicas.")
                    st.download_button(
                        label="📥 Descargar Presentación Completa",
                        data=f,
                        file_name=f"Clase_Medica_{date.today()}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
            os.unlink(tmp.name)
            status_text.empty()

        except Exception as e:
            st.error(f"Ocurrió un error: {e}")
