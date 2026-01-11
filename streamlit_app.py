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

# Configuración
st.set_page_config(page_title="PPT Médico Pro", page_icon="🏥")
st.title("🏥 Generador Médico Profesional")

if "GROQ_API_KEY" not in st.secrets:
    st.error("❌ Falta la GROQ_API_KEY en Secrets")
    st.stop()

client = Groq(api_key=st.secrets["GROQ_API_KEY"])

# Colores y Medidas
PRIMARY = RGBColor(0, 51, 102)
ACCENT = RGBColor(0, 102, 204)
WHITE = RGBColor(255, 255, 255)
SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

def add_pro_style(slide, title_text):
    # Barra superior
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.2))
    bar.fill.solid()
    bar.fill.fore_color.rgb = PRIMARY
    bar.line.fill.background()
    # Título
    txt = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.5), Inches(1))
    p = txt.text_frame.paragraphs[0]
    p.text = title_text
    p.font.size = Pt(34)
    p.font.bold = True
    p.font.color.rgb = PRIMARY

num_slides = st.slider("Número de diapositivas", 5, 15, 10)
uploaded_file = st.file_uploader("Sube tu PDF médico", type="pdf")

if uploaded_file:
    st.success(f"✅ PDF cargado: {uploaded_file.name}")
    if st.button("🚀 Generar Presentación Enriquecida"):
        try:
            with st.spinner("Analizando contenido..."):
                reader = PdfReader(uploaded_file)
                pdf_text = ""
                for page in reader.pages:
                    pdf_text += (page.extract_text() or "") + "\n"

                # Prompt para la IA
                prompt = f"Genera {num_slides} diapositivas médicas profesionales. Para cada una usa: ---SLIDE--- TÍTULO: [título] CONTENIDO: • [punto 1] • [punto 2] • [punto 3]. Texto: {pdf_text[:15000]}"
                
                chat = client.chat.completions.create(
                    messages=[{"role": "user", "content": prompt}],
                    model="llama-3.1-8b-instant",
                )
                response = chat.choices[0].message.content

            with st.spinner("Creando diseño visual..."):
                prs = Presentation()
                prs.slide_width = SLIDE_W
                prs.slide_height = SLIDE_H

                # Portada
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
                rect.fill.solid()
                rect.fill.fore_color.rgb = PRIMARY
                title = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(11.3), Inches(2))
                p = title.text_frame.paragraphs[0]
                p.text = uploaded_file.name.replace(".pdf", "").upper()
                p.font.size = Pt(44)
                p.font.bold = True
                p.font.color.rgb = WHITE
                p.alignment = PP_ALIGN.CENTER

                # Contenido
                slides_raw = response.split("---SLIDE---")
                for s_raw in slides_raw:
                    if "TÍTULO:" in s_raw:
                        slide = prs.slides.add_slide(prs.slide_layouts[6])
                        lines = s_raw.strip().split("\n")
                        t_text = lines[0].replace("TÍTULO:", "").strip()
                        add_pro_style(slide, t_text)
                        
                        body = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(11), Inches(4.5))
                        tf = body.text_frame
                        for line in lines:
                            if line.strip().startswith("•"):
                                p = tf.add_paragraph()
                                p.text = line.strip()
                                p.font.size = Pt(22)
                                p.space_after = Pt(12)

                # Algoritmo Visual (Flujo)
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                add_pro_style(slide, "Algoritmo de Manejo Clínico")
                for i in range(4):
                    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1 + i*3), Inches(3), Inches(2.5), Inches(1.5))
                    box.fill.solid()
                    box.fill.fore_color.rgb = ACCENT
                    box.text = f"Paso {i+1}"

                # Guardar
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
                    prs.save(tmp.name)
                    with open(tmp.name, "rb") as f:
                        st.download_button("📥 Descargar PowerPoint Pro", f, file_name="presentacion_pro.pptx")
                os.unlink(tmp.name)

        except Exception as e:
            st.error(f"Error: {e}")
