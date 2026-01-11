import streamlit as st
from groq import Groq
from pptx import Presentation
from pptx.util import Pt
from pypdf import PdfReader
import tempfile
import os
import re

st.set_page_config(page_title="IA Médica PPTX", page_icon="🏥")
st.title("🏥 Generador Médico (Powered by Groq)")

# Verificar API KEY
if "GROQ_API_KEY" not in st.secrets:
    st.error("❌ Falta la GROQ_API_KEY en Secrets")
    st.stop()

client = Groq(api_key=st.secrets["GROQ_API_KEY"])

num_slides = st.slider("Número de diapositivas", 3, 15, 8)
uploaded_file = st.file_uploader("Sube tu PDF médico", type="pdf")

if uploaded_file:
    st.success(f"✅ Archivo cargado: {uploaded_file.name}")

    if st.button("🚀 Generar PowerPoint"):
        try:
            with st.spinner("📄 Leyendo PDF..."):
                reader = PdfReader(uploaded_file)
                text = ""
                for page in reader.pages:
                    text += page.extract_text() + "\n"

            with st.spinner("🤖 Analizando con IA..."):
                chat_completion = client.chat.completions.create(
                    messages=[
                        {
                            "role": "system",
                            "content": "Eres un especialista médico experto en crear presentaciones profesionales."
                        },
                        {
                            "role": "user",
                            "content": f"""A partir del siguiente texto médico, genera EXACTAMENTE {num_slides} diapositivas.

Para CADA diapositiva usa este formato:

---SLIDE---
TÍTULO: [título claro y conciso]
CONTENIDO:
• [punto clave 1]
• [punto clave 2]
• [punto clave 3]

Responde SOLO con las diapositivas en ese formato. No agregues introducciones ni conclusiones fuera del formato.

Texto:
{text[:18000]}"""
                        }
                    ],
                    model="llama-3.1-8b-instant",
                    temperature=0.7,
                    max_tokens=4000
                )
                response_text = chat_completion.choices[0].message.content

            # Procesar respuesta y crear diapositivas
            with st.spinner("📊 Creando presentación..."):
                prs = Presentation()
                prs.slide_width = Pt(1280)
                prs.slide_height = Pt(720)

                # Diapositiva de título
                slide_layout = prs.slide_layouts[0]
                slide = prs.slides.add_slide(slide_layout)
                title = slide.shapes.title
                subtitle = slide.placeholders[1]
                title.text = "Análisis Médico IA"
                subtitle.text = f"Basado en: {uploaded_file.name}"

                # Dividir por slides
                slides_data = response_text.split("---SLIDE---")
                
                for slide_text in slides_data:
                    if not slide_text.strip():
                        continue
                    
                    # Extraer título y contenido
                    lines = slide_text.strip().split("\n")
                    slide_title = ""
                    slide_content = []
                    
                    for line in lines:
                        if line.startswith("TÍTULO:"):
                            slide_title = line.replace("TÍTULO:", "").strip()
                        elif line.startswith("CONTENIDO:"):
                            continue
                        elif line.strip().startswith("•") or line.strip().startswith("-"):
                            slide_content.append(line.strip())
                    
                    # Crear slide
                    if slide_title:
                        slide_layout = prs.slide_layouts[1]
                        slide = prs.slides.add_slide(slide_layout)
                        
                        # Título
                        title_shape = slide.shapes.title
                        title_shape.text = slide_title
                        
                        # Contenido
                        body_shape = slide.placeholders[1]
                        tf = body_shape.text_frame
                        tf.clear()
                        
                        for point in slide_content[:6]:  # Máximo 6 puntos
                            p = tf.add_paragraph()
                            p.text = point
                            p.level = 0

                # Guardar
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
                    prs.save(tmp.name)
                    with open(tmp.name, "rb") as f:
                        st.success(f"✅ Presentación creada con {len(prs.slides)} diapositivas")
                        st.download_button(
                            "📥 Descargar PowerPoint",
                            f,
                            file_name="presentacion_medica.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                os.unlink(tmp.name)

        except Exception as e:
            st.error(f"Error: {e}")
else:
    st.info("⬆️ Sube un PDF para comenzar")
