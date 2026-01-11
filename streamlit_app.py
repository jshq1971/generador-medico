import streamlit as st
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches
from pypdf import PdfReader
import tempfile
import os

st.set_page_config(page_title="Generador Médico IA", page_icon="🏥")
st.title("🏥 Generador de Presentaciones Médicas")

# API KEY
if "GEMINI_API_KEY" not in st.secrets:
    st.error("❌ Falta la GEMINI_API_KEY en Secrets")
    st.stop()

genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

num_slides = st.slider("Número de diapositivas", 1, 20, 10)
uploaded_file = st.file_uploader("Sube tu PDF médico", type="pdf")

if uploaded_file:
    st.success(f"✅ Archivo cargado: {uploaded_file.name}")

    if st.button("🚀 Generar PowerPoint"):
        try:
            with st.spinner("Leyendo PDF..."):
                reader = PdfReader(uploaded_file)
                text = ""
                for page in reader.pages:
                    text += page.extract_text() + "\n"

            with st.spinner("Analizando con IA..."):
                model = genai.GenerativeModel("gemini-pro")

                prompt = f"""
Eres un especialista médico.
A partir del siguiente texto, crea una presentación de {num_slides} diapositivas.

Para cada diapositiva incluye:
- Título
- 3–4 puntos clave clínicos

Texto:
{text[:12000]}
"""
                response = model.generate_content(prompt)

            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = "Resumen Médico IA"
            slide.placeholders[1].text = response.text[:1500]

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
                prs.save(tmp.name)
                with open(tmp.name, "rb") as f:
                    st.download_button(
                        "📥 Descargar PowerPoint",
                        f,
                        file_name="presentacion_medica.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )

            os.unlink(tmp.name)

        except Exception as e:
            st.error(f"Error: {e}")
else:
    st.info("⬆️ Sube un PDF para comenzar")
