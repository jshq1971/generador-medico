import streamlit as st
from groq import Groq
from pptx import Presentation
from pypdf import PdfReader
import tempfile
import os

st.set_page_config(page_title="IA Médica PPTX", page_icon="🏥")
st.title("🏥 Generador Médico (Powered by Groq)")

# Verificar API KEY de Groq
if "GROQ_API_KEY" not in st.secrets:
    st.error("❌ Falta la GROQ_API_KEY en Secrets")
    st.stop()

client = Groq(api_key=st.secrets["GROQ_API_KEY"])

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

            with st.spinner("Analizando con IA (Groq)..."):
                # Usamos Llama 3 70b, que es excelente para medicina
                chat_completion = client.chat.completions.create(
                    messages=[
                        {
                            "role": "system",
                            "content": "Eres un experto médico. Crea presentaciones profesionales."
                        },
                        {
                            "role": "user",
                            "content": f"A partir de este texto, genera {num_slides} diapositivas con Título y 3 puntos clave cada una. Responde en ESPAÑOL.\n\nTexto: {text[:15000]}"
                        }
                    ],
                    model="llama3-70b-8192",
                )
                response_text = chat_completion.choices[0].message.content

            # Crear PPTX
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = "Análisis Médico IA"
            slide.placeholders[1].text = response_text[:1500]

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
                prs.save(tmp.name)
                with open(tmp.name, "rb") as f:
                    st.download_button("📥 Descargar PowerPoint", f, file_name="presentacion.pptx")
            os.unlink(tmp.name)

        except Exception as e:
            st.error(f"Error: {e}")
else:
    st.info("⬆️ Sube un PDF para comenzar")
