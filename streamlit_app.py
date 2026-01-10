import streamlit as st
import google.generativeai as genai
from pptx import Presentation
import tempfile
import os

# Configuración
st.set_page_config(page_title="IA Médica PPTX", page_icon="🏥")
st.title("🏥 Generador de Presentaciones Médicas")

# 1. Verificar API KEY
if "GEMINI_API_KEY" not in st.secrets:
    st.error("⚠️ Falta la GEMINI_API_KEY en los Secrets de Streamlit.")
    st.stop()

genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

# 2. Interfaz
num_slides = st.slider("Número de diapositivas", 1, 20, 10)
uploaded_file = st.file_uploader("Sube tu PDF médico", type="pdf")

if uploaded_file is not None:
    st.success(f"✅ Archivo cargado: {uploaded_file.name}")
    
    if st.button("🚀 Generar PowerPoint"):
        try:
            with st.spinner("Analizando PDF con Gemini 1.5 Pro..."):
                # Guardar PDF temporalmente
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    tmp.write(uploaded_file.getvalue())
                    tmp_path = tmp.name
                
                # Subir archivo a Google
                g_file = genai.upload_file(tmp_path)
                
                # USAR EL MODELO CORRECTO
                model = genai.GenerativeModel("gemini-1.5-pro-latest")
                
                prompt = f"Analiza este PDF y crea una presentación de {num_slides} diapositivas. Responde en español con Títulos y Puntos Clave."
                response = model.generate_content([prompt, g_file])
                
                # Crear PPTX
                prs = Presentation()
                slide = prs.slides.add_slide(prs.slide_layouts[1])
                slide.shapes.title.text = "Análisis Médico IA"
                slide.placeholders[1].text = response.text[:1000] # Resumen inicial
                
                # Guardar y descargar
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_pptx:
                    prs.save(tmp_pptx.name)
                    with open(tmp_pptx.name, "rb") as f:
                        st.download_button(
                            label="📥 Descargar Presentación",
                            data=f,
                            file_name="presentacion_medica.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                
                os.unlink(tmp_path)
                
        except Exception as e:
            st.error(f"Error de API: {str(e)}")
            st.info("Consejo: Si el error persiste, intenta con un PDF más pequeño o verifica tu API Key.")
else:
    st.info("👆 Selecciona un archivo PDF para comenzar.")
