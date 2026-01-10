import streamlit as st
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches
import tempfile
import os

# Configuración de página
st.set_page_config(page_title="IA Médica PPTX", page_icon="🏥")

st.title("🏥 Generador de Presentaciones Médicas")

# 1. Verificar API KEY
if "GEMINI_API_KEY" not in st.secrets:
    st.error("⚠️ Falta la GEMINI_API_KEY en los Secrets de Streamlit.")
    st.stop()

genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

# 2. Interfaz de usuario
num_slides = st.slider("Número de diapositivas", 1, 20, 10)
uploaded_file = st.file_uploader("Sube tu PDF médico", type="pdf")

if uploaded_file is not None:
    st.success(f"✅ Archivo cargado: {uploaded_file.name}")
    
    if st.button("🚀 Generar PowerPoint"):
        try:
            with st.spinner("Analizando PDF con IA..."):
                # Guardar temporalmente para subir a Gemini
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    tmp.write(uploaded_file.getvalue())
                    tmp_path = tmp.name
                
                # Subir a Gemini
                model = genai.GenerativeModel("gemini-1.5-flash") # Flash es más rápido para esto
                g_file = genai.upload_file(tmp_path)
                
                prompt = f"Analiza este PDF y crea una presentación de {num_slides} diapositivas. Para cada diapositiva dame: Título y Contenido (3 puntos clave). Responde en español."
                response = model.generate_content([prompt, g_file])
                
                # Crear el PPTX
                prs = Presentation()
                
                # Diapositiva de Título
                slide = prs.slides.add_slide(prs.slide_layouts[0])
                slide.shapes.title.text = "Resumen Médico IA"
                slide.placeholders[1].text = f"Basado en: {uploaded_file.name}"
                
                # Diapositiva de Contenido (Resumen de la IA)
                slide_cont = prs.slides.add_slide(prs.slide_layouts[1])
                slide_cont.shapes.title.text = "Análisis del Documento"
                tf = slide_cont.placeholders[1].text_frame
                tf.text = response.text[:500] # Ponemos el texto de la IA
                
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
                
                # Limpiar archivos temporales
                os.unlink(tmp_path)
                
        except Exception as e:
            st.error(f"Hubo un problema: {str(e)}")
else:
    st.info("👆 Por favor, selecciona un archivo PDF para comenzar.")
