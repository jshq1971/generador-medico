import streamlit as st
import google.generativeai as genai
import json
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import tempfile

# Configuración
st.set_page_config(page_title="Asistente Médico IA", page_icon="🏥")

st.title("🏥 Generador de Presentaciones Médicas")
st.write("Sube tus PDFs y genera diapositivas profesionales.")

# Cargar API Key desde Secrets
if "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    
    num_slides = st.slider("Número de diapositivas", 1, 20, 10)
    uploaded_files = st.file_uploader("Sube tus PDFs", type="pdf", accept_multiple_files=True)
    
    if st.button("🚀 Generar PowerPoint"):
        if uploaded_files:
            try:
                with st.spinner("Analizando contenido..."):
                    model = genai.GenerativeModel("gemini-1.5-pro")
                    # Procesar archivos
                    google_files = []
                    for f in uploaded_files:
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                            tmp.write(f.getvalue())
                            g_file = genai.upload_file(tmp.name)
                            google_files.append(g_file)
                    
                    prompt = f"Genera un JSON Array de {num_slides} diapositivas médicas. Responde SOLO el JSON puro."
                    response = model.generate_content([prompt] + google_files)
                    
                    # Crear PPTX simple
                    prs = Presentation()
                    slide = prs.slides.add_slide(prs.slide_layouts[6])
                    title = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(10), Inches(2))
                    title.text_frame.text = "PRESENTACIÓN GENERADA"
                    
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_pptx:
                        prs.save(tmp_pptx.name)
                        with open(tmp_pptx.name, "rb") as f:
                            st.success("✅ ¡Listo!")
                            st.download_button("📥 Descargar PowerPoint", f, "presentacion.pptx")
            except Exception as e:
                st.error(f"Error: {e}")
        else:
            st.warning("Por favor, sube al menos un PDF.")
else:
    st.error("Falta la GEMINI_API_KEY en los Secrets de Streamlit.")
