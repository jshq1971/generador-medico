import streamlit as st
from groq import Groq
from pypdf import PdfReader
import fitz  # PyMuPDF
import tempfile
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# --- CONFIG ---
st.set_page_config(page_title="PPT Médico Estable", page_icon="🏥")

if "GROQ_API_KEY" not in st.secrets:
    st.error("Falta la API Key en Secrets")
    st.stop()

client = Groq(api_key=st.secrets["GROQ_API_KEY"])

# --- ESTADO ---
if "slides" not in st.session_state: st.session_state.slides = []
if "pdf_text" not in st.session_state: st.session_state.pdf_text = ""
if "total_pedidas" not in st.session_state: st.session_state.total_pedidas = 60

# --- UI ---
st.title("🏥 Generador Médico Estable")

# 1. Configuración Inicial
if not st.session_state.pdf_text:
    st.session_state.total_pedidas = st.number_input("¿Cuántas diapositivas quieres en total?", 10, 100, 60)
    uploaded_file = st.file_uploader("Sube tu PDF médico", type="pdf")
    if uploaded_file:
        with st.spinner("Leyendo PDF..."):
            reader = PdfReader(uploaded_file)
            text = ""
            for page in reader.pages:
                text += (page.extract_text() or "") + "\n"
            st.session_state.pdf_text = text
            st.session_state.pdf_bytes = uploaded_file.getvalue()
            st.rerun()

# 2. Generación por Bloques
else:
    current = len(st.session_state.slides)
    total = st.session_state.total_pedidas
    
    st.write(f"### Progreso: {current} / {total} diapositivas")
    st.progress(min(current / total, 1.0))

    if current < total:
        if st.button(f"▶️ GENERAR SIGUIENTE BLOQUE (+5 SLIDES)"):
            with st.spinner("La IA está escribiendo..."):
                try:
                    # Llamada ultra-rápida
                    segmento = st.session_state.pdf_text[(current*2000):(current*2000)+10000]
                    prompt = f"Eres oncólogo. Genera 5 diapositivas médicas. Formato: ---SLIDE--- TÍTULO: [título] PUNTOS: - [punto 1] - [punto 2]. Texto: {segmento}"
                    
                    chat = client.chat.completions.create(
                        messages=[{"role": "user", "content": prompt}],
                        model="llama-3.1-8b-instant",
                        temperature=0.2,
                        max_tokens=2000
                    )
                    
                    respuesta = chat.choices[0].message.content
                    nuevas = [s for s in respuesta.split("---SLIDE---") if "TÍTULO:" in s]
                    
                    if nuevas:
                        st.session_state.slides.extend(nuevas)
                        st.success(f"¡Bloque generado! Ahora tienes {len(st.session_state.slides)} slides.")
                        st.rerun()
                    else:
                        st.error("La IA no respondió en el formato correcto. Intenta de nuevo.")
                except Exception as e:
                    st.error(f"Error de conexión: {e}")

    # 3. Finalización
    if current >= total or (current > 0 and st.button("🏁 Finalizar ahora y descargar")):
        if st.button("🎨 CREAR ARCHIVO POWERPOINT"):
            with st.spinner("Creando PPTX..."):
                prs = Presentation()
                prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
                
                for s_data in st.session_state.slides:
                    slide = prs.slides.add_slide(prs.slide_layouts[6])
                    lines = s_data.strip().split("\n")
                    
                    # Título
                    title_text = lines[0].replace("TÍTULO:", "").strip()
                    txt = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(1))
                    p = txt.text_frame.paragraphs[0]
                    p.text = title_text.upper()
                    p.font.size = Pt(28)
                    p.font.bold = True
                    p.font.color.rgb = RGBColor(0, 33, 71)
                    
                    # Puntos
                    body = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.5), Inches(5))
                    for l in lines:
                        if l.strip().startswith("-"):
                            p_body = body.text_frame.add_paragraph()
                            p_body.text = "• " + l.strip().lstrip("-").strip()
                            p_body.font.size = Pt(20)

                with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
                    prs.save(tmp.name)
                    with open(tmp.name, "rb") as f:
                        st.download_button("📥 DESCARGAR AHORA", f, file_name="Presentacion_Medica.pptx")

    if st.button("🗑️ Reiniciar todo"):
        st.session_state.slides = []
        st.session_state.pdf_text = ""
        st.rerun()
