import streamlit as st
from groq import Groq
from pypdf import PdfReader
import tempfile
import os
import re
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# --- CONFIG ---
st.set_page_config(page_title="PPT Médico Robusto", page_icon="🏥")

if "GROQ_API_KEY" not in st.secrets:
    st.error("Falta la API Key en Secrets")
    st.stop()

client = Groq(api_key=st.secrets["GROQ_API_KEY"])

# --- ESTADO ---
if "slides" not in st.session_state: st.session_state.slides = []
if "pdf_text" not in st.session_state: st.session_state.pdf_text = ""
if "total_pedidas" not in st.session_state: st.session_state.total_pedidas = 60

# --- UI ---
st.title("🏥 Generador Médico Robusto")

if not st.session_state.pdf_text:
    st.session_state.total_pedidas = st.number_input("¿Cuántas diapositivas quieres?", 10, 100, 60)
    uploaded_file = st.file_uploader("Sube tu PDF médico", type="pdf")
    if uploaded_file:
        with st.spinner("Analizando PDF..."):
            reader = PdfReader(uploaded_file)
            text = ""
            for page in reader.pages:
                text += (page.extract_text() or "") + "\n"
            st.session_state.pdf_text = text
            st.rerun()

else:
    current = len(st.session_state.slides)
    total = st.session_state.total_pedidas
    
    st.write(f"### Progreso: {current} / {total} diapositivas")
    st.progress(min(current / total, 1.0))

    if current < total:
        if st.button(f"▶️ GENERAR BLOQUE (+5 SLIDES)"):
            with st.spinner("IA procesando bloque médico..."):
                try:
                    # Seleccionamos una parte del texto basada en el progreso
                    inicio = (current * 1500) % len(st.session_state.pdf_text)
                    segmento = st.session_state.pdf_text[inicio : inicio + 12000]
                    
                    prompt = f"""Actúa como un oncólogo experto. Genera exactamente 5 diapositivas técnicas basadas en el texto.
                    Usa ESTE FORMATO ESTRICTO para cada diapositiva:
                    
                    ---SLIDE---
                    TÍTULO: [Escribe aquí el título]
                    PUNTOS:
                    - [Punto técnico 1]
                    - [Punto técnico 2]
                    - [Punto técnico 3]
                    
                    TEXTO: {segmento}"""
                    
                    chat = client.chat.completions.create(
                        messages=[{"role": "system", "content": "Eres un asistente médico que solo responde en el formato ---SLIDE--- solicitado."},
                                  {"role": "user", "content": prompt}],
                        model="llama-3.1-8b-instant",
                        temperature=0.1, # Bajamos la temperatura para que sea más preciso
                        max_tokens=2500
                    )
                    
                    respuesta = chat.choices[0].message.content
                    
                    # Buscamos las diapositivas usando una expresión regular más flexible
                    nuevas_raw = re.split(r'---SLIDE---|---slide---|Slide \d+:', respuesta)
                    nuevas_limpias = [s.strip() for s in nuevas_raw if "TÍTULO:" in s.upper()]
                    
                    if nuevas_limpias:
                        st.session_state.slides.extend(nuevas_limpias[:5])
                        st.success(f"✅ ¡Bloque de {len(nuevas_limpias[:5])} slides añadido!")
                        st.rerun()
                    else:
                        st.error("La IA no usó el formato ---SLIDE---. Reintentando con un ajuste...")
                        st.text_area("Respuesta de la IA (para depurar):", respuesta, height=100)
                except Exception as e:
                    st.error(f"Error: {e}")

    # Botón para descargar lo que se lleve generado
    if current > 0:
        if st.button("🎨 DESCARGAR POWERPOINT AHORA"):
            with st.spinner("Construyendo archivo..."):
                prs = Presentation()
                prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
                
                for s_data in st.session_state.slides:
                    slide = prs.slides.add_slide(prs.slide_layouts[6])
                    
                    # Extraer título y puntos con Regex para mayor seguridad
                    titulo_match = re.search(r"TÍTULO:\s*(.*)", s_data, re.IGNORECASE)
                    titulo_text = titulo_match.group(1).strip() if titulo_match else "Diapositiva Médica"
                    
                    # Título
                    txt = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(1))
                    p = txt.text_frame.paragraphs[0]
                    p.text = titulo_text.upper()
                    p.font.size = Pt(28)
                    p.font.bold = True
                    p.font.color.rgb = RGBColor(0, 33, 71)
                    
                    # Puntos
                    body = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.5), Inches(5.5))
                    tf = body.text_frame
                    tf.word_wrap = True
                    
                    puntos = re.findall(r"-\s*(.*)", s_data)
                    for pt in puntos:
                        p_body = tf.add_paragraph()
                        p_body.text = "• " + pt.strip()
                        p_body.font.size = Pt(20)
                        p_body.space_after = Pt(10)

                with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
                    prs.save(tmp.name)
                    with open(tmp.name, "rb") as f:
                        st.download_button("📥 DESCARGAR PPTX", f, file_name="Clase_Medica.pptx")

    if st.button("🗑️ Reiniciar"):
        st.session_state.slides = []
        st.session_state.pdf_text = ""
        st.rerun()
