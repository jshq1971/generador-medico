import streamlit as st
from groq import Groq
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pypdf import PdfReader
import tempfile
import os

st.set_page_config(page_title="IA Médica PPTX", page_icon="🏥")
st.title("🏥 Generador de Presentaciones Médicas Profesionales")

# Verificar API KEY
if "GROQ_API_KEY" not in st.secrets:
    st.error("❌ Falta la GROQ_API_KEY en Secrets")
    st.stop()

client = Groq(api_key=st.secrets["GROQ_API_KEY"])

num_slides = st.slider("Número de diapositivas", 3, 15, 8)
uploaded_file = st.file_uploader("Sube tu PDF médico", type="pdf")

if uploaded_file:
    st.success(f"✅ Archivo cargado: {uploaded_file.name}")

    if st.button("🚀 Generar PowerPoint Profesional"):
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
                            "content": "Eres un especialista médico experto en crear presentaciones profesionales para congresos y sesiones clínicas."
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
• [punto clave 4]

REGLAS:
- Máximo 4 puntos por slide
- Cada punto debe ser conciso (máximo 15 palabras)
- Usa lenguaje médico preciso
- Enfócate en lo más relevante clínicamente

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

            # =========================
            # CREAR PRESENTACIÓN PRO
            # =========================
            with st.spinner("📊 Creando presentación profesional..."):
                prs = Presentation()

                # Tamaño 16:9
                prs.slide_width = Inches(13.33)
                prs.slide_height = Inches(7.5)

                # COLORES MÉDICOS
                PRIMARY = RGBColor(0, 51, 102)     # Azul clínico
                SECONDARY = RGBColor(64, 64, 64)   # Gris oscuro
                ACCENT = RGBColor(0, 102, 204)     # Azul claro

                # -------------------------
                # SLIDE DE PORTADA
                # -------------------------
                slide = prs.slides.add_slide(prs.slide_layouts[6])

                # Fondo de color
                background = slide.background
                fill = background.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(245, 245, 245)

                # Título principal
                title_box = slide.shapes.add_textbox(
                    Inches(1.5), Inches(2.5), Inches(10), Inches(1.5)
                )
                tf = title_box.text_frame
                p = tf.paragraphs[0]
                p.text = "Análisis Clínico"
                p.font.size = Pt(44)
                p.font.bold = True
                p.font.color.rgb = PRIMARY
                p.alignment = PP_ALIGN.CENTER

                # Subtítulo
                p = tf.add_paragraph()
                p.text = uploaded_file.name.replace(".pdf", "")
                p.font.size = Pt(20)
                p.font.color.rgb = SECONDARY
                p.alignment = PP_ALIGN.CENTER

                # Línea decorativa
                line = slide.shapes.add_shape(
                    1,  # Line shape
                    Inches(4), Inches(4.5), Inches(5.33), Inches(0)
                )
                line.line.color.rgb = ACCENT
                line.line.width = Pt(3)

                # -------------------------
                # SLIDES DE CONTENIDO
                # -------------------------
                slides_data = response_text.split("---SLIDE---")

                for slide_text in slides_data:
                    if not slide_text.strip():
                        continue

                    title = ""
                    bullets = []

                    for line in slide_text.split("\n"):
                        if line.startswith("TÍTULO:"):
                            title = line.replace("TÍTULO:", "").strip()
                        elif line.strip().startswith("•"):
                            bullets.append(line.replace("•", "").strip())
                        elif line.strip().startswith("-"):
                            bullets.append(line.replace("-", "").strip())

                    if not title or not bullets:
                        continue

                    slide = prs.slides.add_slide(prs.slide_layouts[6])

                    # Fondo
                    background = slide.background
                    fill = background.fill
                    fill.solid()
                    fill.fore_color.rgb = RGBColor(255, 255, 255)

                    # Barra superior de color
                    header = slide.shapes.add_shape(
                        1,  # Rectangle
                        Inches(0), Inches(0), Inches(13.33), Inches(0.15)
                    )
                    header.fill.solid()
                    header.fill.fore_color.rgb = PRIMARY
                    header.line.fill.background()

                    # TÍTULO
                    title_box = slide.shapes.add_textbox(
                        Inches(0.8), Inches(0.6), Inches(11.5), Inches(1.2)
                    )
                    tf = title_box.text_frame
                    p = tf.paragraphs[0]
                    p.text = title
                    p.font.size = Pt(32)
                    p.font.bold = True
                    p.font.color.rgb = PRIMARY

                    # CONTENIDO
                    content_box = slide.shapes.add_textbox(
                        Inches(1.2), Inches(2.2), Inches(11), Inches(4.5)
                    )
                    tf = content_box.text_frame
                    tf.clear()

                    for i, bullet in enumerate(bullets[:4]):
                        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
                        p.text = "• " + bullet
                        p.font.size = Pt(22)
                        p.font.color.rgb = RGBColor(0, 0, 0)
                        p.space_after = Pt(14)
                        p.line_spacing = 1.3

                # -------------------------
                # SLIDE DE CIERRE
                # -------------------------
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                
                background = slide.background
                fill = background.fill
                fill.solid()
                fill.fore_color.rgb = PRIMARY

                closing_box = slide.shapes.add_textbox(
                    Inches(2), Inches(3), Inches(9.33), Inches(1.5)
                )
                tf = closing_box.text_frame
                p = tf.paragraphs[0]
                p.text = "Gracias"
                p.font.size = Pt(48)
                p.font.bold = True
                p.font.color.rgb = RGBColor(255, 255, 255)
                p.alignment = PP_ALIGN.CENTER

                # Guardar
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
                    prs.save(tmp.name)
                    with open(tmp.name, "rb") as f:
                        st.success(f"✅ Presentación creada con {len(prs.slides)} diapositivas")
                        st.download_button(
                            "📥 Descargar PowerPoint Profesional",
                            f,
                            file_name="presentacion_medica_profesional.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                os.unlink(tmp.name)

        except Exception as e:
            st.error(f"Error: {e}")
            st.info("Si el error persiste, intenta con un PDF más pequeño.")
else:
    st.info("⬆️ Sube un PDF médico para comenzar")
