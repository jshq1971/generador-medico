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


# =========================
# CONFIG
# =========================
st.set_page_config(page_title="PPT Médico Pro", page_icon="🏥", layout="centered")
st.title("🏥 Generador de Presentaciones Médicas (Pro)")

if "GROQ_API_KEY" not in st.secrets:
    st.error("❌ Falta la GROQ_API_KEY en Secrets")
    st.stop()

client = Groq(api_key=st.secrets["GROQ_API_KEY"])

with st.sidebar:
    st.header("Opciones")
    num_slides = st.slider("Diapositivas (contenido)", 4, 20, 10)
    include_agenda = st.checkbox("Incluir Agenda (índice)", True)
    include_figures = st.checkbox("Añadir figuras del PDF (como imágenes)", True)
    max_fig_slides = st.slider("Cantidad de figuras a incluir", 0, 5, 2, disabled=not include_figures)
    include_flow = st.checkbox("Añadir 1 slide de algoritmo/flujo clínico", True)
    logo_file = st.file_uploader("Logo (PNG/JPG) opcional", type=["png", "jpg", "jpeg"])

uploaded_file = st.file_uploader("Sube tu PDF médico", type="pdf")


# =========================
# PPT HELPERS
# =========================
PRIMARY = RGBColor(0, 51, 102)       # Azul clínico
ACCENT = RGBColor(0, 102, 204)       # Azul acento
DARK = RGBColor(40, 40, 40)
LIGHT_BG = RGBColor(246, 248, 250)
WHITE = RGBColor(255, 255, 255)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


def _set_slide_bg(slide, rgb):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb


def _add_top_bar(slide, color=PRIMARY, height=Inches(0.18)):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), SLIDE_W, height)
    bar.fill.solid()
    bar.fill.fore_color.rgb = color
    bar.line.fill.background()


def _add_footer(slide, left_text="", right_text=""):
    # Línea fina
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.7), Inches(7.18), Inches(11.93), Inches(0.02))
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(220, 225, 230)
    line.line.fill.background()

    # Texto izq
    box_l = slide.shapes.add_textbox(Inches(0.7), Inches(7.22), Inches(8.5), Inches(0.3))
    tf = box_l.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = left_text
    p.font.size = Pt(10)
    p.font.color.rgb = RGBColor(120, 120, 120)

    # Texto der
    box_r = slide.shapes.add_textbox(Inches(9.5), Inches(7.22), Inches(3.1), Inches(0.3))
    tf = box_r.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = right_text
    p.alignment = PP_ALIGN.RIGHT
    p.font.size = Pt(10)
    p.font.color.rgb = RGBColor(120, 120, 120)


def _add_logo(slide, logo_path):
    if not logo_path:
        return
    try:
        slide.shapes.add_picture(logo_path, Inches(11.7), Inches(0.35), height=Inches(0.6))
    except Exception:
        pass


def _add_title(slide, title):
    tb = slide.shapes.add_textbox(Inches(0.9), Inches(0.55), Inches(11.2), Inches(1.1))
    tf = tb.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    return tb


def _add_bullets(slide, bullets, left=Inches(1.0), top=Inches(1.85), width=Inches(7.4), height=Inches(5.1)):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True

    for i, b in enumerate(bullets[:6]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"• {b}"
        p.level = 0
        p.font.size = Pt(20)
        p.font.color.rgb = DARK
        p.space_after = Pt(10)
        p.line_spacing = 1.15
    return box


def _add_callout(slide, text, left=Inches(8.7), top=Inches(2.05), width=Inches(3.8), height=Inches(1.2)):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = LIGHT_BG
    shp.line.color.rgb = RGBColor(210, 215, 220)
    tf = shp.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    p.alignment = PP_ALIGN.CENTER
    return shp


def _safe_filename(s):
    return re.sub(r"[^a-zA-Z0-9._-]+", "_", s)


# =========================
# LLM + PARSING
# =========================
def generate_slides_outline(pdf_text, n_slides):
    prompt = f"""
Eres un editor médico y diseñador de presentaciones para sesión clínica / congreso.
Genera EXACTAMENTE {n_slides} diapositivas de CONTENIDO (sin contar portada/agenda/cierre).

FORMATO OBLIGATORIO (repite para cada diapositiva):
---SLIDE---
TÍTULO: <máx 8 palabras, estilo título>
MENSAJE_CLAVE: <1 frase, máx 18 palabras>
CONTENIDO:
• <bullet 1, máx 14 palabras>
• <bullet 2, máx 14 palabras>
• <bullet 3, máx 14 palabras>

REGLAS:
- No pongas introducciones fuera del formato.
- Evita texto vago (“importante”, “relevante”).
- Enfatiza decisiones clínicas, criterios, perlas, riesgos, outcomes, cifras si están en el texto.
- Si faltan datos, no inventes números.

TEXTO:
{pdf_text[:22000]}
""".strip()

    completion = client.chat.completions.create(
        model="llama-3.1-8b-instant",
        messages=[
            {"role": "system", "content": "Responde estrictamente en el formato pedido."},
            {"role": "user", "content": prompt},
        ],
        temperature=0.4,
        max_tokens=4500,
    )
    return completion.choices[0].message.content


def parse_slides(text):
    blocks = [b.strip() for b in text.split("---SLIDE---") if b.strip()]
    slides = []
    for b in blocks:
        title = ""
        key = ""
        bullets = []
        mode = None
        for line in b.splitlines():
            line = line.strip()
            if not line:
                continue
            if line.startswith("TÍTULO:"):
                title = line.replace("TÍTULO:", "").strip()
                mode = None
            elif line.startswith("MENSAJE_CLAVE:"):
                key = line.replace("MENSAJE_CLAVE:", "").strip()
                mode = None
            elif line.startswith("CONTENIDO:"):
                mode = "bullets"
            elif mode == "bullets" and (line.startswith("•") or line.startswith("-")):
                bullets.append(line.lstrip("•-").strip())
        if title and bullets:
            slides.append({"title": title, "key": key, "bullets": bullets})
    return slides


def generate_flow_steps(pdf_text):
    prompt = f"""
A partir del texto, crea un algoritmo/flujo clínico en 6 pasos (máximo).
Devuelve SOLO una lista numerada (1., 2., 3...) sin explicaciones.

TEXTO:
{pdf_text[:18000]}
""".strip()

    completion = client.chat.completions.create(
        model="llama-3.1-8b-instant",
        messages=[
            {"role": "system", "content": "Sé clínico y concreto."},
            {"role": "user", "content": prompt},
        ],
        temperature=0.3,
        max_tokens=600,
    )
    raw = completion.choices[0].message.content.strip()
    steps = []
    for line in raw.splitlines():
        m = re.match(r"^\s*\d+\.\s+(.*)$", line.strip())
        if m:
            steps.append(m.group(1).strip())
    return steps[:6]


# =========================
# PDF FIGURES (RENDER PAGES)
# =========================
def pick_figure_pages(pdf_bytes, k):
    """Selecciona páginas con más imágenes embebidas; fallback a primeras páginas."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    scores = []
    for i in range(len(doc)):
        page = doc[i]
        img_count = len(page.get_images(full=True))
        scores.append((img_count, i))
    doc.close()
    scores.sort(reverse=True, key=lambda x: x[0])
    chosen = [i for cnt, i in scores if cnt > 0][:k]
    if not chosen:
        chosen = list(range(min(k, 3)))  # fallback suave
    return chosen[:k]


def render_page_to_png(pdf_bytes, page_index, zoom=2.0):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[page_index]
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    pix.save(tmp.name)
    doc.close()
    return tmp.name


# =========================
# MAIN
# =========================
if uploaded_file:
    st.success(f"✅ Archivo cargado: {uploaded_file.name}")

    if st.button("🚀 Generar PowerPoint Pro"):
        try:
            pdf_bytes = uploaded_file.getvalue()

            # Guardar logo opcional
            logo_path = None
            if logo_file is not None:
                t = tempfile.NamedTemporaryFile(delete=False, suffix="." + logo_file.name.split(".")[-1])
                t.write(logo_file.getvalue())
                t.close()
                logo_path = t.name

            with st.spinner("📄 Extrayendo texto del PDF..."):
                reader = PdfReader(uploaded_file)
                pdf_text = ""
                for page in reader.pages:
                    extracted = page.extract_text() or ""
                    pdf_text += extracted + "\n"

                if len(pdf_text.strip()) < 200:
                    st.warning("⚠️ El PDF tiene poco texto extraíble. Si es escaneado, puede salir pobre. (Aun así lo intento).")

            with st.spinner("🧠 Generando guión de diapositivas (IA)..."):
                outline = generate_slides_outline(pdf_text, num_slides)
                slide_items = parse_slides(outline)

                if len(slide_items) < max(3, num_slides // 2):
                    st.warning(
                        f"⚠️ La IA devolvió pocas diapositivas parseables ({len(slide_items)}). "
                        "Aun así generaré la PPT. Si quieres, luego endurecemos el formato."
                    )

            flow_steps = []
            if include_flow:
                with st.spinner("🧭 Creando algoritmo/flujo clínico..."):
                    flow_steps = generate_flow_steps(pdf_text)

            figure_pngs = []
            if include_figures and max_fig_slides > 0:
                with st.spinner("🖼️ Extrayendo figuras (render de páginas)..."):
                    pages = pick_figure_pages(pdf_bytes, max_fig_slides)
                    for p in pages:
                        figure_pngs.append((p, render_page_to_png(pdf_bytes, p)))

            with st.spinner("🎨 Armando PowerPoint profesional..."):
                prs = Presentation()
                prs.slide_width = SLIDE_W
                prs.slide_height = SLIDE_H

                doc_name = uploaded_file.name.replace(".pdf", "")
                today = date.today().isoformat()

                # PORTADA
                s = prs.slides.add_slide(prs.slide_layouts[6])
                _set_slide_bg(s, LIGHT_BG)
                _add_top_bar(s, PRIMARY, Inches(0.22))
                _add_logo(s, logo_path)

                tbox = s.shapes.add_textbox(Inches(1.2), Inches(2.2), Inches(11.0), Inches(2.0))
                tf = tbox.text_frame
                tf.clear()
                p = tf.paragraphs[0]
                p.text = "Resumen Clínico"
                p.font.size = Pt(46)
                p.font.bold = True
                p.font.color.rgb = PRIMARY
                p.alignment = PP_ALIGN.CENTER

                p = tf.add_paragraph()
                p.text = doc_name
                p.font.size = Pt(20)
                p.font.color.rgb = DARK
                p.alignment = PP_ALIGN.CENTER

                p = tf.add_paragraph()
                p.text = f"Generado automáticamente · {today}"
                p.font.size = Pt(14)
                p.font.color.rgb = RGBColor(110, 110, 110)
                p.alignment = PP_ALIGN.CENTER

                # AGENDA
                if include_agenda:
                    s = prs.slides.add_slide(prs.slide_layouts[6])
                    _set_slide_bg(s, WHITE)
                    _add_top_bar(s)
                    _add_logo(s, logo_path)
                    _add_footer(s, left_text=doc_name, right_text="Agenda")

                    _add_title(s, "Agenda")
                    titles = [it["title"] for it in slide_items[:10]]  # agenda compacta
                    _add_bullets(s, titles, left=Inches(1.0), top=Inches(1.8), width=Inches(11.3), height=Inches(5.2))

                # CONTENIDO
                for idx, it in enumerate(slide_items, start=1):
                    s = prs.slides.add_slide(prs.slide_layouts[6])
                    _set_slide_bg(s, WHITE)
                    _add_top_bar(s)
                    _add_logo(s, logo_path)
                    _add_footer(s, left_text=doc_name, right_text=f"{idx}/{len(slide_items)}")

                    _add_title(s, it["title"])
                    if it.get("key"):
                        _add_callout(s, it["key"], left=Inches(8.7), top=Inches(1.65), width=Inches(3.8), height=Inches(1.25)

                    _add_bullets(s, it["bullets"], left=Inches(1.0), top=Inches(2.1), width=Inches(11.3), height=Inches(5.0))

                # FLOW SLIDE
                if include_flow and len(flow_steps) >= 3:
                    s = prs.slides.add_slide(prs.slide_layouts[6])
                    _set_slide_bg(s, WHITE)
                    _add_top_bar(s, ACCENT)
                    _add_logo(s, logo_path)
                    _add_footer(s, left_text=doc_name, right_text="Algoritmo")

                    _add_title(s, "Algoritmo / Flujo clínico (resumen)")

                    left = Inches(1.2)
                    top = Inches(1.7)
                    box_w = Inches(10.9)
                    box_h = Inches(0.7)
                    gap = Inches(0.15)

                    for i, step in enumerate(flow_steps[:6]):
                        y = top + i * (box_h + gap)
                        shp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, y, box_w, box_h)
                        shp.fill.solid()
                        shp.fill.fore_color.rgb = LIGHT_BG
                        shp.line.color.rgb = RGBColor(210, 215, 220)

                        tf = shp.text_frame
                        tf.clear()
                        p = tf.paragraphs[0]
                        p.text = f"{i+1}. {step}"
                        p.font.size = Pt(16)
                        p.font.bold = True if i == 0 else False
                        p.font.color.rgb = PRIMARY

                        # Flecha entre pasos
                        if i < min(5, len(flow_steps[:6]) - 1):
                            arr = s.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(6.4), y + box_h, Inches(0.5), gap)
                            arr.fill.solid()
                            arr.fill.fore_color.rgb = ACCENT
                            arr.line.fill.background()

                # FIGURE SLIDES
                if include_figures and figure_pngs:
                    for (pidx, png_path) in figure_pngs:
                        s = prs.slides.add_slide(prs.slide_layouts[6])
                        _set_slide_bg(s, WHITE)
                        _add_top_bar(s, PRIMARY)
                        _add_logo(s, logo_path)
                        _add_footer(s, left_text=doc_name, right_text=f"Figura (pág. {pidx+1})")

                        _add_title(s, f"Figura / Tabla (pág. {pidx+1})")

                        # Imagen centrada
                        # Ajuste simple para que quepa dentro del área
                        s.shapes.add_picture(png_path, Inches(1.0), Inches(1.7), width=Inches(11.33))

                # CIERRE
                s = prs.slides.add_slide(prs.slide_layouts[6])
                _set_slide_bg(s, PRIMARY)
                _add_logo(s, logo_path)

                box = s.shapes.add_textbox(Inches(2.0), Inches(2.9), Inches(9.3), Inches(1.5))
                tf = box.text_frame
                tf.clear()
                p = tf.paragraphs[0]
                p.text = "Gracias"
                p.font.size = Pt(54)
                p.font.bold = True
                p.font.color.rgb = WHITE
                p.alignment = PP_ALIGN.CENTER

                p = tf.add_paragraph()
                p.text = "¿Preguntas?"
                p.font.size = Pt(22)
                p.font.color.rgb = WHITE
                p.alignment = PP_ALIGN.CENTER

            # Guardar y descargar
            out_name = f"presentacion_medica_pro_{_safe_filename(doc_name)}.pptx"
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
                prs.save(tmp.name)
                with open(tmp.name, "rb") as f:
                    st.success("✅ Presentación generada (Pro).")
                    st.download_button(
                        "📥 Descargar PowerPoint",
                        data=f,
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )

            # Limpieza
            try:
                os.unlink(tmp.name)
            except Exception:
                pass

            if logo_path:
                try:
                    os.unlink(logo_path)
                except Exception:
                    pass

            for _, png in figure_pngs:
                try:
                    os.unlink(png)
                except Exception:
                    pass

        except Exception as e:
            st.error(f"Error: {e}")
else:
    st.info("⬆️ Sube un PDF para comenzar.")
