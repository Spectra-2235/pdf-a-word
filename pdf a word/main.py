import streamlit as st
import fitz  # PyMuPDF
from PIL import Image
from io import BytesIO
from docx import Document
from docx.shared import Inches
import tempfile
import os

st.set_page_config(page_title="PDF a Word", layout="centered")
st.title("📄 Convertidor PDF a Word")

pdf_file = st.file_uploader("Sube tu PDF", type=["pdf"])

if pdf_file:
    pdf = fitz.open(stream=pdf_file.read(), filetype="pdf")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        doc = Document()

        # Configurar tamaño carta y márgenes mínimos
        section = doc.sections[-1]
        section.page_width = Inches(8.5)
        section.page_height = Inches(11)
        section.left_margin = Inches(0.2)
        section.right_margin = Inches(0.2)
        section.top_margin = Inches(0.2)
        section.bottom_margin = Inches(0.2)

        usable_width = section.page_width - section.left_margin - section.right_margin

        for i, page in enumerate(pdf):
            # Renderizar la página como imagen
            pix = page.get_pixmap(dpi=300)
            img = Image.open(BytesIO(pix.tobytes("png")))

            # Guardar imagen temporal
            img_path = os.path.join(tempfile.gettempdir(), f"page_{i}.png")
            img.save(img_path)

            # Insertar imagen a ancho completo (alto se ajusta automáticamente)
            doc.add_picture(img_path, width=usable_width)

            # Salto de página si no es la última
            if i < len(pdf) - 1:
                doc.add_page_break()

        doc.save(tmp.name)

        st.download_button(
            label="⬇️ Descargar Word",
            data=open(tmp.name, "rb").read(),
            file_name="PDF.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )