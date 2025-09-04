import streamlit as st
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from io import BytesIO
from PIL import Image
import tempfile
import os

# Para la conversión PDF -> DOCX usamos pdf2docx (asegúrate de instalarlo)
try:
    from pdf2docx import Converter
    HAS_PDF2DOCX = True
except Exception:
    HAS_PDF2DOCX = False

st.set_page_config(page_title="Herramientas PDF", layout="centered")

st.image("assets/combinado.png")
st.title("Herramientas para PDF e Imágenes")
st.write("Usa el menú lateral para seleccionar la operación que deseas realizar.")

menu = st.sidebar.selectbox("Selecciona una opción", [
    "Unir PDFs",
    "Dividir PDF",
    "PDF -> Word (DOCX)",
    "Imagen -> PDF"
])

# ------------------ FUNCIONES AUXILIARES ------------------

def merge_pdfs(uploaded_files: list) -> BytesIO:
    """Unir varios archivos PDF (lista de UploadedFile) y devolver BytesIO listo para descargar."""
    merger = PdfMerger()
    for f in uploaded_files:
        # PyPDF2 acepta objetos file-like directamente
        merger.append(f)
    output = BytesIO()
    merger.write(output)
    merger.close()
    output.seek(0)
    return output


def split_pdf_every_page(uploaded_pdf) -> list:
    """Divide un PDF y devuelve una lista de tuples (nombre, BytesIO) con cada página como PDF independiente."""
    reader = PdfReader(uploaded_pdf)
    results = []
    num_pages = len(reader.pages)
    for i in range(num_pages):
        writer = PdfWriter()
        writer.add_page(reader.pages[i])
        out = BytesIO()
        writer.write(out)
        out.seek(0)
        results.append((f"page_{i+1}.pdf", out))
    return results


def split_pdf_range(uploaded_pdf, start: int, end: int) -> BytesIO:
    """Extrae páginas desde start hasta end (inclusive, 1-based) y devuelve un BytesIO con ese PDF."""
    reader = PdfReader(uploaded_pdf)
    writer = PdfWriter()
    num_pages = len(reader.pages)
    # ajustar índices
    start_idx = max(0, start - 1)
    end_idx = min(end - 1, num_pages - 1)
    for i in range(start_idx, end_idx + 1):
        writer.add_page(reader.pages[i])
    out = BytesIO()
    writer.write(out)
    out.seek(0)
    return out


def pdf_to_docx_bytesio(uploaded_pdf) -> BytesIO:
    """Convierte un archivo PDF (UploadedFile or file-like) a DOCX usando pdf2docx. Devuelve BytesIO."""
    if not HAS_PDF2DOCX:
        raise RuntimeError("pdf2docx no está instalado. Instálalo con: pip install pdf2docx")
    # Guardamos el PDF a un archivo temporal porque pdf2docx trabaja con rutas
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_pdf:
        tmp_pdf.write(uploaded_pdf.read())
        tmp_pdf_path = tmp_pdf.name
    tmp_docx_path = tmp_pdf_path.replace('.pdf', '.docx')
    try:
        cv = Converter(tmp_pdf_path)
        cv.convert(tmp_docx_path, start=0, end=None)
        cv.close()
        # leer el docx en memoria
        with open(tmp_docx_path, 'rb') as f:
            data = f.read()
        bio = BytesIO(data)
        bio.seek(0)
        return bio
    finally:
        # limpiar archivos temporales
        try:
            os.remove(tmp_pdf_path)
        except Exception:
            pass
        try:
            if os.path.exists(tmp_docx_path):
                os.remove(tmp_docx_path)
        except Exception:
            pass


def image_to_pdf_bytesio(uploaded_image) -> BytesIO:
    """Convierte una imagen (png/jpg/jpeg) a PDF y devuelve BytesIO."""
    image = Image.open(uploaded_image)
    # Convertir a RGB si tiene canal alfa
    if image.mode in ("RGBA", "LA"):
        rgb = Image.new("RGB", image.size, (255, 255, 255))
        rgb.paste(image, mask=image.split()[-1])
        image = rgb
    else:
        image = image.convert("RGB")
    out = BytesIO()
    image.save(out, format="PDF")
    out.seek(0)
    return out

# ------------------ OPCIONES DEL MENU ------------------

if menu == "Unir PDFs":
    st.header("Unir archivos PDF")
    uploaded_files = st.file_uploader("Sube varios PDFs (Ctrl/Shift para seleccionar múltiples)", type=['pdf'], accept_multiple_files=True)
    if st.button("Unir archivos"):
        if not uploaded_files or len(uploaded_files) < 2:
            st.warning("Selecciona al menos 2 archivos PDF para unir.")
        else:
            merged = merge_pdfs(uploaded_files)
            st.success("PDFs unidos correctamente. Descarga el resultado abajo.")
            st.download_button("Descargar PDF unido", data=merged.getvalue(), file_name="pdf_unido.pdf", mime='application/pdf')

elif menu == "Dividir PDF":
    st.header("Dividir un PDF")
    uploaded_pdf = st.file_uploader("Sube un PDF", type=['pdf'])
    if uploaded_pdf:
        option = st.radio("Modo de división:", ["Separar cada página", "Extraer un rango de páginas"])
        if option == "Separar cada página":
            if st.button("Dividir en páginas individuales"):
                parts = split_pdf_every_page(uploaded_pdf)
                st.success(f"Se generaron {len(parts)} archivos PDF (uno por página).")
                for name, bio in parts:
                    st.download_button(f"Descargar {name}", data=bio.getvalue(), file_name=name, mime='application/pdf')
        else:
            cols = st.columns(2)
            start = cols[0].number_input("Página inicio (1-based)", min_value=1, value=1, step=1)
            # leer número de páginas para limitar end
            try:
                reader = PdfReader(uploaded_pdf)
                max_p = len(reader.pages)
            except Exception:
                max_p = None
            end_default = max_p if max_p else 1
            end = cols[1].number_input("Página fin (1-based)", min_value=1, value=end_default, step=1)
            if st.button("Extraer rango"):
                try:
                    out = split_pdf_range(uploaded_pdf, start, end)
                    st.success(f"Se extrajo el rango {start}-{end} correctamente.")
                    st.download_button("Descargar rango extraído", data=out.getvalue(), file_name=f"rango_{start}_{end}.pdf", mime='application/pdf')
                except Exception as e:
                    st.error(f"Error al extraer el rango: {e}")

elif menu == "PDF -> Word (DOCX)":
    st.header("Convertir PDF a Word (DOCX)")
    if not HAS_PDF2DOCX:
        st.warning("La biblioteca 'pdf2docx' no está disponible. Instálala: pip install pdf2docx")
    uploaded_pdf = st.file_uploader("Sube un PDF", type=['pdf'])
    if uploaded_pdf and st.button("Convertir a DOCX"):
        if not HAS_PDF2DOCX:
            st.error("Imposible convertir: instala pdf2docx y recarga la app.")
        else:
            try:
                docx_bio = pdf_to_docx_bytesio(uploaded_pdf)
                st.success("Conversión completada. Descarga el DOCX abajo.")
                st.download_button("Descargar DOCX", data=docx_bio.getvalue(), file_name="documento_convertido.docx", mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            except Exception as e:
                st.error(f"Error al convertir: {e}")

elif menu == "Imagen -> PDF":
    st.header("Convertir imagen a PDF")
    uploaded_image = st.file_uploader("Sube una imagen (png/jpg/jpeg)", type=['png','jpg','jpeg'])
    if uploaded_image and st.button("Convertir imagen a PDF"):
        try:
            pdf_bio = image_to_pdf_bytesio(uploaded_image)
            st.success("Imagen convertida a PDF.")
            st.download_button("Descargar PDF", data=pdf_bio.getvalue(), file_name="imagen_convertida.pdf", mime='application/pdf')
        except Exception as e:
            st.error(f"Error al convertir la imagen: {e}")

# ------------------ PIE DE PÁGINA ------------------

st.markdown("---")
st.write("Consejos: si la conversión PDF->DOCX falla en casos complejos (PDFs con muchas imágenes o tablas), considera herramientas externas o servicios especializados.")
