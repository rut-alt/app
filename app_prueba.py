import streamlit as st
import pandas as pd
from io import BytesIO
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, TextStringObject, BooleanObject

# --- FUNCION PRINCIPAL PARA RELLENAR PDF ---
def fill_pdf_text_only(input_pdf, data_dict):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    # Copiar todas las páginas
    for page in reader.pages:
        writer.add_page(page)

    # Forzar NeedAppearances en AcroForm
    if "/AcroForm" in reader.trailer["/Root"]:
        writer._root_object.update({
            NameObject("/AcroForm"): reader.trailer["/Root"]["/AcroForm"]
        })
    else:
        from pypdf.generic import DictionaryObject, ArrayObject
        writer._root_object.update({
            NameObject("/AcroForm"): DictionaryObject({
                NameObject("/Fields"): ArrayObject(),
                NameObject("/NeedAppearances"): BooleanObject(True)
            })
        })

    writer._root_object["/AcroForm"].update({
        NameObject("/NeedAppearances"): BooleanObject(True)
    })

    # Rellenar solo campos de texto
    for page in writer.pages:
        annots = page.get("/Annots")
        if not annots:
            continue
        for annot in annots:
            field = annot.get_object()
            field_name = field.get("/T")
            if not field_name or field_name not in data_dict:
                continue

            if field.get("/FT") == "/Tx":  # Solo campos de texto
                field.update({NameObject("/V"): TextStringObject(str(data_dict[field_name]))})

    # Guardar PDF en memoria
    output = BytesIO()
    writer.write(output)
    output.seek(0)
    return output

# --- STREAMLIT ---
st.set_page_config(page_title="Generador PDF Climagas", page_icon="🧾", layout="centered")
st.title("🧾 Generador de PDF Climagas")
st.markdown("Sube tu PDF plantilla y Excel con productos para generar un PDF rellenando solo campos de texto.")

pdf_file = st.file_uploader("📄 Sube la plantilla PDF", type="pdf")
excel_file = st.file_uploader("📊 Sube el Excel con los productos", type=["xls", "xlsx"])
producto = st.number_input("Introduce el código del producto", step=1, min_value=0)

# Mapeo PDF -> Excel (solo texto)
pdf_to_excel_map = {
    "Potencia Nominal kW": "Potencia Nominal kW",
    "Fecha de solicitud de licencia de obra en su defec": "Fecha de solicitud de licencia de obra en su defect",
    "Potencia total inicial": "Potencia total inicial",
    "Potencia a modificar": "Potencia a modificar",
    "Potencia total final": "Potencia total final"
}

if st.button("🚀 Generar PDF"):
    if not pdf_file or not excel_file:
        st.error("Por favor, sube ambos archivos antes de continuar.")
    else:
        try:
            df = pd.read_excel(excel_file)
            if producto not in df['PRODUCTO'].values:
                st.error("El código del producto no se encuentra en el Excel.")
            else:
                row = df.loc[df['PRODUCTO'] == producto].iloc[0]
                data_dict = {pdf_field: str(row[excel_col]) for pdf_field, excel_col in pdf_to_excel_map.items() if excel_col in row}

                output_pdf = fill_pdf_text_only(pdf_file, data_dict)

                st.success(f"✅ PDF generado correctamente para el producto {producto}")
                st.download_button(
                    label="⬇️ Descargar PDF con campos de texto rellenados",
                    data=output_pdf,
                    file_name=f"pdf_producto_{producto}.pdf",
                    mime="application/pdf"
                )
        except Exception as e:
            st.error(f"❌ Error al generar el PDF: {e}")
