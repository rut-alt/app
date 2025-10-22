import streamlit as st
import pandas as pd
from io import BytesIO
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, BooleanObject, TextStringObject

# --- FUNCION PARA RELLENAR PDF ---
def fill_pdf(input_pdf, data_dict):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    # Copiar todas las páginas del PDF original
    for page in reader.pages:
        writer.add_page(page)

    # Copiar AcroForm si existe, si no crear uno
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

    # Rellenar los campos con los valores del Excel
    writer.update_page_form_field_values(writer.pages[0], data_dict)

    # Forzar que se vean los valores en Adobe
    writer._root_object["/AcroForm"].update({
        NameObject("/NeedAppearances"): BooleanObject(True)
    })

    output = BytesIO()
    writer.write(output)
    output.seek(0)
    return output


# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Generador PDF Climagas", page_icon="🧾", layout="centered")
st.title("🧾 Generador de PDF Climagas")

st.markdown("""
Sube tu **plantilla PDF** y el **Excel con los datos del producto** para generar automáticamente un PDF con los campos rellenados.
""")

# Subida de archivos
pdf_file = st.file_uploader("📄 Sube la plantilla PDF", type="pdf")
excel_file = st.file_uploader("📊 Sube el Excel con los productos", type=["xls", "xlsx"])

# Entrada del código de producto
producto = st.number_input("Introduce el código del producto", step=1, min_value=0)

# Mapeo PDF -> Excel
pdf_to_excel_map = {
    "Potencia Nominal kW": "Potencia Nominal kW",
    "Fecha de solicitud de licencia de obra en su defec": "Fecha de solicitud de licencia de obra en su defect",
    "Potencia total inicial": "Potencia total inicial",
    "Potencia a modificar": "Potencia a modificar",
    "Potencia total final": "Potencia total final",
    "Administrativo": "Administrativo"
}

# Botón principal
if st.button("🚀 Generar PDF"):
    if not pdf_file or not excel_file:
        st.error("Por favor, sube el PDF y el Excel antes de continuar.")
    else:
        try:
            df = pd.read_excel(excel_file)

            if producto not in df['PRODUCTO'].values:
                st.error("El código del producto no se encuentra en el Excel.")
            else:
                row = df.loc[df['PRODUCTO'] == producto].iloc[0]
                data_dict = {}
                for pdf_field, excel_col in pdf_to_excel_map.items():
                    if excel_col in row:
                        data_dict[pdf_field] = str(row[excel_col])

                # Generar PDF con campos visibles
                output_pdf = fill_pdf(pdf_file, data_dict)

                st.success(f"✅ PDF generado correctamente para el producto {producto}")

                st.download_button(
                    label="⬇️ Descargar PDF con campos rellenados",
                    data=output_pdf,
                    file_name=f"pdf_producto_{producto}.pdf",
                    mime="application/pdf"
                )

        except Exception as e:
            st.error(f"❌ Error al generar el PDF: {e}")
