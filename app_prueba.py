import streamlit as st
import pandas as pd
from io import BytesIO
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, TextStringObject, BooleanObject

st.set_page_config(page_title="Generador PDF Climagas", page_icon="🧾", layout="centered")
st.title("🧾 Generador de PDF Climagas")

st.markdown("Sube tu plantilla PDF y el archivo Excel con los datos del producto.")

# Subida de archivos
pdf_file = st.file_uploader("📄 Sube la plantilla PDF", type="pdf")
excel_file = st.file_uploader("📊 Sube el Excel con los productos", type=["xls", "xlsx"])

# Entrada de producto
producto = st.number_input("Introduce el código del producto", step=1, min_value=0)

# Mapeo de campos
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
            # Leer Excel
            df = pd.read_excel(excel_file)
            if producto not in df['PRODUCTO'].values:
                st.error("El código del producto no se encuentra en el Excel.")
            else:
                reader = PdfReader(pdf_file)
                writer = PdfWriter()

                for page in reader.pages:
                    writer.add_page(page)

                for page in writer.pages:
                    annots = page.get("/Annots")
                    if not annots:
                        continue
                    for annot in annots:
                        field = annot.get_object()
                        field_name = field.get("/T")
                        if not field_name or field_name not in pdf_to_excel_map:
                            continue

                        excel_col = pdf_to_excel_map[field_name]
                        value = df.loc[df['PRODUCTO'] == producto, excel_col].values[0]

                        if field.get("/FT") == "/Btn":  # Checkbox
                            field.update({
                                NameObject("/V"): NameObject("/Yes") if value else NameObject("/Off"),
                                NameObject("/AS"): NameObject("/Yes") if value else NameObject("/Off")
                            })
                        else:  # Texto
                            field.update({NameObject("/V"): TextStringObject(str(value))})
                            field.update({NameObject("/Ff"): BooleanObject(True)})

                # Guardar PDF en memoria
                output_pdf = BytesIO()
                writer.write(output_pdf)
                output_pdf.seek(0)

                st.success(f"✅ PDF generado correctamente para el producto {producto}")

                # Descargar el PDF
                st.download_button(
                    label="⬇️ Descargar PDF generado",
                    data=output_pdf,
                    file_name=f"pdf_producto_{producto}.pdf",
                    mime="application/pdf"
                )

        except Exception as e:
            st.error(f"Ocurrió un error: {e}")
