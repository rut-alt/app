import streamlit as st
import pandas as pd
from io import BytesIO
from pdfrw import PdfReader, PdfWriter, PdfName, PdfString

# --- FUNCION PARA RELLENAR PDF CON CAMPOS VISIBLES ---
def fill_pdf_pdfrw(input_pdf, data_dict):
    template_pdf = PdfReader(input_pdf)
    
    for page in template_pdf.pages:
        annotations = page.get("/Annots")
        if annotations:
            for annotation in annotations:
                field_name = annotation.get("/T")
                if field_name:
                    # Quitar paréntesis del nombre del campo
                    key = field_name[1:-1] if field_name.startswith("(") and field_name.endswith(")") else field_name
                    if key in data_dict:
                        annotation.update({
                            PdfName("V"): PdfString(str(data_dict[key])),
                            PdfName("Ff"): 1  # forzar visualización
                        })

    output = BytesIO()
    PdfWriter().write(output, template_pdf)
    output.seek(0)
    return output


# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Generador PDF Climagas", page_icon="🧾", layout="centered")
st.title("🧾 Generador de PDF Climagas")

st.markdown("Sube tu plantilla PDF y el archivo Excel con los datos del producto para generar automáticamente un PDF relleno.")

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
                output_pdf = fill_pdf_pdfrw(pdf_file, data_dict)

                st.success(f"✅ PDF generado correctamente para el producto {producto}")

                st.download_button(
                    label="⬇️ Descargar PDF con campos rellenados",
                    data=output_pdf,
                    file_name=f"pdf_producto_{producto}.pdf",
                    mime="application/pdf"
                )

        except Exception as e:
            st.error(f"❌ Error al generar el PDF: {e}")
