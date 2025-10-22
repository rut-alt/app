import streamlit as st
import pandas as pd
from io import BytesIO
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, TextStringObject, BooleanObject, DictionaryObject

st.set_page_config(page_title="Generador PDF Climagas", page_icon="🧾", layout="centered")
st.title("🧾 Generador de PDF Climagas")

st.markdown("Sube tu plantilla PDF y el archivo Excel con los datos del producto.")

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

if st.button("🚀 Generar PDF"):
    if not pdf_file or not excel_file:
        st.error("Por favor, sube el PDF y el Excel antes de continuar.")
    else:
        try:
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

                        # Campo tipo checkbox
                        if field.get("/FT") == "/Btn":
                            field.update({
                                NameObject("/V"): NameObject("/Yes") if value else NameObject("/Off"),
                                NameObject("/AS"): NameObject("/Yes") if value else NameObject("/Off")
                            })
                        else:  # Campo de texto
                            field.update({
                                NameObject("/V"): TextStringObject(str(value)),
                                NameObject("/Ff"): BooleanObject(True)
                            })
                            # Forzar apariencia visual (para que se vea rellenado)
                            field.update({
                                NameObject("/DA"): TextStringObject("/Helv 0 Tf 0 g")
                            })

                # Marcar el formulario como "no editable" y regenerar apariencias
                if "/AcroForm" in writer._root_object:
                    writer._root_object["/AcroForm"].update({
                        NameObject("/NeedAppearances"): BooleanObject(False),
                        NameObject("/SigFlags"): 3
                    })

                output_pdf = BytesIO()
                writer.write(output_pdf)
                output
