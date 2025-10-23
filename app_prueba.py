import streamlit as st
import pandas as pd
from io import BytesIO
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, TextStringObject, BooleanObject, ArrayObject, DictionaryObject

def fill_pdf_text_only(input_pdf, data_dict):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    # Copiar páginas
    for page in reader.pages:
        writer.add_page(page)

    # Copiar AcroForm completo si existe
    if "/AcroForm" in reader.trailer["/Root"]:
        writer._root_object.update({
            NameObject("/AcroForm"): reader.trailer["/Root"]["/AcroForm"]
        })
    else:
        writer._root_object.update({
            NameObject("/AcroForm"): DictionaryObject({
                NameObject("/Fields"): ArrayObject(),
                NameObject("/NeedAppearances"): BooleanObject(True)
            })
        })

    writer._root_object["/AcroForm"].update({
        NameObject("/NeedAppearances"): BooleanObject(True)
    })

    # Rellenar campos de texto manualmente (empresa + Excel)
    for page in writer.pages:
        annots = page.get("/Annots")
        if not annots:
            continue
        for annot in annots:
            field = annot.get_object()
            field_name = field.get("/T")
            if not field_name or field_name not in data_dict:
                continue
            if field.get("/FT") == "/Tx":  # solo campos de texto
                field.update({NameObject("/V"): TextStringObject(str(data_dict[field_name]))})

    # Guardar en memoria
    output = BytesIO()
    writer.write(output)
    output.seek(0)
    return output

# --- STREAMLIT ---
st.set_page_config(page_title="Generador PDF Climagas", page_icon="🧾", layout="centered")
st.title("🧾 Generador de PDF Climagas")

# Selección de empresa
empresa = st.selectbox(
    "Selecciona la empresa",
    ["Climagas Madrid S.L.", "Instalaciones EcoTerm S.A.", "CalorPlus Energía S.L."]
)

# Datos de empresa
empresas_datos = {
    "Climagas Madrid S.L.": {
        "Textfield-107": "B87512345", "Textfield-108": "González", "Textfield-109": "Ruiz",
        "Textfield-110": "Climagas Madrid S.L.", "Textfield-111": "info@climagas.es",
        "Textfield-112": "Calle", "Textfield-113": "Mayor", "Textfield-114": "25",
        "Textfield-115": "2", "Textfield-116": "A", "Textfield-117": "", "Textfield-118": "1",
        "Textfield-119": "Dcha", "Textfield-120": "Madrid", "Textfield-121": "Madrid",
        "Textfield-122": "28001", "Textfield-123": "913000000", "Textfield-124": "600000000"
    },
    "Instalaciones EcoTerm S.A.": {
        "Textfield-107": "A45896321", "Textfield-108": "López", "Textfield-109": "Martínez",
        "Textfield-110": "Instalaciones EcoTerm S.A.", "Textfield-111": "contacto@ecoterm.com",
        "Textfield-112": "Avenida", "Textfield-113": "Andalucía", "Textfield-114": "12",
        "Textfield-115": "", "Textfield-116": "B", "Textfield-117": "2", "Textfield-118": "",
        "Textfield-119": "", "Textfield-120": "Sevilla", "Textfield-121": "Sevilla",
        "Textfield-122": "41003", "Textfield-123": "955123456", "Textfield-124": "690123456"
    },
    "CalorPlus Energía S.L.": {
        "Textfield-107": "B54321987", "Textfield-108": "Santos", "Textfield-109": "Pérez",
        "Textfield-110": "CalorPlus Energía S.L.", "Textfield-111": "info@calorplus.com",
        "Textfield-112": "Camino", "Textfield-113": "Verde", "Textfield-114": "8",
        "Textfield-115": "", "Textfield-116": "", "Textfield-117": "", "Textfield-118": "",
        "Textfield-119": "", "Textfield-120": "Valencia", "Textfield-121": "Valencia",
        "Textfield-122": "46020", "Textfield-123": "961234567", "Textfield-124": "670234567"
    }
}

# Subida de archivos
pdf_file = st.file_uploader("📄 Sube la plantilla PDF", type="pdf")
excel_file = st.file_uploader("📊 Sube el Excel con los productos", type=["xls", "xlsx"])
producto = st.number_input("Introduce el código del producto", step=1, min_value=0)

pdf_to_excel_map = {
    "Potencia Nominal kW": "Potencia Nominal kW",
    "Fecha de solicitud de licencia de obra en su defec": "Fecha de solicitud de licencia de obra en su defect",
    "Potencia total inicial": "Potencia total inicial",
    "Potencia a modificar": "Potencia a modificar",
    "Potencia total final": "Potencia total final"
}

# Generar PDF
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

                # Combinar con datos de empresa
                data_dict.update(empresas_datos[empresa])

                output_pdf = fill_pdf_text_only(pdf_file, data_dict)

                st.success(f"✅ PDF generado correctamente para el producto {producto} de {empresa}")
                st.download_button(
                    label="⬇️ Descargar PDF",
                    data=output_pdf,
                    file_name=f"{empresa.replace(' ', '_')}_producto_{producto}.pdf",
                    mime="application/pdf"
                )
        except Exception as e:
            st.error(f"❌ Error al generar el PDF: {e}")

