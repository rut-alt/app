import streamlit as st
import pandas as pd
from io import BytesIO
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, TextStringObject, BooleanObject

# --- FUNCIÓN PARA RELLENAR PDF (solo texto) ---
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

    # Rellenar los campos del PDF
    for page in writer.pages:
        annots = page.get("/Annots")
        if not annots:
            continue
        for annot in annots:
            field = annot.get_object()
            field_name = field.get("/T")
            if not field_name or field_name not in data_dict:
                continue

            if field.get("/FT") == "/Tx":  # Solo texto
                field.update({NameObject("/V"): TextStringObject(str(data_dict[field_name]))})

    # Guardar PDF en memoria
    output = BytesIO()
    writer.write(output)
    output.seek(0)
    return output

# --- CONFIGURACIÓN DE LA APP ---
st.set_page_config(page_title="Generador PDF Climagas", page_icon="🧾", layout="centered")
st.title("🧾 Generador de PDF Climagas")

# --- SELECCIÓN DE EMPRESA ---
st.subheader("Selecciona la empresa")
empresa = st.selectbox(
    "Nombre de la empresa",
    ["Climagas Madrid S.L.", "Instalaciones EcoTerm S.A.", "CalorPlus Energía S.L."]
)

# --- DATOS POR EMPRESA ---
empresas_datos = {
    "Climagas Madrid S.L.": {
        "textfield-107": "B87512345",  # NIF
        "textfield-108": "González",
        "textfield-109": "Ruiz",
        "textfield-110": "Climagas Madrid S.L.",
        "textfield-111": "info@climagas.es",
        "textfield-112": "Calle",
        "textfield-113": "Mayor",
        "textfield-114": "25",
        "textfield-115": "2",
        "textfield-116": "A",
        "textfield-117": "",
        "textfield-118": "1",
        "textfield-119": "Dcha",
        "textfield-120": "Madrid",
        "textfield-121": "Madrid",
        "textfield-122": "28001",
        "textfield-123": "913000000",
        "textfield-124": "600000000",
    },
    "Instalaciones EcoTerm S.A.": {
        "textfield-107": "A45896321",
        "textfield-108": "López",
        "textfield-109": "Martínez",
        "textfield-110": "Instalaciones EcoTerm S.A.",
        "textfield-111": "contacto@ecoterm.com",
        "textfield-112": "Avenida",
        "textfield-113": "Andalucía",
        "textfield-114": "12",
        "textfield-115": "",
        "textfield-116": "B",
        "textfield-117": "2",
        "textfield-118": "",
        "textfield-119": "",
        "textfield-120": "Sevilla",
        "textfield-121": "Sevilla",
        "textfield-122": "41003",
        "textfield-123": "955123456",
        "textfield-124": "690123456",
    },
    "CalorPlus Energía S.L.": {
        "textfield-107": "B54321987",
        "textfield-108": "Santos",
        "textfield-109": "Pérez",
        "textfield-110": "CalorPlus Energía S.L.",
        "textfield-111": "info@calorplus.com",
        "textfield-112": "Camino",
        "textfield-113": "Verde",
        "textfield-114": "8",
        "textfield-115": "",
        "textfield-116": "",
        "textfield-117": "",
        "textfield-118": "",
        "textfield-119": "",
        "textfield-120": "Valencia",
        "textfield-121": "Valencia",
        "textfield-122": "46020",
        "textfield-123": "961234567",
        "textfield-124": "670234567",
    }
}

# --- SUBIDA DE ARCHIVOS ---
pdf_file = st.file_uploader("📄 Sube la plantilla PDF", type="pdf")
excel_file = st.file_uploader("📊 Sube el Excel con los productos", type=["xls", "xlsx"])
producto = st.number_input("Introduce el código del producto", step=1, min_value=0)

# --- MAPEOS DE CAMPOS PDF → EXCEL ---
pdf_to_excel_map = {
    "Potencia Nominal kW": "Potencia Nominal kW",
    "Fecha de solicitud de licencia de obra en su defec": "Fecha de solicitud de licencia de obra en su defect",
    "Potencia total inicial": "Potencia total inicial",
    "Potencia a modificar": "Potencia a modificar",
    "Potencia total final": "Potencia total final"
}

# --- BOTÓN GENERAR ---
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

                # Añadir los datos de la empresa seleccionada
                data_dict.update(empresas_datos[empresa])

                output_pdf = fill_pdf_text_only(pdf_file, data_dict)

                st.success(f"✅ PDF generado correctamente para el producto {producto} de {empresa}")
                st.download_button(
                    label="⬇️ Descargar PDF con campos de texto rellenados",
                    data=output_pdf,
                    file_name=f"{empresa.replace(' ', '_')}_producto_{producto}.pdf",
                    mime="application/pdf"
                )
        except Exception as e:
            st.error(f"❌ Error al generar el PDF: {e}")
