import streamlit as st
import pandas as pd
from io import BytesIO
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, BooleanObject

# --- FUNCIÓN PARA RELLENAR PDF (Texto y apariencia) ---
def fill_pdf_text_only(input_pdf, data_dict):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    # Copiar páginas
    for page in reader.pages:
        writer.add_page(page)

    # Copiar AcroForm si existe
    if "/AcroForm" in reader.trailer["/Root"]:
        writer._root_object.update({
            NameObject("/AcroForm"): reader.trailer["/Root"]["/AcroForm"]
        })

    # Forzar NeedAppearances
    writer._root_object["/AcroForm"].update({
        NameObject("/NeedAppearances"): BooleanObject(True)
    })

    # Actualizar campos (rellena y fuerza visualización)
    writer.update_page_form_field_values(writer.pages, data_dict)

    # Guardar PDF en memoria
    output = BytesIO()
    writer.write(output)
    output.seek(0)
    return output

# --- CONFIGURACIÓN STREAMLIT ---
st.set_page_config(page_title="Generador PDF Climagas", page_icon="🧾", layout="centered")
st.title("🧾 Generador de PDF Climagas")

# --- SELECCIÓN DE EMPRESA ---
empresa = st.selectbox(
    "Selecciona la empresa",
    ["Climagas Madrid S.L.", "Instalaciones EcoTerm S.A.", "CalorPlus Energía S.L."]
)

# --- DATOS DE LAS EMPRESAS ---
empresas_datos = {
    "Climagas Madrid S.L.": {
        "Textfield-107": "B87512345",  # NIF
        "Textfield-108": "González",   # Primer Apellido
        "Textfield-109": "Ruiz",       # Segundo Apellido
        "Textfield-110": "Climagas Madrid S.L.",  # Nombre/Razón Social
        "Textfield-111": "info@climagas.es",     # Correo electrónico
        "Textfield-112": "Calle",    # Tipo de vía
        "Textfield-113": "Mayor",    # Nombre vía
        "Textfield-114": "25",       # Nº
        "Textfield-115": "2",        # Bloque
        "Textfield-116": "A",        # Portal
        "Textfield-117": "",         # Escalera
        "Textfield-118": "1",        # Piso
        "Textfield-119": "Dcha",     # Puerta
        "Textfield-120": "Madrid",   # Localidad
        "Textfield-121": "Madrid",   # Provincia
        "Textfield-122": "28001",    # CP
        "Textfield-123": "913000000",# Teléfono Fijo
        "Textfield-124": "600000000" # Teléfono Móvil
    },
    "Instalaciones EcoTerm S.A.": {
        "Textfield-107": "A45896321",
        "Textfield-108": "López",
        "Textfield-109": "Martínez",
        "Textfield-110": "Instalaciones EcoTerm S.A.",
        "Textfield-111": "contacto@ecoterm.com",
        "Textfield-112": "Avenida",
        "Textfield-113": "Andalucía",
        "Textfield-114": "12",
        "Textfield-115": "",
        "Textfield-116": "B",
        "Textfield-117": "2",
        "Textfield-118": "",
        "Textfield-119": "",
        "Textfield-120": "Sevilla",
        "Textfield-121": "Sevilla",
        "Textfield-122": "41003",
        "Textfield-123": "955123456",
        "Textfield-124": "690123456",
    },
    "CalorPlus Energía S.L.": {
        "Textfield-107": "B54321987",
        "Textfield-108": "Santos",
        "Textfield-109": "Pérez",
        "Textfield-110": "CalorPlus Energía S.L.",
        "Textfield-111": "info@calorplus.com",
        "Textfield-112": "Camino",
        "Textfield-113": "Verde",
        "Textfield-114": "8",
        "Textfield-115": "",
        "Textfield-116": "",
        "Textfield-117": "",
        "Textfield-118": "",
        "Textfield-119": "",
        "Textfield-120": "Valencia",
        "Textfield-121": "Valencia",
        "Textfield-122": "46020",
        "Textfield-123": "961234567",
        "Textfield-124": "670234567",
    }
}

# --- SUBIDA DE ARCHIVOS ---
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
                data_dict = {
                    pdf_field: str(row[excel_col])
                    for pdf_field, excel_col in pdf_to_excel_map.items()
                    if excel_col in row
                }

                # Agregar datos de la empresa seleccionada
                data_dict.update(empresas_datos[empresa])

                output_pdf = fill_pdf_text_only(pdf_file, data_dict)

                st.success(f"✅ PDF generado correctamente para el producto {producto} de {empresa}")
                st.download_button(
                    label="⬇️ Descargar PDF con campos rellenados",
                    data=output_pdf,
                    file_name=f"{empresa.replace(' ', '_')}_producto_{producto}.pdf",
                    mime="application/pdf"
                )
        except Exception as e:
            st.error(f"❌ Error al generar el PDF: {e}")
