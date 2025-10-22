import tkinter as tk
from tkinter import messagebox
import pandas as pd
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, TextStringObject, BooleanObject

# Archivos
input_pdf = "IT315 PLANTILLA CLIMAGAS ACT.pdf"
output_pdf = "pdf_campos_rellenados.pdf"
excel_file = "datos_productos.xlsx"

# Mapeo PDF -> Excel
pdf_to_excel_map = {
    "Potencia Nominal kW": "Potencia Nominal kW",
    "Fecha de solicitud de licencia de obra en su defec": "Fecha de solicitud de licencia de obra en su defect",
    "Potencia total inicial": "Potencia total inicial",
    "Potencia a modificar": "Potencia a modificar",
    "Potencia total final": "Potencia total final",
    "Administrativo": "Administrativo"
}

def generar_pdf():
    try:
        # Leer producto
        producto = int(entry_producto.get())
        df = pd.read_excel(excel_file)

        reader = PdfReader(input_pdf)
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
                    # Forzar actualización de apariencia
                    field.update({NameObject("/Ff"): BooleanObject(True)})

        # Guardar PDF
        with open(output_pdf, "wb") as f:
            writer.write(f)

        messagebox.showinfo("¡Listo!", f"PDF generado para el producto {producto}")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# GUI
root = tk.Tk()
root.title("Generador de PDF Climagas")
root.geometry("350x150")

tk.Label(root, text="Introduce el código del producto:").pack(padx=10, pady=10)
entry_producto = tk.Entry(root)
entry_producto.pack(padx=10, pady=5)

tk.Button(root, text="Generar PDF", command=generar_pdf).pack(pady=20)

root.mainloop()