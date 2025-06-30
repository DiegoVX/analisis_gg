# excel_loader.py
import pandas as pd
from tkinter import filedialog, messagebox

def cargar_excel():
    archivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not archivo:
        return None

    try:
        df = pd.read_excel(archivo, sheet_name="DETALLE FAC", engine="openpyxl")

        columnas_necesarias = ["Número Material", "Número Factura", "Cantidad UMC"]
        for col in columnas_necesarias:
            if col not in df.columns:
                messagebox.showerror("Error", f"Falta la columna requerida: {col}")
                return None

        return df

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo cargar el archivo:\n{e}")
        return None