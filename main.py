# main.py
import tkinter as tk
import re
from tkinter import ttk, messagebox, filedialog
#from excel_loader import cargar_excel
from sql_checker import buscar_coincidencia_siadal
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import ttk as ttk_native

# Variables globales
df_detalle = None

# Interfaz gráfica
ventana = tk.Tk()
ventana.title("Validación de Materiales con SIADAL")
ventana.geometry("1000x600")

# Barra de progreso global (debe ir después de crear ventana)
progress_var = tk.DoubleVar()
progress_bar = ttk_native.Progressbar(ventana, variable=progress_var, maximum=100, length=400)
progress_label = tk.Label(ventana, text="")

# Función de normalización avanzada para comparar materiales
def normalizar_material(mat):
    if not isinstance(mat, str):
        mat = str(mat)
    mat = mat.strip().upper()
    mat = re.sub(r'[\s\-.\/]', '', mat)  # Quita espacios, guiones, puntos, diagonales
    mat = mat.lstrip('0')
    return mat

# Mostrar datos en tabla con colores y coincidencias
def mostrar_materiales(df):
    for row in tabla.get_children():
        tabla.delete(row)

    for idx, row in df.iterrows():
        numero_material = str(row["Número Material"]).strip()
        numero_factura = str(row["Número Factura"]).strip()
        cantidad_umc = row["Cantidad UMC"]

        resultado = buscar_coincidencia_siadal(numero_material, numero_factura, cantidad_umc)
        match = resultado["match"]
        material_siadal = resultado.get("matno", "")

        if match == "exacto":
            color_tag = "verde"
            estado = "Coincidencia Exacta"
        elif match == "parcial":
            color_tag = "amarillo"
            estado = "Coincidencia Parcial"
        else:
            color_tag = "rojo"
            estado = "Sin Coincidencia"
            if not material_siadal:
                np_siadal = row.get("NP SIADAL", "")
                # Si NP SIADAL es válido, usarlo
                if not pd.isna(np_siadal) and str(np_siadal).strip().lower() != "nan":
                    material_siadal = str(np_siadal).strip()
                else:
                    # 1. Buscar coincidencia por número de parte y cantidad (no factura)
                    filtro_np_cant = df[
                        (df["Número Material"].astype(str).str.strip() == numero_material) &
                        (df["Cantidad UMC"] == cantidad_umc) &
                        (df["Número Factura"].astype(str).str.strip() != numero_factura)
                    ]
                    if not filtro_np_cant.empty:
                        fila = filtro_np_cant.iloc[0]
                        almacen = fila.get("Almacen", "")
                        if almacen and not pd.isna(almacen):
                            material_siadal = f"{fila['Número Material']} (Almacén: {almacen})"
                        else:
                            material_siadal = str(fila['Número Material'])
                    else:
                        # 2. Buscar coincidencia por factura y cantidad (no número de parte)
                        filtro_fac_cant = df[
                            (df["Número Factura"].astype(str).str.strip() == numero_factura) &
                            (df["Cantidad UMC"] == cantidad_umc) &
                            (df["Número Material"].astype(str).str.strip() != numero_material)
                        ]
                        if not filtro_fac_cant.empty:
                            fila = filtro_fac_cant.iloc[0]
                            almacen = fila.get("Almacen", "")
                            if almacen and not pd.isna(almacen):
                                material_siadal = f"{fila['Número Material']} (Almacén: {almacen})"
                            else:
                                material_siadal = str(fila['Número Material'])
                        else:
                            # 3. Buscar coincidencia solo por factura (ya lo tienes)
                            filtro = df[df["Número Factura"].astype(str).str.strip() == numero_factura]
                            if not filtro.empty:
                                for _, fila in filtro.iterrows():
                                    mat_encontrado = str(fila["Número Material"]).strip()
                                    if mat_encontrado != numero_material:
                                        almacen = fila.get("Almacen", "")
                                        if almacen and not pd.isna(almacen):
                                            material_siadal = f"{mat_encontrado} (Almacén: {almacen})"
                                        else:
                                            material_siadal = mat_encontrado
                                        break
                            else:
                                material_siadal = ""

        tabla.insert("", "end", values=(idx + 1, numero_material, numero_factura, cantidad_umc, material_siadal, estado), tags=(color_tag,))

    tabla.tag_configure("verde", background="lightgreen")
    tabla.tag_configure("amarillo", background="khaki")
    tabla.tag_configure("rojo", background="salmon")

# Evento para cargar Excel con barra de progreso

def evento_cargar():
    global df_detalle
    archivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not archivo:
        return
    progress_bar.pack(pady=(20, 0))
    progress_label.pack()
    ventana.update()
    try:
        # Intentar leer la hoja 'DETALLE FAC', si no existe, leer la primera hoja disponible
        try:
            df_tmp = pd.read_excel(archivo, sheet_name="DETALLE FAC", engine="openpyxl")
        except Exception:
            # Si falla, leer la primera hoja
            xl = pd.ExcelFile(archivo, engine="openpyxl")
            primera_hoja = xl.sheet_names[0]
            df_tmp = xl.parse(primera_hoja)
        total = len(df_tmp)
        if total == 0:
            raise Exception("El archivo no contiene datos.")
        df_detalle = pd.DataFrame()
        for i, row in df_tmp.iterrows():
            df_detalle = pd.concat([df_detalle, pd.DataFrame([row])], ignore_index=True)
            if i % max(1, total // 100) == 0 or i == total - 1:
                percent = (i + 1) / total * 100
                progress_var.set(percent)
                progress_label.config(text=f"Cargando: {percent:.1f}%")
                ventana.update()
        mostrar_materiales(df_detalle)
        progress_label.config(text="Carga completada.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo cargar el archivo:\n{e}")
        progress_label.config(text="Error al cargar.")
    finally:
        ventana.after(1200, lambda: (progress_bar.pack_forget(), progress_label.pack_forget()))

# Sidebar para los botones (estilo Bootstrap 5)
sidebar = tk.Frame(ventana, width=200, bg="#f8f9fa")
sidebar.pack(side="left", fill="y")

estilo_btn = {
    'font': ("Segoe UI", 11, "bold"),
    'bd': 0,
    'relief': 'flat',
    'activebackground': '#e2e6ea',
    'cursor': 'hand2',
    'height': 2,
    'width': 18,
    'highlightthickness': 0,
}

btn_cargar = tk.Button(sidebar, text="Cargar Excel", command=evento_cargar, bg="#0d6efd", fg="white", activeforeground="white", **estilo_btn)
btn_cargar.pack(pady=(30,10), padx=10)

# --- Botón y función para exportar el análisis a Excel ---
def exportar_excel():
    # Exporta los datos que se muestran en la tabla, no el DataFrame original
    columnas_tabla = [tabla.heading(col)["text"] for col in tabla["columns"]]
    datos_tabla = []
    for item in tabla.get_children():
        fila = tabla.item(item)["values"]
        datos_tabla.append(fila)
    if not datos_tabla:
        messagebox.showwarning("Advertencia", "No hay datos para exportar.")
        return
    df_export = pd.DataFrame(datos_tabla, columns=columnas_tabla)
    # Solo reemplaza vacíos/nulos en la columna I (índice 8) por 'REVISAR MANUAL' si existe
    if len(df_export.columns) > 8:
        col_i = df_export.columns[8]
        df_export[col_i] = df_export[col_i].fillna("REVISAR MANUAL")
        df_export[col_i] = df_export[col_i].replace(r'^\s*$', 'REVISAR MANUAL', regex=True)
    archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not archivo:
        return
    try:
        df_export.to_excel(archivo, index=False, engine="openpyxl")
        messagebox.showinfo("Éxito", f"Archivo exportado correctamente:\n{archivo}")
    except PermissionError:
        messagebox.showerror("Error", "El archivo está abierto o protegido. Ciérralo e inténtalo de nuevo.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo exportar el archivo:\n{e}")

btn_exportar = tk.Button(sidebar, text="Exportar Excel", command=exportar_excel, bg="#198754", fg="white", activeforeground="white", **estilo_btn)
btn_exportar.pack(pady=10, padx=10)

# --- Botón y función para mostrar estadísticos ---
def mostrar_estadisticos():
    estados = [tabla.item(item)["values"][5] for item in tabla.get_children()]
    total = len(estados)
    exacta = estados.count("Coincidencia Exacta")
    parcial = estados.count("Coincidencia Parcial")
    sin = estados.count("Sin Coincidencia")
    win = tk.Toplevel(ventana)
    win.title("Estadísticos de Coincidencias")
    win.geometry("350x220")
    win.configure(bg="#f8f9fa")
    tk.Label(win, text=f"Total de materiales: {total}", font=("Segoe UI", 12, "bold"), bg="#f8f9fa", fg="#212529").pack(pady=(20,10))
    tk.Label(win, text=f"Coincidencia Exacta: {exacta}", font=("Segoe UI", 12), bg="#f8f9fa", fg="#198754").pack(pady=5)
    tk.Label(win, text=f"Coincidencia Parcial: {parcial}", font=("Segoe UI", 12), bg="#f8f9fa", fg="#ffc107").pack(pady=5)
    tk.Label(win, text=f"Sin Coincidencia: {sin}", font=("Segoe UI", 12), bg="#f8f9fa", fg="#dc3545").pack(pady=5)
    tk.Button(win, text="Cerrar", command=win.destroy, bg="#0d6efd", fg="white", font=("Segoe UI", 11, "bold"), bd=0, relief='flat', activebackground="#0b5ed7", cursor='hand2', height=1, width=12).pack(pady=15)

btn_estadisticos = tk.Button(sidebar, text="Ver Estadísticos", command=mostrar_estadisticos, bg="#ffc107", fg="#212529", activeforeground="#212529", **estilo_btn)
btn_estadisticos.pack(pady=10, padx=10)

btn_grafico = tk.Button(sidebar, text="Gráfico de Pastel", command=lambda: mostrar_grafico_pastel(), bg="#dc3545", fg="white", activeforeground="white", **estilo_btn)
btn_grafico.pack(pady=10, padx=10)

# --- Función para mostrar gráfico de pastel ---
def mostrar_grafico_pastel():
    estados = [tabla.item(item)["values"][5] for item in tabla.get_children()]
    labels = ["Coincidencia Exacta", "Coincidencia Parcial", "Sin Coincidencia"]
    sizes = [estados.count(l) for l in labels]
    colors = ["#198754", "#ffc107", "#dc3545"]
    if sum(sizes) == 0:
        messagebox.showinfo("Sin datos", "No hay datos para graficar.")
        return
    win = tk.Toplevel(ventana)
    win.title("Gráfico de Coincidencias")
    win.geometry("420x420")
    fig, ax = plt.subplots(figsize=(4,4))
    wedges, texts, autotexts = ax.pie(sizes, labels=labels, autopct="%1.1f%%", colors=colors, startangle=90, textprops={'fontsize': 11, 'fontname': 'Segoe UI'})
    ax.axis("equal")
    plt.tight_layout()
    canvas = FigureCanvasTkAgg(fig, master=win)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True)
    tk.Button(win, text="Cerrar", command=win.destroy, bg="#0d6efd", fg="white", font=("Segoe UI", 11, "bold"), bd=0, relief='flat', activebackground="#0b5ed7", cursor='hand2', height=1, width=12).pack(pady=10)

frame_tabla = tk.Frame(ventana)
frame_tabla.pack(fill="both", expand=True, padx=(0,20), pady=20)

scroll_y = tk.Scrollbar(frame_tabla, orient="vertical")
scroll_x = tk.Scrollbar(frame_tabla, orient="horizontal")

tabla = ttk.Treeview(
    frame_tabla,
    columns=["n", "material", "factura", "cantidad", "material_siadal", "estado"],
    show="headings",
    height=20,
    yscrollcommand=scroll_y.set,
    xscrollcommand=scroll_x.set
)

headers = [
    ("n", "N°"),
    ("material", "Número Material"),
    ("factura", "Número Factura"),
    ("cantidad", "Cantidad UMC"),
    ("material_siadal", "Material SIADAL"),
    ("estado", "Estado")
]

for col, txt in headers:
    tabla.heading(col, text=txt)

scroll_y.config(command=tabla.yview)
scroll_x.config(command=tabla.xview)
scroll_y.pack(side="right", fill="y")
scroll_x.pack(side="bottom", fill="x")
tabla.pack(fill="both", expand=True)

btn_salir = tk.Button(sidebar, text="Salir", command=ventana.destroy, bg="#6c757d", fg="white", activeforeground="white", **estilo_btn)
btn_salir.pack(pady=(40,10), padx=10)

ventana.mainloop()