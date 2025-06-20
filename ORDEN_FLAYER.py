import pandas as pd
import tkinter as tk
import matplotlib.pyplot as plt
import time
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import filedialog, ttk, messagebox

# ---------------------------------------------
# INICIO CODIGO PRINCIPAL
# ---------------------------------------------

# Variable global para almacenar el DataFrame original procesado
df_original = None
materiales_siadal = set()
materiales_encontrados_avanzados = set()
materiales_vistos = set()

# Función principal para cargar y procesar el archivo Excel
def cargar_excel():
    global df_original
    archivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not archivo:
        return

    try:
        progress['value'] = 10
        ventana.update()

        detalle_df = pd.read_excel(archivo, sheet_name="DETALLE FAC", engine="openpyxl")
        progress['value'] = 30
        ventana.update()

        relacion_df = pd.read_excel(archivo, sheet_name="RELACIÓN FAC-PED", engine="openpyxl", header=None)
        relacion_df.columns = relacion_df.iloc[0]
        relacion_df = relacion_df[1:]
        relacion_df = relacion_df[["NumeroFactura", "NumeroPedimento"]]
        progress['value'] = 50
        ventana.update()

        encabezado_df = pd.read_excel(archivo, sheet_name="ENCABEZADO FAC", engine="openpyxl", header=None)
        encabezado_df.columns = ["NumeroFactura", "?", "TipoOperacion"] + list(encabezado_df.columns[3:])
        encabezado_df = encabezado_df[1:]
        encabezado_df = encabezado_df[["NumeroFactura", "TipoOperacion"]]
        progress['value'] = 70
        ventana.update()

        detalle_df["Número Factura"] = detalle_df["Número Factura"].astype(str)
        relacion_df["NumeroFactura"] = relacion_df["NumeroFactura"].astype(str)
        encabezado_df["NumeroFactura"] = encabezado_df["NumeroFactura"].astype(str)

        merged_df = detalle_df.merge(relacion_df, left_on="Número Factura", right_on="NumeroFactura", how="left")
        merged_df = merged_df.merge(encabezado_df, left_on="Número Factura", right_on="NumeroFactura", how="left")

        merged_df["TipoOperacion"] = merged_df["TipoOperacion"].map({
            1: "Importación",
            2: "Exportación"
        })

        df = merged_df[["Número Material", "Número Factura", "Cantidad UMC", "NumeroPedimento", "TipoOperacion"]]
        df = df.dropna(subset=["Número Material"])
        df = df.drop_duplicates(subset=["Número Material"])

        df_original = df
        aplicar_filtro()

        progress['value'] = 100
        ventana.update()
        time.sleep(0.5)
        progress['value'] = 0

    except Exception as e:
        progress['value'] = 0
        messagebox.showerror("Error", f"No se pudo procesar el archivo:\n{e}")

# Aplica el filtro por tipo de operación seleccionado en el combobox
def aplicar_filtro(*args):
    if df_original is None:
        return

    tipo = filtro_operacion.get()
    if tipo == "Todos":
        df_filtrado = df_original
    else:
        # Filtra según el tipo de operación: Importación o Exportación
        df_filtrado = df_original[df_original["TipoOperacion"] == tipo]

    mostrar_datos(df_filtrado)

# Muestra los datos en la tabla visual
def mostrar_datos(df):
    # Limpia la tabla antes de mostrar nuevos datos
    for row in tabla.get_children():
        tabla.delete(row)

    # Inserta cada fila del DataFrame en la tabla
    for _, row in df.iterrows():
        numero_material = str(row["Número Material"]).strip()
        estado = ""
        if materiales_siadal:
            estado = "Existente" if numero_material in materiales_siadal else "No Existe"

        tabla.insert("", "end", values=(
            row["Número Material"],
            row["Número Factura"],
            row["Cantidad UMC"],
            row["NumeroPedimento"],
            row["TipoOperacion"],
            estado
        ))

# Guarda los datos actualmente mostrados en la tabla a un archivo Excel
def guardar_datos():
    # Recupera los valores actuales de la tabla
    filas = [tabla.item(i)['values'] for i in tabla.get_children()]
    columnas = ["Número Material", "Número Factura", "Cantidad UMC", "NumeroPedimento", "TipoOperacion"]
    df_guardar = pd.DataFrame(filas, columns=columnas)

    # Diálogo para guardar el archivo
    archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if archivo:
        df_guardar.to_excel(archivo, index=False)
        messagebox.showinfo("Guardado", "Datos guardados correctamente.")

# ---------------------------------------------
# FIN CODIGO PRINCIPAL
# ---------------------------------------------

# ---------------------------------------------
# INICIO GRAFICOS - ESTADISTICAS
# ---------------------------------------------

def mostrar_estadisticas(materiales_excel, materiales_sql):
    ventana_estadisticas = tk.Toplevel()
    ventana_estadisticas.title("Estadísticas de Materiales")
    ventana_estadisticas.geometry("500x500")

    # Cálculos
    total = len(materiales_excel)
    encontrados = sum(1 for m in materiales_excel if m in materiales_sql)
    no_encontrados = total - encontrados

    porcentaje_encontrados = round((encontrados / total) * 100, 2)
    porcentaje_no_encontrados = round((no_encontrados / total) * 100, 2)

    # Tabla de resultados
    frame_tabla = tk.Frame(ventana_estadisticas)
    frame_tabla.pack(pady=10)

    tabla_resultados = ttk.Treeview(frame_tabla, columns=["Estado", "Cantidad", "Porcentaje"], show="headings", height=3)
    tabla_resultados.heading("Estado", text="Estado")
    tabla_resultados.heading("Cantidad", text="Cantidad")
    tabla_resultados.heading("Porcentaje", text="Porcentaje")

    tabla_resultados.insert("", "end", values=(" Existentes", encontrados, f"{porcentaje_encontrados}%"))
    tabla_resultados.insert("", "end", values=("No encontrados", no_encontrados, f"{porcentaje_no_encontrados}%"))
    tabla_resultados.insert("", "end", values=("Total", total, "100%"))
    tabla_resultados.pack()

    # Gráfico circular
    fig, ax = plt.subplots(figsize=(4, 4))
    ax.pie([encontrados, no_encontrados],
           labels=["Existentes", "No encontrados"],
           colors=["lightgreen", "salmon"],
           autopct='%1.1f%%',
           startangle=90)
    ax.axis('equal')

    canvas = FigureCanvasTkAgg(fig, master=ventana_estadisticas)
    canvas.draw()
    canvas.get_tk_widget().pack(pady=30)

# --------------------------------------------
# FIN GRAFICOS - ESTADISTICAS
# --------------------------------------------

# --------------------------------------------
# INICIO COMPARACION
# --------------------------------------------

def mostrar_comparacion_excel_siadal(materiales_excel, materiales_sql):
    ventana_comparacion = tk.Toplevel()
    ventana_comparacion.title("Comparación Materiales Excel vs SIADAL")
    ventana_comparacion.geometry("600x500")

    # Unión con coincidencias avanzadas
    materiales_existentes_totales = materiales_sql.union(materiales_encontrados_avanzados)

    frame_comparacion = tk.Frame(ventana_comparacion)
    frame_comparacion.pack(fill="both", expand=True, padx=20, pady=20)

    scroll_y = tk.Scrollbar(frame_comparacion, orient="vertical")
    scroll_x = tk.Scrollbar(frame_comparacion, orient="horizontal")

    tabla_comp = ttk.Treeview(
        frame_comparacion,
        columns=["excel", "siadal"],
        show="headings",
        yscrollcommand=scroll_y.set,
        xscrollcommand=scroll_x.set
    )

    tabla_comp.heading("excel", text="Número Material (Excel)")
    tabla_comp.heading("siadal", text="Número Material (SIADAL")

    scroll_y.config(command=tabla_comp.yview)
    scroll_x.config(command=tabla_comp.xview)
    scroll_y.pack(side="right", fill="y")
    scroll_x.pack(side="bottom", fill="x")
    tabla_comp.pack(fill="both", expand=True)

    tabla_comp.tag_configure("sin_coincidencia", background="salmon")

    for mat in sorted(materiales_excel):
        siadal_match = mat if mat in materiales_existentes_totales else ""
        tag = "sin_coincidencia" if siadal_match == "" else ""
        tabla_comp.insert("", "end", values=(mat, siadal_match), tags=(tag,))

# --------------------------------------------
# FIN COMPARACION
# --------------------------------------------

# ---------------------------------------------
# INICIO CONSULTA SQL SERVER Y MUESTRA RESULTADOS
# ---------------------------------------------

def consultar_y_colorear():
    import pyodbc

    if df_original is None:
        messagebox.showwarning("Advertencia", "Primero debes cargar un archivo Excel.")
        return

    try:
        progress['value'] = 10
        ventana.update()

        server = 'PRACTICAS_TI\\MSSQLSERVER1'
        database = 'dbSiadalGoGlobal'
        username = 'sa'
        password = 'root'

        conexion = pyodbc.connect(
            f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        )
        cursor = conexion.cursor()

        cursor.execute("SELECT MatNoParte FROM siadalgoglobaluser.tblMaterial")
        resultados = cursor.fetchall()

        materiales_sql = set(str(fila[0]).strip() for fila in resultados)
        global materiales_siadal
        materiales_siadal = materiales_sql
        progress['value'] = 40
        ventana.update()

        for row in tabla.get_children():
            tabla.delete(row)

        global materiales_vistos
        btn_estadisticas = tk.Button(frame_botones, text="Ver Estadísticas", width=25,
                                     command=lambda: mostrar_estadisticas(materiales_vistos, materiales_siadal))
        btn_estadisticas.pack(side="left", padx=5)

        materiales_vistos = set()
        total = len(df_original)
        procesados = 0

        for _, row in df_original.iterrows():
            numero_material = str(row["Número Material"]).strip()
            if numero_material in materiales_vistos:
                continue
            materiales_vistos.add(numero_material)

            tag_color = "verde" if numero_material in materiales_sql else "rojo"
            estado = "Existente" if tag_color == "verde" else "No existente"

            tabla.insert("", "end", values=(
                numero_material,
                row["Número Factura"],
                row["Cantidad UMC"],
                row["NumeroPedimento"],
                row["TipoOperacion"],
                estado
            ), tags=(tag_color,))

            procesados += 1
            progreso_actual = 40 + int((procesados / total) * 60)
            progress['value'] = progreso_actual
            ventana.update()

        tabla.tag_configure("verde", background="lightgreen")
        tabla.tag_configure("rojo", background="salmon")

        cursor.close()
        conexion.close()

        mostrar_estadisticas(materiales_vistos, materiales_sql)

        # Mostrar tabla comparativa
        mostrar_comparacion_excel_siadal(materiales_vistos, materiales_sql)

        progress['value'] = 100
        ventana.update()
        time.sleep(0.5)
        progress['value'] = 0

    except Exception as e:
        progress['value'] = 0
        messagebox.showerror("Error SQL", f"No se pudo ejecutar la consulta:\n{e}")

# ---------------------------------------------
# FIN CONSULTA SQL SERVER Y MUESTRA RESULTADOS
# ---------------------------------------------

# --------------------------------------------
# INICIO BUSQUEDA AVANZADA
# --------------------------------------------

def buscar_coincidencias_avanzadas():
    import pyodbc

    if df_original is None or not materiales_siadal:
        messagebox.showwarning("Advertencia", "Primero debes cargar y verificar el archivo Excel.")
        return

    try:
        progress['value'] = 10
        ventana.update()

        server = 'PRACTICAS_TI\\MSSQLSERVER1'
        database = 'dbSiadalGoGlobal'
        username = 'sa'
        password = 'root'

        conexion = pyodbc.connect(
            f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        )
        cursor = conexion.cursor()

        df_no_encontrados = df_original[~df_original["Número Material"].astype(str).isin(materiales_siadal)]

        print("---- Coincidencias Avanzadas ----")
        print(f"Total de materiales en Excel: {len(df_original)}")
        print(f"Materiales no encontrados en SIADAL: {len(df_no_encontrados)}")

        if df_no_encontrados.empty:
            messagebox.showinfo("Información", "No hay materiales no encontrados en SIADAL para buscar coincidencias avanzadas.")
            progress['value'] = 0
            return

        resultados_totales = []
        total = len(df_no_encontrados)
        procesados = 0

        for _, fila in df_no_encontrados.iterrows():
            numero_material = str(fila["Número Material"]).strip()
            numero_factura = str(fila["Número Factura"]).strip()
            cantidad_umc = int(fila["Cantidad UMC"])

            query = f"""
                SELECT 
                    c.FactEntFechaEnt AS FECHA, 
                    p.ProvNombre, 
                    c.FactEntPedimento, 
                    c.FactEntNofact AS Factura, 
                    c.FactEntFolio AS Folio, 
                    c.FactEntContenedor AS Caja_CTR, 
                    a.MatNoParte AS NúmeroMaterial, 
                    a.MatDescr AS Descripción, 
                    SUM(b.MovExistente) AS Cantidad
                FROM siadalgoglobaluser.tblMaterial AS a
                INNER JOIN siadalgoglobaluser.tblMovimientos AS b ON a.idMaterial = b.idMaterial
                INNER JOIN siadalgoglobaluser.tblFacturaEnt AS c ON b.idFactEnt = c.idFactEnt
                INNER JOIN siadalgoglobaluser.tblDescrFactEnt AS d 
                    ON a.idMaterial = d.idMaterial 
                    AND c.idFactEnt = d.idFactEnt 
                    AND b.iddescrFERenTras = d.idDescrFactEnt
                INNER JOIN siadalgoglobaluser.tblProveedor AS p ON c.idProveedor = p.idProveedor
                WHERE (c.idAlmacen = 17) 
                  AND (c.FactEntFechaEnt >= 20240101)
                  AND a.MatNoParte LIKE ?
                  AND c.FactEntNofact = ?
                GROUP BY 
                    c.FactEntFechaEnt, p.ProvNombre, c.FactEntPedimento, c.FactEntNofact, 
                    c.FactEntFolio, c.FactEntContenedor, a.MatDescr, a.MatNoParte
                HAVING SUM(b.MovExistente) = ?
                ORDER BY FECHA DESC
            """

            cursor.execute(query, (f"%{numero_material}%", numero_factura, cantidad_umc))
            filas = cursor.fetchall()

            for f in filas:
                resultados_totales.append([
                    f.FECHA, f.ProvNombre, f.FactEntPedimento, f.Factura,
                    f.Folio, f.Caja_CTR, f.NúmeroMaterial, f.Descripción, f.Cantidad
                ])

            procesados += 1
            progreso_actual = 10 + int((procesados / total) * 90)
            progress['value'] = progreso_actual
            ventana.update()

        cursor.close()
        conexion.close()

        progress['value'] = 100
        ventana.update()
        time.sleep(0.3)
        progress['value'] = 0

        ventana_resultados = tk.Toplevel()
        ventana_resultados.title("Coincidencias Avanzadas en SIADAL")
        ventana_resultados.geometry("1100x500")

        frame_tabla = tk.Frame(ventana_resultados)
        frame_tabla.pack(fill="both", expand=True, padx=10, pady=10)

        scroll_y = tk.Scrollbar(frame_tabla, orient="vertical")
        scroll_x = tk.Scrollbar(frame_tabla, orient="horizontal")

        tabla_coincidencias = ttk.Treeview(
            frame_tabla,
            columns=["Fecha", "Proveedor", "Pedimento", "Factura", "Folio", "Caja_CTR", "NúmeroMaterial", "Descripción", "Cantidad"],
            show="headings",
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set,
            height=20
        )

        for col in tabla_coincidencias["columns"]:
            tabla_coincidencias.heading(col, text=col)
            tabla_coincidencias.column(col, width=120, anchor="w")

        scroll_y.config(command=tabla_coincidencias.yview)
        scroll_x.config(command=tabla_coincidencias.xview)
        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        tabla_coincidencias.pack(fill="both", expand=True)

        for fila in resultados_totales:
            tabla_coincidencias.insert("", "end", values=fila)

        # Guardar coincidencias avanzadas encontradas
        global materiales_encontrados_avanzados
        materiales_encontrados_avanzados = set()

        if resultados_totales:
            materiales_encontrados_avanzados = set(f[6] for f in resultados_totales)

            df_exportar = pd.DataFrame(resultados_totales, columns=[
                "Fecha", "Proveedor", "Pedimento", "Factura", "Folio", "Caja_CTR",
                "NúmeroMaterial", "Descripción", "Cantidad"
            ])

            archivo_guardado = filedialog.asksaveasfilename(
                title="Guardar Coincidencias Avanzadas",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )

        if archivo_guardado:
            try:
                df_exportar.to_excel(archivo_guardado, index=False)
                messagebox.showinfo("Archivo Guardado",
                                    f"Coincidencias avanzadas guardadas en:\n{archivo_guardado}")
            except Exception as e:
                messagebox.showerror("Error al guardar", f"No se pudo guardar el archivo:\n{e}")
        else:
            messagebox.showinfo("Sin Resultados",
                            "No se encontraron coincidencias avanzadas para los materiales no existentes.")

    except Exception as e:
        progress['value'] = 0
        messagebox.showerror("Error búsqueda avanzada", f"Error al buscar coincidencias avanzadas:\n{e}")
# --------------------------------------------
# FIN BUSQUEDA AVANZADA
# --------------------------------------------

# ---------------------------------------------
# INICIO INTERFAZ GRAFICA
# ---------------------------------------------

# Crear ventana principal
ventana = tk.Tk()
ventana.title("Analizador de Datos")
ventana.geometry("950x650")

# Botón para cargar el archivo Excel
btn_cargar = tk.Button(ventana, text="Cargar Excel", command=cargar_excel)
btn_cargar.pack(pady=10)

# Sección para el filtro de tipo de operación
frame_filtro = tk.Frame(ventana)
frame_filtro.pack()

# Etiqueta + ComboBox para filtrar tipo de operación
tk.Label(frame_filtro, text="Filtrar por Tipo de Operación:").pack(side="left", padx=5)

filtro_operacion = ttk.Combobox(frame_filtro, values=["Todos", "Importación", "Exportación"])
filtro_operacion.set("Todos")  # Valor por defecto
filtro_operacion.pack(side="left")
filtro_operacion.bind("<<ComboboxSelected>>", aplicar_filtro)  # Aplica filtro al cambiar opción

# Frame que contendrá la tabla y los scrollbars
frame_tabla = tk.Frame(ventana)
frame_tabla.pack(fill="both", expand=True, padx=20, pady=20)

# Scrollbars
scroll_y = tk.Scrollbar(frame_tabla, orient="vertical")
scroll_x = tk.Scrollbar(frame_tabla, orient="horizontal")

# Tabla con scroll
tabla = ttk.Treeview(
    frame_tabla,
    columns=["material", "factura", "umc", "pedimento", "tipo", "estado"],
    show="headings",
    height=20,
    yscrollcommand=scroll_y.set,
    xscrollcommand=scroll_x.set
)

# Configurar columnas
tabla.heading("material", text="Número Material")
tabla.heading("factura", text="Número Factura")
tabla.heading("umc", text="Cantidad UMC")
tabla.heading("pedimento", text="Número Pedimento")
tabla.heading("tipo", text="Tipo Operación")
tabla.heading("estado", text="Estado")

# Asociar scrollbars a la tabla
scroll_y.config(command=tabla.yview)
scroll_x.config(command=tabla.xview)

# Ubicar widgets
scroll_y.pack(side="right", fill="y")
scroll_x.pack(side="bottom", fill="x")
tabla.pack(fill="both", expand=True)

# -------------------------------------------------
# FIN INTERFAZ GRAFICA
# -------------------------------------------------

# Barra de progreso
# Barra de progreso
progress = ttk.Progressbar(ventana, orient="horizontal", length=400, mode="determinate")
progress.pack(pady=10)

# Frame para contener los botones en forma vertical, alineados a la derecha
frame_botones = tk.Frame(ventana)
frame_botones.pack(side="bottom", pady=10)

btn_guardar = tk.Button(frame_botones, text="Guardar resultados", width=25, command=guardar_datos)
btn_guardar.pack(side="left", padx=5)

btn_verificar_sql = tk.Button(frame_botones, text="Verificar Materiales con SQL", width=25, command=consultar_y_colorear)
btn_verificar_sql.pack(side="left", padx=5)

btn_buscar_avanzadas = tk.Button(frame_botones, text="Buscar Coincidencias Avanzadas", width=25, command=buscar_coincidencias_avanzadas)
btn_buscar_avanzadas.pack(side="left", padx=5)

btn_comparar = tk.Button(frame_botones, text="Comparar Excel vs SIADAL", width=25, command=lambda: mostrar_comparacion_excel_siadal(
    [str(x).strip() for x in df_original["Número Material"].dropna().unique()],
    materiales_siadal
))
btn_comparar.pack(side="left", padx=5)

# Inicia el bucle principal de la aplicación
try:
    ventana.mainloop()
except KeyboardInterrupt:
    print("Programa detenido por el usuario.")