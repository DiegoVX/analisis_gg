import pandas as pd
import tkinter as tk
import matplotlib.pyplot as plt
import time
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import filedialog, ttk, messagebox

# INICIO CODIGO PRINCIPAL
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

        print("Filas reales del detalle:", len(detalle_df)) # Debug

        relacion_df = pd.read_excel(archivo, sheet_name="RELACIÓN FAC-PED", engine="openpyxl", header=None)
        relacion_df.columns = relacion_df.iloc[0]
        relacion_df = relacion_df[1:]
        relacion_df = relacion_df[["NumeroFactura", "NumeroPedimento"]]
        relacion_df = relacion_df.drop_duplicates(subset=["NumeroFactura"])
        progress['value'] = 50
        ventana.update()

        encabezado_df = pd.read_excel(archivo, sheet_name="ENCABEZADO FAC", engine="openpyxl", header=None)
        encabezado_df.columns = ["NumeroFactura", "?", "TipoOperacion"] + list(encabezado_df.columns[3:])
        encabezado_df = encabezado_df[1:]
        encabezado_df = encabezado_df[["NumeroFactura", "TipoOperacion"]]
        encabezado_df = encabezado_df.drop_duplicates(subset=["NumeroFactura"])
        progress['value'] = 70
        ventana.update()

        # Unificar tipos
        detalle_df["Número Factura"] = detalle_df["Número Factura"].astype(str)
        relacion_df["NumeroFactura"] = relacion_df["NumeroFactura"].astype(str)
        encabezado_df["NumeroFactura"] = encabezado_df["NumeroFactura"].astype(str)

        # Realizar los merge sin multiplicación de filas
        merged_df = detalle_df.merge(relacion_df, left_on="Número Factura", right_on="NumeroFactura", how="left")
        merged_df = merged_df.merge(encabezado_df, left_on="Número Factura", right_on="NumeroFactura", how="left")
        merged_df["TipoOperacion"] = merged_df["TipoOperacion"].map({
            1: "Importación",
            2: "Exportación"
        })

        df = merged_df[["Número Material", "Número Factura", "Cantidad UMC", "NumeroPedimento", "TipoOperacion"]]

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
    for idx, (_,row) in enumerate(df.iterrows(), start = 1):
        numero_material = str(row["Número Material"]).strip() if pd.notna(row["Número Material"]) else ""
        estado = ""
        if materiales_siadal:
            estado = "Existente" if numero_material in materiales_siadal else "No Existe"

        tabla.insert("", "end", values=(
            idx,
            numero_material,
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
    columnas = ["N°", "Número Material", "Número Factura", "Cantidad UMC", "NumeroPedimento", "TipoOperacion"]
    filas = [tabla.item(i)['values'] for i in tabla.get_children()]
    df_guardar = pd.DataFrame(filas, columns=columnas)

    # Diálogo para guardar el archivo
    archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if archivo:
        df_guardar.to_excel(archivo, index=False)
        messagebox.showinfo("Guardado", "Datos guardados correctamente.")

# INICIO GRAFICOS - ESTADISTICAS
def mostrar_estadisticas():
    if df_original is None:
        return

    ventana_estadisticas = tk.Toplevel()
    ventana_estadisticas.title("Estadísticas de Materiales")
    ventana_estadisticas.geometry("500x500")

    # Obtener el filtro actual
    tipo = filtro_operacion.get()

    # Filtrar datos según selección
    if tipo == "Todos":
        df_filtrado = df_original
    else:
        df_filtrado = df_original[df_original["TipoOperacion"] == tipo]

    # Calcular estadísticas sobre TODOS los materiales (incluyendo duplicados)
    total_registros = len(df_filtrado)
    materiales_excel = df_filtrado["Número Material"].astype(str).str.strip()
    encontrados = sum(1 for m in materiales_excel if m in materiales_siadal)
    no_encontrados = total_registros - encontrados

    porcentaje_encontrados = round((encontrados / total_registros) * 100, 2) if total_registros > 0 else 0
    porcentaje_no_encontrados = round((no_encontrados / total_registros) * 100, 2) if total_registros > 0 else 0

    # Tabla de resultados
    frame_tabla = tk.Frame(ventana_estadisticas)
    frame_tabla.pack(pady=10)

    tabla_resultados = ttk.Treeview(frame_tabla, columns=["Estado", "Cantidad", "Porcentaje"], show="headings",
                                    height=3)
    tabla_resultados.heading("Estado", text="Estado")
    tabla_resultados.heading("Cantidad", text="Cantidad")
    tabla_resultados.heading("Porcentaje", text="Porcentaje")

    tabla_resultados.insert("", "end", values=("Existentes", encontrados, f"{porcentaje_encontrados}%"))
    tabla_resultados.insert("", "end", values=("No encontrados", no_encontrados, f"{porcentaje_no_encontrados}%"))
    tabla_resultados.insert("", "end", values=("Total registros", total_registros, "100%"))
    tabla_resultados.pack()

    # Gráfico circular
    fig, ax = plt.subplots(figsize=(4, 4))
    if total_registros > 0:
        ax.pie([encontrados, no_encontrados],
               labels=["Existentes", "No encontrados"],
               colors=["lightgreen", "salmon"],
               autopct='%1.1f%%',
               startangle=90)
        ax.axis('equal')
    else:
        ax.text(0.5, 0.5, "No hay datos para mostrar", ha='center', va='center')

    canvas = FigureCanvasTkAgg(fig, master=ventana_estadisticas)
    canvas.draw()
    canvas.get_tk_widget().pack(pady=30)

# INICIO COMPARACION
def mostrar_comparacion_excel_siadal():
    if df_original is None:
        return

    ventana_comparacion = tk.Toplevel()
    ventana_comparacion.title("Comparación Completa Excel vs SIADAL")
    ventana_comparacion.geometry("900x650")  # Ventana más grande para más columnas

    # Obtener el filtro actual
    tipo = filtro_operacion.get()

    # Filtrar datos según selección
    if tipo == "Todos":
        df_filtrado = df_original.copy()
        titulo = f"Comparación Completa (Todos los {len(df_filtrado)} registros)"
    else:
        df_filtrado = df_original[df_original["TipoOperacion"] == tipo].copy()
        titulo = f"Comparación ({len(df_filtrado)} registros de {tipo})"

    ventana_comparacion.title(titulo)

    frame_comparacion = tk.Frame(ventana_comparacion)
    frame_comparacion.pack(fill="both", expand=True, padx=10, pady=10)

    scroll_y = tk.Scrollbar(frame_comparacion, orient="vertical")
    scroll_x = tk.Scrollbar(frame_comparacion, orient="horizontal")

    # Columnas adicionales para mostrar toda la información
    tabla_comp = ttk.Treeview(
        frame_comparacion,
        columns=["n", "material", "factura", "cantidad", "pedimento", "siadal", "tipo", "estado"],
        show="headings",
        yscrollcommand=scroll_y.set,
        xscrollcommand=scroll_x.set,
        height=25
    )

    # Configurar columnas
    columnas = [
        ("n", "N°", 50),
        ("material", "Material (Excel)", 150),
        ("factura", "N° Factura", 100),
        ("cantidad", "Cantidad", 80),
        ("pedimento", "Pedimento", 100),
        ("siadal", "Material (SIADAL)", 150),
        ("tipo", "Tipo Operación", 100),
        ("estado", "Estado", 120)
    ]

    for col, text, width in columnas:
        tabla_comp.heading(col, text=text)
        tabla_comp.column(col, width=width, anchor="w")

    scroll_y.config(command=tabla_comp.yview)
    scroll_x.config(command=tabla_comp.xview)
    scroll_y.pack(side="right", fill="y")
    scroll_x.pack(side="bottom", fill="x")
    tabla_comp.pack(fill="both", expand=True)

    # Configurar estilos
    tabla_comp.tag_configure("sin_coincidencia", background="#FFCCCB")  # Rojo claro
    tabla_comp.tag_configure("sql_match", background="#90EE90")  # Verde claro
    tabla_comp.tag_configure("avanzado_match", background="#ADD8E6")  # Azul claro

    # Procesar TODOS los registros del DataFrame filtrado (no solo únicos)
    for idx, (_, row) in enumerate(df_filtrado.iterrows(), start=1):
        material_excel = str(row["Número Material"]).strip() if pd.notna(row["Número Material"]) else ""
        factura = str(row["Número Factura"])
        cantidad = str(row["Cantidad UMC"])
        pedimento = str(row["NumeroPedimento"])
        tipo_op = row["TipoOperacion"]

        # Determinar coincidencia
        if material_excel in materiales_siadal:
            material_siadal = material_excel
            estado = "Existente (SQL)"
            tag = "sql_match"
        elif material_excel in materiales_encontrados_avanzados:
            material_siadal = material_excel
            estado = "Existente (Avanzado)"
            tag = "avanzado_match"
        else:
            material_siadal = ""
            estado = "No encontrado"
            tag = "sin_coincidencia"

        tabla_comp.insert("", "end", values=(
            idx,
            material_excel,
            factura,
            cantidad,
            pedimento,
            material_siadal,
            tipo_op,
            estado
        ), tags=(tag,))

    # Añadir contador de registros
    frame_contador = tk.Frame(ventana_comparacion)
    frame_contador.pack(pady=5)
    tk.Label(frame_contador, text=f"Total de registros mostrados: {len(df_filtrado)}",
             font=('Arial', 10, 'bold')).pack()

    # Botón para exportar a Excel
    btn_exportar = tk.Button(ventana_comparacion, text="Exportar a Excel",
                             command=lambda: exportar_comparacion(df_filtrado))
    btn_exportar.pack(pady=5)


def exportar_comparacion(df):
    archivo = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Guardar comparación como"
    )
    if archivo:
        try:
            df.to_excel(archivo, index=False)
            messagebox.showinfo("Éxito", "Comparación exportada correctamente")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar: {str(e)}")

# INICIO CONSULTA SQL SERVER Y MUESTRA RESULTADOS
def consultar_y_colorear():
    import pyodbc
    if df_original is None:
        messagebox.showwarning("Advertencia", "Primero debes cargar un archivo Excel.")
        return
    try:
        progress['value'] = 10
        ventana.update()

        # Conexión SQL (mantener igual)
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

        # Limpiar tabla
        for row in tabla.get_children():
            tabla.delete(row)

        # MODIFICACIÓN CLAVE: Mostrar TODOS los registros (4,722) incluyendo duplicados
        total = len(df_original)
        procesados = 0

        # Obtener el tipo de operación seleccionado
        tipo_operacion = filtro_operacion.get()

        # Filtrar según el combobox
        if tipo_operacion == "Todos":
            df_a_mostrar = df_original.copy()
        else:
            df_a_mostrar = df_original[df_original["TipoOperacion"] == tipo_operacion]

        # Mostrar todos los registros del filtro aplicado
        for idx, row in df_a_mostrar.iterrows():
            numero_material = str(row["Número Material"]).strip() if pd.notna(row["Número Material"]) else ""

            tag_color = "verde" if numero_material in materiales_sql else "rojo"
            estado = "Existente" if tag_color == "verde" else "No existente"

            tabla.insert("", "end", values=(
                idx + 1,  # Número consecutivo
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

        # Configurar colores
        tabla.tag_configure("verde", background="lightgreen")
        tabla.tag_configure("rojo", background="salmon")

        cursor.close()
        conexion.close()

        # Actualizar materiales_vistos con TODOS los materiales únicos (sin duplicados)
        global materiales_vistos
        materiales_vistos = set(str(x).strip() for x in df_original["Número Material"].dropna().unique())

        # Mostrar estadísticas con los materiales únicos
        mostrar_estadisticas(materiales_vistos, materiales_siadal)
        mostrar_comparacion_excel_siadal(materiales_vistos, materiales_siadal)

        progress['value'] = 100
        ventana.update()
        time.sleep(0.5)
        progress['value'] = 0

    except Exception as e:
        progress['value'] = 0
        messagebox.showerror("Error SQL", f"No se pudo ejecutar la consulta:\n{e}")

# INICIO BUSQUEDA AVANZADA
def buscar_coincidencias_avanzadas():
    import pyodbc

    if df_original is None or not materiales_siadal:
        messagebox.showwarning("Advertencia", "Primero debes cargar y verificar el archivo Excel.")
        return

    try:
        progress['value'] = 10
        ventana.update()

        # Configuración de conexión
        server = 'PRACTICAS_TI\\MSSQLSERVER1'
        database = 'dbSiadalGoGlobal'
        username = 'sa'
        password = 'root'

        conexion = pyodbc.connect(
            f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        )
        cursor = conexion.cursor()

        # Obtener el filtro actual
        tipo = filtro_operacion.get()

        # Filtrar según el combobox pero mantener todos los registros para evaluación
        if tipo == "Todos":
            df_a_evaluar = df_original.copy()
        else:
            df_a_evaluar = df_original[df_original["TipoOperacion"] == tipo].copy()

        # Identificar materiales no encontrados (considerando todos los registros)
        df_no_encontrados = df_a_evaluar[
            ~df_a_evaluar["Número Material"].astype(str).str.strip().isin(materiales_siadal)
        ]

        if df_no_encontrados.empty:
            messagebox.showinfo("Información",
                                "No hay materiales no encontrados en SIADAL para buscar coincidencias avanzadas.")
            progress['value'] = 0
            return

        resultados_totales = []
        total = len(df_no_encontrados)
        procesados = 0

        # Buscar coincidencias avanzadas para cada registro no encontrado
        for _, fila in df_no_encontrados.iterrows():
            numero_material = str(fila["Número Material"]).strip() if pd.notna(fila["Número Material"]) else ""
            numero_factura = str(fila["Número Factura"]).strip()
            cantidad_umc = float(fila["Cantidad UMC"])  # Convertir a float para evitar errores

            # Consulta SQL optimizada para obtener TODAS las coincidencias posibles
            query = """
            SELECT 
                c.FactEntFechaEnt AS FECHA, 
                p.ProvNombre, 
                c.FactEntPedimento, 
                c.FactEntNofact AS Factura, 
                c.FactEntFolio AS Folio, 
                c.FactEntContenedor AS Caja_CTR, 
                a.MatNoParte AS NúmeroMaterial, 
                a.MatDescr AS Descripción, 
                b.MovExistente AS Cantidad
            FROM siadalgoglobaluser.tblMaterial AS a
            INNER JOIN siadalgoglobaluser.tblMovimientos AS b ON a.idMaterial = b.idMaterial
            INNER JOIN siadalgoglobaluser.tblFacturaEnt AS c ON b.idFactEnt = c.idFactEnt
            INNER JOIN siadalgoglobaluser.tblDescrFactEnt AS d 
                ON a.idMaterial = d.idMaterial 
                AND c.idFactEnt = d.idFactEnt 
                AND b.iddescrFERenTras = d.idDescrFactEnt
            INNER JOIN siadalgoglobaluser.tblProveedor AS p ON c.idProveedor = p.idProveedor
            WHERE c.idAlmacen = 17 
              AND c.FactEntFechaEnt >= ?
              AND (
                  a.MatNoParte LIKE ? OR  -- Coincidencia parcial
                  a.MatNoParte LIKE ? OR  -- Coincidencia al final
                  a.MatNoParte LIKE ?     -- Coincidencia al inicio
              )
              AND c.FactEntNofact = ?
              AND b.MovExistente = ?
            ORDER BY FECHA DESC
            """

            # Buscar coincidencias de diferentes formas
            patrones = [
                f"%{numero_material}%",  # Cualquier coincidencia parcial
                f"%{numero_material}",  # Coincidencia al final
                f"{numero_material}%"  # Coincidencia al inicio
            ]

            # Versión optimizada
            patrones_unicos = set(patrones)  # Eliminar patrones duplicados si es posible

            # Estructura para verificación rápida de duplicados
            duplicados_verificados = set()

            for patron in patrones_unicos:
                cursor.execute(query, ('20240101', patron, patron, patron, numero_factura, cantidad_umc))
                filas = cursor.fetchall()

                for f in filas:
                    # Crear una clave única para verificar duplicados
                    clave_duplicado = (f.NúmeroMaterial, f.Factura)

                    if clave_duplicado not in duplicados_verificados:
                        resultados_totales.append([
                            f.FECHA, f.ProvNombre, f.FactEntPedimento, f.Factura,
                            f.Folio, f.Caja_CTR, f.NúmeroMaterial, f.Descripción, f.Cantidad,
                            numero_material, cantidad_umc
                        ])
                        duplicados_verificados.add(clave_duplicado)

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

    except Exception as e:
        progress['value'] = 0
        messagebox.showerror("Error", f"Error en búsqueda avanzada:\n{str(e)}")

    # Mostrar resultados
    ventana_resultados = tk.Toplevel()
    ventana_resultados.title(f"Coincidencias Avanzadas en SIADAL ({len(resultados_totales)} resultados)")
    ventana_resultados.geometry("1400x700")

    # Frame principal
    frame_principal = tk.Frame(ventana_resultados, padx=20, pady=20)
    frame_principal.pack(fill="both", expand=True)

    # Frame para controles
    frame_controles = tk.Frame(frame_principal)
    frame_controles.pack(fill="x", pady=(0, 15))

    # Botones con estilo mejorado
    btn_exportar = tk.Button(frame_controles, text="Exportar a Excel", width=20,
                             command=lambda: exportar_resultados(resultados_totales, df_a_evaluar),
                             bg="#4CAF50", fg="white")
    btn_exportar.pack(side="left", padx=5)

    btn_reinyectar = tk.Button(frame_controles, text="Reinyectar Coincidencias", width=20,
                               command=lambda: reinyectar_coincidencias(resultados_totales),
                               bg="#2196F3", fg="white")
    btn_reinyectar.pack(side="left", padx=5)

    btn_cerrar = tk.Button(frame_controles, text="Cerrar", width=20,
                           command=ventana_resultados.destroy,
                           bg="#f44336", fg="white")
    btn_cerrar.pack(side="left", padx=5)

    # Frame para la tabla
    frame_tabla = tk.LabelFrame(frame_principal, text="Resultados de Coincidencias", padx=10, pady=10)
    frame_tabla.pack(fill="both", expand=True)

    # Scrollbars
    scroll_y = tk.Scrollbar(frame_tabla, orient="vertical")
    scroll_x = tk.Scrollbar(frame_tabla, orient="horizontal")

    # Configuración de columnas mejorada
    columnas = [
        ("n", "N°", 50, "center"),
        ("Fecha", "Fecha", 120, "center"),
        ("Proveedor", "Proveedor", 180, "w"),
        ("Pedimento", "Pedimento", 120, "center"),
        ("Factura", "Factura", 120, "center"),
        ("Folio", "Folio", 80, "center"),
        ("Caja_CTR", "Caja/CTR", 100, "center"),
        ("Material_SIADAL", "Material SIADAL", 180, "w"),
        ("Descripcion", "Descripción", 250, "w"),
        ("Cantidad_SIADAL", "Cantidad SIADAL", 100, "center"),
        ("Material_Excel", "Material Excel", 180, "w"),
        ("Cantidad_Excel", "Cantidad Excel", 100, "center"),
        ("Tipo_Operacion", "Tipo Operación", 120, "center")
    ]

    # Estilo para la tabla
    style = ttk.Style()
    style.configure("Treeview", font=('Arial', 10), rowheight=25)
    style.configure("Treeview.Heading", font=('Arial', 10, 'bold'))

    tabla_coincidencias = ttk.Treeview(
        frame_tabla,
        columns=[col[0] for col in columnas],
        show="headings",
        yscrollcommand=scroll_y.set,
        xscrollcommand=scroll_x.set,
        height=20
    )

    # Configurar columnas
    for col, text, width, anchor in columnas:
        tabla_coincidencias.heading(col, text=text)
        tabla_coincidencias.column(col, width=width, anchor=anchor)

    scroll_y.config(command=tabla_coincidencias.yview)
    scroll_x.config(command=tabla_coincidencias.xview)
    scroll_y.pack(side="right", fill="y")
    scroll_x.pack(side="bottom", fill="x")
    tabla_coincidencias.pack(fill="both", expand=True)

    # Insertar datos con colores alternados
    for idx, fila in enumerate(resultados_totales, start=1):
        tipo_op = df_a_evaluar[
            (df_a_evaluar["Número Factura"] == fila[3]) &
            (df_a_evaluar["Número Material"].astype(str).str.strip() == fila[9])
            ]["TipoOperacion"].values[0]

        tags = ('evenrow',) if idx % 2 == 0 else ('oddrow',)

        tabla_coincidencias.insert("", "end", values=[
            idx,
            fila[0],  # Fecha
            fila[1],  # Proveedor
            fila[2],  # Pedimento
            fila[3],  # Factura
            fila[4],  # Folio
            fila[5],  # Caja_CTR
            fila[6],  # Material_SIADAL
            fila[7],  # Descripción
            f"{fila[8]:.2f}",  # Cantidad SIADAL
            fila[9],  # Material Excel
            f"{fila[10]:.2f}",  # Cantidad Excel
            tipo_op  # Tipo Operación
        ], tags=tags)

    # Configurar colores alternados
    tabla_coincidencias.tag_configure('oddrow', background='white')
    tabla_coincidencias.tag_configure('evenrow', background='#f5f5f5')

    # Contador de resultados
    lbl_contador = tk.Label(frame_principal,
                            text=f"Total de coincidencias encontradas: {len(resultados_totales)}",
                            font=('Arial', 10, 'bold'))
    lbl_contador.pack(pady=(10, 0))

def reinyectar_coincidencias(resultados):
    """Función mejorada para reinyectar coincidencias"""
    global materiales_siadal, materiales_encontrados_avanzados

    # Extraer materiales únicos encontrados
    nuevos_materiales = {str(fila[6]).strip() for fila in resultados if str(fila[6]).strip()}

    if not nuevos_materiales:
        messagebox.showwarning("Advertencia", "No hay coincidencias para reinyectar")
        return

    # Agregar a los conjuntos globales
    materiales_siadal.update(nuevos_materiales)
    materiales_encontrados_avanzados.update(nuevos_materiales)

    # Actualizar vista principal
    aplicar_filtro()

    # Resaltar coincidencias en la tabla principal
    for row in tabla.get_children():
        valores = tabla.item(row)['values']
        if len(valores) > 1 and valores[1] in nuevos_materiales:
            tabla.item(row, tags=("encontrado",))

    tabla.tag_configure("encontrado", background="lightgreen")

    # Mostrar resumen
    resumen = (
        f"Se reinyectaron {len(nuevos_materiales)} materiales:\n"
        f"- Coincidencias exactas: {len([m for m in nuevos_materiales if m in df_original['Número Material'].astype(str).str.strip().values])}\n"
        f"- Coincidencias parciales: {len(nuevos_materiales) - len([m for m in nuevos_materiales if m in df_original['Número Material'].astype(str).str.strip().values])}"
    )

    messagebox.showinfo("Éxito", resumen)

# INICIO MATERIALES CON/SIN ESPACIOS
def procesar_patrones(patrones, cursor, numero_factura, cantidad_umc, numero_material):
    resultados_totales = []
    duplicados_verificados = set()

    # Primera pasada: búsqueda normal
    for patron in set(patrones):  # Usamos set() para eliminar patrones duplicados
        query = """
        SELECT FECHA, ProvNombre, FactEntPedimento, Factura, Folio, Caja_CTR, 
               NúmeroMaterial, Descripción, Cantidad
        FROM siadalgoglobaluser.tblMaterial
        WHERE FECHA >= ? 
        AND (NúmeroMaterial = ? OR Factura = ? OR FactEntPedimento = ?)
        AND Factura = ?
        AND Cantidad = ?
        """

        cursor.execute(query, ('20240101', patron, patron, patron, numero_factura, cantidad_umc))
        filas = cursor.fetchall()

        for f in filas:
            clave_duplicado = (f.NúmeroMaterial, f.Factura)
            if clave_duplicado not in duplicados_verificados:
                resultados_totales.append([
                    f.FECHA, f.ProvNombre, f.FactEntPedimento, f.Factura,
                    f.Folio, f.Caja_CTR, f.NúmeroMaterial, f.Descripción, f.Cantidad,
                    numero_material, cantidad_umc, "Existente (SQL)"
                ])
                duplicados_verificados.add(clave_duplicado)

    # Segunda pasada: validación de materiales con espacios
    resultados_totales = validar_materiales_con_espacios(resultados_totales, cursor)

    return resultados_totales

def validar_materiales_con_espacios(resultados_totales, cursor):
    query_sin_espacios = """
    SELECT FECHA, ProvNombre, FactEntPedimento, Factura, Folio, Caja_CTR, 
           NúmeroMaterial, Descripción, Cantidad
    FROM siadalgoglobaluser.tblMaterial
    WHERE REPLACE(NúmeroMaterial, ' ', '') = ?
    AND FECHA >= ?
    """

    marcados = []
    materiales_procesados = set()

    for registro in resultados_totales:
        material_excel = registro[9]  # Posición del material de Excel
        if not material_excel:  # Si no hay material de Excel, usar el de SIADAL
            material_excel = registro[6]

        material_sin_espacios = material_excel.replace(" ", "")

        # Si el material tiene espacios y no lo hemos procesado aún
        if (" " in material_excel and material_excel != material_sin_espacios
                and material_sin_espacios not in materiales_procesados):

            cursor.execute(query_sin_espacios, (material_sin_espacios, '20240101'))
            filas = cursor.fetchall()

            for f in filas:
                registro_marcado = [
                    f.FECHA, f.ProvNombre, f.FactEntPedimento, f.Factura,
                    f.Folio, f.Caja_CTR, f.NúmeroMaterial, f.Descripción, f.Cantidad,
                    material_excel,  # Material original con espacios
                    registro[10],  # Cantidad UMC
                    "VALIDADO (Espacios)"  # Estado en rojo
                ]
                marcados.append(registro_marcado)

            materiales_procesados.add(material_sin_espacios)

    # Añadir los nuevos registros marcados
    resultados_totales.extend(marcados)
    return resultados_totales

# INICIO COINCIDENCIAS COMPLETAS
def buscar_coincidencias_completas(patrones, cursor, numero_factura, cantidad_umc, numero_material_excel):
    resultados_totales = []
    materiales_procesados = set()

    # Normalizamos el material de Excel (eliminamos espacios para comparación)
    material_normalizado = numero_material_excel.replace(" ", "") if numero_material_excel else ""

    # Consulta SQL optimizada para buscar por diferentes campos
    query = """
    SELECT FECHA, ProvNombre, FactEntPedimento, Factura, Folio, Caja_CTR, 
           NúmeroMaterial, Descripción, Cantidad,
           REPLACE(NúmeroMaterial, ' ', '') AS MaterialSinEspacios
    FROM siadalgoglobaluser.tblMaterial
    WHERE FECHA >= '20240101'
    AND (NúmeroMaterial = ? OR Factura = ? OR FactEntPedimento = ? 
         OR REPLACE(NúmeroMaterial, ' ', '') = ?)
    AND Factura = ?
    AND Cantidad = ?
    """

    for patron in set(patrones):
        # Buscamos tanto el patrón original como su versión sin espacios
        patron_sin_espacios = patron.replace(" ", "")

        cursor.execute(query, (
            patron, patron, patron, patron_sin_espacios,
            numero_factura, cantidad_umc
        ))

        filas = cursor.fetchall()

        for f in filas:
            # Creamos una clave única para evitar duplicados
            clave_unica = (f.Factura, f.MaterialSinEspacios, f.Cantidad)

            if clave_unica not in materiales_procesados:
                # Determinamos el estado según el tipo de coincidencia
                if f.MaterialSinEspacios == material_normalizado:
                    estado = "COINCIDENCIA EXACTA"
                elif patron_sin_espacios == f.MaterialSinEspacios:
                    estado = "COINCIDENCIA POR PATRÓN"
                else:
                    estado = "COINCIDENCIA POR OTRO CAMPO"

                resultados_totales.append([
                    f.FECHA, f.ProvNombre, f.FactEntPedimento, f.Factura,
                    f.Folio, f.Caja_CTR, f.NúmeroMaterial, f.Descripción, f.Cantidad,
                    numero_material_excel,  # Conservamos el material original con espacios
                    cantidad_umc,
                    estado
                ])

                materiales_procesados.add(clave_unica)

    # Validación adicional para materiales con espacios no encontrados en la primera pasada
    if " " in numero_material_excel and material_normalizado not in materiales_procesados:
        query_espacios = """
        SELECT FECHA, ProvNombre, FactEntPedimento, Factura, Folio, Caja_CTR, 
               NúmeroMaterial, Descripción, Cantidad
        FROM siadalgoglobaluser.tblMaterial
        WHERE REPLACE(NúmeroMaterial, ' ', '') = ?
        AND FECHA >= '20240101'
        AND Factura = ?
        AND Cantidad = ?
        """

        cursor.execute(query_espacios, (material_normalizado, numero_factura, cantidad_umc))
        filas_espacios = cursor.fetchall()

        for f in filas_espacios:
            clave_unica = (f.Factura, material_normalizado, f.Cantidad)

            if clave_unica not in materiales_procesados:
                resultados_totales.append([
                    f.FECHA, f.ProvNombre, f.FactEntPedimento, f.Factura,
                    f.Folio, f.Caja_CTR, f.NúmeroMaterial, f.Descripción, f.Cantidad,
                    numero_material_excel,
                    cantidad_umc,
                    "COINCIDENCIA IGNORANDO ESPACIOS"
                ])

                materiales_procesados.add(clave_unica)

    return resultados_totales

# INICIO EXPORTAR RESULTADOS
def exportar_resultados(resultados, df_originales):
    # Crear DataFrame para exportar
    datos_exportar = []

    for res in resultados:
        datos_exportar.append({
            "Fecha": res[0],
            "Proveedor": res[1],
            "Pedimento": res[2],
            "Factura": res[3],
            "Folio": res[4],
            "Caja_CTR": res[5],
            "Material_SIADAL": res[6],
            "Descripcion": res[7],
            "Cantidad_SIADAL": res[8],
            "Material_Excel": res[9],
            "Cantidad_Excel": res[10],
            "Tipo_Operacion": df_originales[
                (df_originales["Número Factura"] == res[3]) &
                (df_originales["Número Material"].astype(str).str.strip() == res[9])
                ]["TipoOperacion"].values[0]
        })

    df_export = pd.DataFrame(datos_exportar)

    archivo = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Guardar coincidencias avanzadas"
    )

    if archivo:
        try:
            df_export.to_excel(archivo, index=False)
            messagebox.showinfo("Éxito", "Datos exportados correctamente")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar: {str(e)}")

def actualizar_vista_principal():
    # Actualizar la tabla principal con los nuevos materiales encontrados
    if df_original is not None:
        aplicar_filtro()  # Esto refrescará la vista

        # Resaltar los nuevos materiales encontrados
        for row in tabla.get_children():
            valores = tabla.item(row)['values']
            if len(valores) > 1 and valores[1] in materiales_encontrados_avanzados:
                tabla.item(row, tags=("encontrado",))

        tabla.tag_configure("encontrado", background="lightgreen")

# INICIO INTERFAZ GRÁFICA
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
    columns=["n", "material", "factura", "umc", "pedimento", "tipo", "estado"],
    show="headings",
    height=20,
    yscrollcommand=scroll_y.set,
    xscrollcommand=scroll_x.set
)

# Configurar columnas
tabla.heading("n", text ="N°")
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

btn_estadisticas = tk.Button(frame_botones, text="Ver Estadísticas", width=25,
                            command=mostrar_estadisticas)
btn_estadisticas.pack(side="left", padx=5)

btn_comparar = tk.Button(frame_botones, text="Comparar Excel vs SIADAL", width=25,
                        command=mostrar_comparacion_excel_siadal)
btn_comparar.pack(side="left", padx=5)

# Inicia el bucle principal de la aplicación
try:
    ventana.mainloop()
except KeyboardInterrupt:
    print("Programa detenido por el usuario.")