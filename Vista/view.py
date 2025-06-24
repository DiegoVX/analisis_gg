import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class DataView:
    """Vista para la interfaz gráfica de la aplicación."""

    def __init__(self, root):
        self.root = root
        self.root.title("Analizador de Datos")
        self.root.geometry("950x650")

        self.tabla = None
        self.progress = None
        self.filtro_operacion = None
        self.btn_verificar_sql = None
        self.btn_buscar_avanzadas = None
        self.setup_ui()

    def setup_ui(self):
        """Configura la interfaz gráfica principal."""
        # Botón para cargar Excel
        btn_cargar = tk.Button(self.root, text="Cargar Excel", command=lambda: self.controller.cargar_excel())
        btn_cargar.pack(pady=10)

        # Filtro de tipo de operación
        frame_filtro = tk.Frame(self.root)
        frame_filtro.pack()
        tk.Label(frame_filtro, text="Filtrar por Tipo de Operación:").pack(side="left", padx=5)
        self.filtro_operacion = ttk.Combobox(frame_filtro, values=["Todos", "Importación", "Exportación"])
        self.filtro_operacion.set("Todos")
        self.filtro_operacion.pack(side="left")
        self.filtro_operacion.bind("<<ComboboxSelected>>", lambda e: self.controller.aplicar_filtro())

        # Frame de la tabla
        frame_tabla = tk.Frame(self.root)
        frame_tabla.pack(fill="both", expand=True, padx=20, pady=20)
        scroll_y = tk.Scrollbar(frame_tabla, orient="vertical")
        scroll_x = tk.Scrollbar(frame_tabla, orient="horizontal")

        self.tabla = ttk.Treeview(
            frame_tabla,
            columns=["n", "material", "factura", "umc", "pedimento", "tipo", "estado"],
            show="headings",
            height=20,
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set
        )
        self.tabla.heading("n", text="N°")
        self.tabla.heading("material", text="Número Material")
        self.tabla.heading("factura", text="Número Factura")
        self.tabla.heading("umc", text="Cantidad UMC")
        self.tabla.heading("pedimento", text="Número Pedimento")
        self.tabla.heading("tipo", text="Tipo Operación")
        self.tabla.heading("estado", text="Estado")
        self.tabla.column("n", width=50)
        self.tabla.column("material", width=150)
        self.tabla.column("factura", width=100)
        self.tabla.column("umc", width=80)
        self.tabla.column("pedimento", width=100)
        self.tabla.column("tipo", width=100)
        self.tabla.column("estado", width=120)

        scroll_y.config(command=self.tabla.yview)
        scroll_x.config(command=self.tabla.xview)
        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        self.tabla.pack(fill="both", expand=True)

        # Barra de progreso
        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=400, mode="determinate")
        self.progress.pack(pady=10)

        # Botones
        frame_botones = tk.Frame(self.root)
        frame_botones.pack(side="bottom", pady=10, fill="x")
        btn_guardar = tk.Button(frame_botones, text="Guardar resultados", width=25,
                                command=lambda: self.controller.guardar_datos())
        btn_guardar.pack(side="left", padx=5, fill="x", expand=True)
        self.btn_verificar_sql = tk.Button(frame_botones, text="Verificar Materiales con SQL", width=25,
                                           command=lambda: self.controller.consultar_y_colorear(), state="disabled")
        self.btn_verificar_sql.pack(side="left", padx=5, fill="x", expand=True)
        self.btn_buscar_avanzadas = tk.Button(frame_botones, text="Buscar Coincidencias Avanzadas", width=25,
                                              command=lambda: self.controller.buscar_coincidencias_avanzadas(),
                                              state="disabled")
        self.btn_buscar_avanzadas.pack(side="left", padx=5, fill="x", expand=True)
        btn_estadisticas = tk.Button(frame_botones, text="Ver Estadísticas", width=25,
                                     command=lambda: self.controller.mostrar_estadisticas())
        btn_estadisticas.pack(side="left", padx=5, fill="x", expand=True)
        btn_comparar = tk.Button(frame_botones, text="Comparar Excel vs SIADAL", width=25,
                                 command=lambda: self.controller.mostrar_comparacion_excel_siadal())
        btn_comparar.pack(side="left", padx=5, fill="x", expand=True)

        # Configurar estilos de la tabla
        self.tabla.tag_configure("verde", background="lightgreen")
        self.tabla.tag_configure("rojo", background="salmon")
        self.tabla.tag_configure("encontrado", background="lightgreen")
        self.tabla.tag_configure("sin_coincidencia", background="#FFCCCB")
        self.tabla.tag_configure("sql_match", background="#90EE90")
        self.tabla.tag_configure("avanzado_match", background="#ADD8E6")
        self.tabla.tag_configure("oddrow", background="white")
        self.tabla.tag_configure("evenrow", background="#f5f5f5")

    def mostrar_datos(self, datos):
        """Muestra los datos en la tabla principal."""
        if not hasattr(self, "controller"):
            return

        for row in self.tabla.get_children():
            self.tabla.delete(row)

        for idx, row in datos.iterrows():
            numero_material = str(row["Número Material"]).strip() if pd.notna(row["Número Material"]) else ""
            estado = "Existente" if numero_material in self.controller.model.materiales_siadal else "No Existe"
            tag = "verde" if estado == "Existente" else "rojo"
            self.tabla.insert("", "end", values=(
                idx + 1,
                numero_material,
                row["Número Factura"],
                row["Cantidad UMC"],
                str(row["NumeroPedimento"]) if pd.notna(row["NumeroPedimento"]) else "",
                row["TipoOperacion"],
                estado
            ), tags=(tag,))

    def mostrar_estadisticas(self, df_filtrado):
        """Muestra estadísticas en una ventana emergente con tabla y gráfico circular."""
        ventana_estadisticas = tk.Toplevel(self.root)
        ventana_estadisticas.title("Estadísticas de Materiales")
        ventana_estadisticas.geometry("500x500")

        total_registros = len(df_filtrado)
        materiales_excel = df_filtrado["Número Material"].astype(str).str.strip()
        encontrados = sum(1 for m in materiales_excel if m in self.controller.model.materiales_siadal)
        no_encontrados = total_registros - encontrados
        porcentaje_encontrados = round((encontrados / total_registros) * 100, 2) if total_registros > 0 else 0
        porcentaje_no_encontrados = round((no_encontrados / total_registros) * 100, 2) if total_registros > 0 else 0

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

        fig, ax = plt.subplots(figsize=(4, 4))
        if total_registros > 0:
            ax.pie([encontrados, no_encontrados], labels=["Existentes", "No encontrados"],
                   colors=["lightgreen", "salmon"], autopct='%1.1f%%', startangle=90)
            ax.axis('equal')
        else:
            ax.text(0.5, 0.5, "No hay datos para mostrar", ha='center', va='center')

        canvas = FigureCanvasTkAgg(fig, master=ventana_estadisticas)
        canvas.draw()
        canvas.get_tk_widget().pack(pady=30)

    def mostrar_comparacion_excel_siadal(self, df_filtrado, tipo_operacion):
        """Muestra la comparación entre Excel y SIADAL en una ventana emergente."""
        ventana_comparacion = tk.Toplevel(self.root)
        titulo = f"Comparación ({len(df_filtrado)} registros de {tipo_operacion})" if tipo_operacion != "Todos" else f"Comparación Completa (Todos los {len(df_filtrado)} registros)"
        ventana_comparacion.title(titulo)
        ventana_comparacion.geometry("900x650")

        frame_comparacion = tk.Frame(ventana_comparacion)
        frame_comparacion.pack(fill="both", expand=True, padx=10, pady=10)
        scroll_y = tk.Scrollbar(frame_comparacion, orient="vertical")
        scroll_x = tk.Scrollbar(frame_comparacion, orient="horizontal")

        tabla_comp = ttk.Treeview(
            frame_comparacion,
            columns=["n", "material", "factura", "cantidad", "pedimento", "siadal", "tipo", "estado"],
            show="headings",
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set,
            height=25
        )

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

        for idx, row in df_filtrado.iterrows():
            material_excel = str(row["Número Material"]).strip() if pd.notna(row["Número Material"]) else ""
            factura = str(row["Número Factura"])
            cantidad = str(row["Cantidad UMC"])
            pedimento = str(row["NumeroPedimento"]) if pd.notna(row["NumeroPedimento"]) else ""
            tipo_op = row["TipoOperacion"]
            material_siadal, estado, tag = (
                (material_excel, "Existente (SQL)", "sql_match")
                if material_excel in self.controller.model.materiales_siadal
                else (material_excel, "Existente (Avanzado)", "avanzado_match")
                if material_excel in self.controller.model.materiales_encontrados_avanzados
                else ("", "No encontrado", "sin_coincidencia")
            )

            tabla_comp.insert("", "end", values=(
                idx + 1, material_excel, factura, cantidad, pedimento, material_siadal, tipo_op, estado
            ), tags=(tag,))

        frame_contador = tk.Frame(ventana_comparacion)
        frame_contador.pack(pady=5)
        tk.Label(frame_contador, text=f"Total de registros mostrados: {len(df_filtrado)}",
                 font=('Arial', 10, 'bold')).pack()

        btn_exportar = tk.Button(ventana_comparacion, text="Exportar a Excel",
                                 command=lambda: self.controller.exportar_comparacion(df_filtrado))
        btn_exportar.pack(pady=5)

    def mostrar_coincidencias_avanzadas(self, resultados, df_original):
        """Muestra los resultados de las coincidencias avanzadas en una ventana emergente."""
        ventana_resultados = tk.Toplevel(self.root)
        ventana_resultados.title(f"Coincidencias Avanzadas en SIADAL ({len(resultados)} resultados)")
        ventana_resultados.geometry("1400x700")

        frame_principal = tk.Frame(ventana_resultados, padx=20, pady=20)
        frame_principal.pack(fill="both", expand=True)

        frame_controles = tk.Frame(frame_principal)
        frame_controles.pack(fill="x", pady=(0, 15))

        btn_exportar = tk.Button(frame_controles, text="Exportar a Excel", width=20,
                                 command=lambda: self.controller.exportar_resultados(resultados, df_original),
                                 bg="#4CAF50", fg="white")
        btn_exportar.pack(side="left", padx=5)

        btn_reinyectar = tk.Button(frame_controles, text="Reinyectar Coincidencias", width=20,
                                   command=lambda: self.controller.reinyectar_coincidencias(resultados),
                                   bg="#2196F3", fg="white")
        btn_reinyectar.pack(side="left", padx=5)

        btn_cerrar = tk.Button(frame_controles, text="Cerrar", width=20,
                               command=ventana_resultados.destroy,
                               bg="#f44336", fg="white")
        btn_cerrar.pack(side="left", padx=5)

        frame_tabla = tk.LabelFrame(frame_principal, text="Resultados de Coincidencias", padx=10, pady=10)
        frame_tabla.pack(fill="both", expand=True)

        scroll_y = tk.Scrollbar(frame_tabla, orient="vertical")
        scroll_x = tk.Scrollbar(frame_tabla, orient="horizontal")

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

        for col, text, width, anchor in columnas:
            tabla_coincidencias.heading(col, text=text)
            tabla_coincidencias.column(col, width=width, anchor=anchor)

        scroll_y.config(command=tabla_coincidencias.yview)
        scroll_x.config(command=tabla_coincidencias.xview)
        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        tabla_coincidencias.pack(fill="both", expand=True)

        for idx, fila in enumerate(resultados, start=1):
            tipo_op = df_original[
                (df_original["Número Factura"] == fila[3]) &
                (df_original["Número Material"].astype(str).str.strip() == fila[9])
                ]["TipoOperacion"].values[0] if len(df_original[
                                                        (df_original["Número Factura"] == fila[3]) &
                                                        (df_original["Número Material"].astype(str).str.strip() == fila[
                                                            9])
                                                        ]) > 0 else ""

            tags = ('evenrow',) if idx % 2 == 0 else ('oddrow',)
            tabla_coincidencias.insert("", "end", values=[
                idx, fila[0], fila[1], fila[2], fila[3], fila[4], fila[5],
                fila[6], fila[7], f"{fila[8]:.2f}", fila[9], f"{fila[10]:.2f}", tipo_op
            ], tags=tags)

        lbl_contador = tk.Label(frame_principal,
                                text=f"Total de coincidencias encontradas: {len(resultados)}",
                                font=('Arial', 10, 'bold'))
        lbl_contador.pack(pady=(10, 0))

    def update_progress(self, value):
        """Actualiza la barra de progreso."""
        self.progress['value'] = value
        self.root.update()

    def show_message(self, title, message):
        """Muestra un mensaje emergente."""
        messagebox.showinfo(title, message)

    def show_error(self, title, message):
        """Muestra un mensaje de error."""
        messagebox.showerror(title, message)

    def set_controller(self, controller):
        """Establece el controlador para la vista."""
        self.controller = controller