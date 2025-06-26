from tkinter import filedialog
import time


class DataController:
    """Controlador para coordinar el modelo y la vista."""

    def __init__(self, model, view):
        self.model = model
        self.view = view
        self.view.set_controller(self)

    def cargar_excel(self):
        """Carga un archivo Excel y actualiza la vista."""
        archivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        self.ruta_excel = archivo

        if not archivo:
            return

        self.view.update_progress(10)
        success, message = self.model.cargar_excel(archivo)
        self.view.update_progress(30)

        if success:
            self.view.btn_verificar_sql.config(state="normal")
            self.view.btn_buscar_avanzadas.config(state="normal")

            df_filtrado = self.model.filtrar_datos(self.view.filtro_operacion.get())
            self.view.mostrar_datos(df_filtrado, colorear=False)

            self.view.update_progress(100)
            time.sleep(0.5)
            self.view.update_progress(0)
        else:
            self.view.update_progress(0)
            self.view.show_error("Error", message)

    def aplicar_filtro(self):
        """Aplica el filtro de tipo de operación y muestra los datos."""
        tipo_operacion = self.view.filtro_operacion.get()
        df_filtrado = self.model.filtrar_datos(tipo_operacion)
        if df_filtrado is not None:
            self.view.mostrar_datos(df_filtrado)

    def consultar_y_colorear(self):
        """Consulta los materiales en la base de datos y actualiza la vista."""
        self.view.update_progress(10)
        success, message = self.model.consultar_sql()
        self.view.update_progress(40)

        if success:
            total = len(self.model.df_original)
            procesados = 0
            tipo_operacion = self.view.filtro_operacion.get()
            df_a_mostrar = self.model.filtrar_datos(tipo_operacion)

            for _ in df_a_mostrar.iterrows():
                procesados += 1
                if procesados % max(total // 10, 1) == 0:
                    progreso = 40 + int((procesados / total) * 60)
                    self.view.update_progress(progreso)

            self.aplicar_filtro()
            self.view.update_progress(100)
            time.sleep(0.5)
            self.view.update_progress(0)
        else:
            self.view.update_progress(0)
            self.view.show_error("Error SQL", message)

    def buscar_coincidencias_avanzadas(self):
        """Realiza una búsqueda avanzada de coincidencias no encontradas."""
        self.view.update_progress(10)
        tipo_operacion = self.view.filtro_operacion.get()
        success, message, resultados = self.model.buscar_coincidencias_avanzadas(tipo_operacion)
        self.view.update_progress(100)
        time.sleep(0.3)
        self.view.update_progress(0)

        if success:
            if resultados:
                self.view.mostrar_coincidencias_avanzadas(resultados, self.model.df_original)
            else:
                self.view.show_message("Información", message)
        else:
            self.view.show_error("Error", message)

    def guardar_datos(self):
        """Guarda los datos actualmente visibles en un archivo Excel."""
        filas = [self.view.tabla.item(i)['values'] for i in self.view.tabla.get_children()]
        archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if archivo:
            success, message = self.model.guardar_datos(filas, archivo)
            if success:
                self.view.show_message("Guardado", message)
            else:
                self.view.show_error("Error", message)

    def exportar_comparacion(self, df):
        """Exporta los datos comparativos a un archivo Excel."""
        archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if archivo:
            success, message = self.model.guardar_datos(df, archivo)
            if success:
                self.view.show_message("Éxito", message)
            else:
                self.view.show_error("Error", message)

    def exportar_resultados(self, resultados, df_original):
        """Exporta los resultados de coincidencias avanzadas con detalles."""
        datos_exportar = [{
            "Fecha": res[0], "Proveedor": res[1], "Pedimento": res[2], "Factura": res[3],
            "Folio": res[4], "Caja_CTR": res[5], "Material_SIADAL": res[6],
            "Descripcion": res[7], "Cantidad_SIADAL": res[8], "Material_Excel": res[9],
            "Cantidad_Excel": res[10],
            "Tipo_Operacion": df_original[
                (df_original["Número Factura"] == res[3]) &
                (df_original["Número Material"].astype(str).str.strip() == res[9])
            ]["TipoOperacion"].values[0] if len(df_original[
                (df_original["Número Factura"] == res[3]) &
                (df_original["Número Material"].astype(str).str.strip() == res[9])
            ]) > 0 else ""
        } for res in resultados]

        archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if archivo:
            success, message = self.model.guardar_datos(datos_exportar, archivo)
            if success:
                self.view.show_message("Éxito", message)
            else:
                self.view.show_error("Error", message)

    def reinyectar_coincidencias(self, resultados):
        """Agrega coincidencias encontradas a los conjuntos válidos y actualiza la vista."""
        success, message = self.model.reinyectar_coincidencias(resultados)
        if success:
            self.aplicar_filtro()
            self.view.show_message("Éxito", message)
        else:
            self.view.show_error("Advertencia", message)

    def mostrar_estadisticas(self):
        """Muestra estadísticas generales de los datos filtrados."""
        tipo_operacion = self.view.filtro_operacion.get()
        df_filtrado = self.model.filtrar_datos(tipo_operacion)
        if df_filtrado is not None:
            self.view.mostrar_estadisticas(df_filtrado)

    def mostrar_comparacion_excel_siadal(self):
        """Muestra la comparación entre el archivo Excel y los datos de SIADAL."""
        tipo_operacion = self.view.filtro_operacion.get()
        df_filtrado = self.model.filtrar_datos(tipo_operacion)
        if df_filtrado is not None:
            self.view.mostrar_comparacion_excel_siadal()

    def actualizar_excel_con_siadal(self):
        tipo_operacion = self.view.filtro_operacion.get()
        self.view.update_progress(5)

        ok, msg, resultados = self.model.buscar_coincidencias_avanzadas(tipo_operacion)
        if not ok:
            self.view.show_error("Error", msg)
            return

        self.view.update_progress(40)
        ruta_nueva = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")],
            title="Guardar archivo actualizado con SIADAL"
        )
        if not ruta_nueva.endswith('.xlsx'):
            ruta_nueva += 'xlsx'
            return

        self.view.update_progress(70)
        exito, mensaje = self.model.escribir_resultados_en_excel(self.ruta_excel, ruta_nueva, resultados)

        if exito:
            self.view.show_message("Éxito", mensaje)
        else:
            self.view.show_error("Error", mensaje)

        self.view.update_progress(100)