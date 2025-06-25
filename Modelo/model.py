import pandas as pd
import pyodbc
from openpyxl import load_workbook

class DataModel:
    """Modelo para manejar datos de Excel y consultas SQL."""

    def __init__(self):
        self.df_original = None
        self.materiales_siadal = set()
        self.materiales_encontrados_avanzados = set()
        self.materiales_vistos = set()

    def cargar_varios_excel(self, archivos):
        """Carga múltiples archivos Excel, buscando las hojas que contengan los datos necesarios sin depender de nombres fijos."""
        if not archivos:
            return False, "No se seleccionaron archivos."

        lista_dfs = []

        for archivo in archivos:
            try:
                xl = pd.ExcelFile(archivo, engine="openpyxl")
                hojas = xl.sheet_names

                detalle_df = None
                relacion_df = None
                encabezado_df = None

                for hoja in hojas:
                    df = xl.parse(hoja, header=None)

                    # Detectar "RELACIÓN FAC-PED"
                    if relacion_df is None and "NumeroFactura" in df.iloc[0].values and "NumeroPedimento" in df.iloc[
                        0].values:
                        df.columns = df.iloc[0]
                        relacion_df = df[1:][["NumeroFactura", "NumeroPedimento"]].drop_duplicates(
                            subset=["NumeroFactura"])

                    # Detectar "ENCABEZADO FAC"
                    elif encabezado_df is None and ("NumeroFactura" in df.iloc[0].values or df.shape[1] >= 3):
                        df.columns = ["NumeroFactura", "?", "TipoOperacion"] + list(
                            df.iloc[0][3:])  # Ajustar según tus datos reales
                        encabezado_df = df[1:][["NumeroFactura", "TipoOperacion"]].drop_duplicates(
                            subset=["NumeroFactura"])

                    # Detectar "DETALLE FAC" (buscamos columnas clave conocidas)
                    elif detalle_df is None:
                        df_temp = xl.parse(hoja)
                        columnas_esperadas = {"Número Material", "Número Factura", "Cantidad UMC"}
                        if columnas_esperadas.issubset(set(df_temp.columns)):
                            detalle_df = df_temp

                # Si no se encontró alguna tabla crítica, se omite
                if detalle_df is None or relacion_df is None or encabezado_df is None:
                    print(f"[Advertencia] Archivo omitido por falta de datos suficientes: {archivo}")
                    continue

                # Asegurar tipos
                detalle_df["Número Factura"] = detalle_df["Número Factura"].astype(str)
                relacion_df["NumeroFactura"] = relacion_df["NumeroFactura"].astype(str)
                encabezado_df["NumeroFactura"] = encabezado_df["NumeroFactura"].astype(str)

                # Unir todo
                merged_df = detalle_df.merge(relacion_df, left_on="Número Factura", right_on="NumeroFactura",
                                             how="left")
                merged_df = merged_df.merge(encabezado_df, left_on="Número Factura", right_on="NumeroFactura",
                                            how="left")
                merged_df["TipoOperacion"] = merged_df["TipoOperacion"].map({1: "Importación", 2: "Exportación"})

                final_df = merged_df[[
                    "Número Material", "Número Factura", "Cantidad UMC", "NumeroPedimento", "TipoOperacion"
                ]]

                lista_dfs.append(final_df)

            except Exception as e:
                print(f"[Error] No se pudo procesar el archivo {archivo}: {e}")
                continue

        if not lista_dfs:
            return False, "Ningún archivo válido fue cargado."

        self.df_original = pd.concat(lista_dfs, ignore_index=True)
        return True, f"{len(lista_dfs)} archivo(s) cargado(s) correctamente."

    def filtrar_datos(self, tipo_operacion):
        if self.df_original is None:
            return None
        if tipo_operacion == "Todos":
            return self.df_original
        return self.df_original[self.df_original["TipoOperacion"] == tipo_operacion]

    def consultar_sql(self):
        if self.df_original is None:
            return False, "No hay datos cargados."

        try:
            conexion = pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                'SERVER=PRACTICAS_TI\\MSSQLSERVER1;'
                'DATABASE=dbSiadalGoGlobal;'
                'UID=sa;PWD=root'
            )
            cursor = conexion.cursor()

            cursor.execute("SELECT MatNoParte FROM siadalgoglobaluser.tblMaterial")
            self.materiales_siadal = set(str(r[0]).strip() for r in cursor.fetchall() if r[0])
            self.materiales_vistos = set(str(x).strip() for x in self.df_original["Número Material"].dropna().unique())

            return True, "Consulta SQL completada."
        except Exception as e:
            return False, f"Error: {e}"

    def buscar_coincidencias_avanzadas(self, tipo_operacion):
        if self.df_original is None or not self.materiales_siadal:
            return False, "Datos no cargados o consulta SQL no realizada.", []

        df_filtrado = self.filtrar_datos(tipo_operacion)
        df_no_encontrados = df_filtrado[
            ~df_filtrado["Número Material"].astype(str).str.strip().isin(self.materiales_siadal)
        ]

        if df_no_encontrados.empty:
            return True, "No hay materiales no encontrados.", []

        try:
            conexion = pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                'SERVER=PRACTICAS_TI\\MSSQLSERVER1;'
                'DATABASE=dbSiadalGoGlobal;'
                'UID=sa;PWD=root'
            )
            cursor = conexion.cursor()
            resultados_totales = []
            duplicados_verificados = set()

            for fila in df_no_encontrados.itertuples():
                mat = str(fila._1).strip()
                factura = fila._2
                try:
                    cantidad = float(fila._3)
                except (ValueError, TypeError):
                    continue

                query = f"""
                SELECT 
                    c.FactEntFechaEnt AS FECHA, p.ProvNombre, c.FactEntPedimento, c.FactEntNofact, 
                    c.FactEntFolio, c.FactEntContenedor AS Caja_CTR, a.MatNoParte, a.MatDescr, 
                    SUM(b.MovExistente) as qty
                FROM siadalgoglobaluser.tblMaterial AS a
                INNER JOIN siadalgoglobaluser.tblMovimientos AS b ON a.idMaterial = b.idMaterial
                INNER JOIN siadalgoglobaluser.tblFacturaEnt AS c ON b.idFactEnt = c.idFactEnt
                INNER JOIN siadalgoglobaluser.tblDescrFactEnt AS d ON a.idMaterial = d.idMaterial 
                    AND c.idFactEnt = d.idFactEnt AND b.iddescrFERenTras = d.idDescrFactEnt
                INNER JOIN siadalgoglobaluser.tblProveedor AS p ON c.idProveedor = p.idProveedor
                WHERE (c.idAlmacen = 17) AND (c.FactEntFechaEnt >= 20240101)
                    AND a.MatNoParte LIKE '{mat}%'
                    AND c.FactEntNofact = '{factura}'
                GROUP BY c.FactEntFechaEnt, p.ProvNombre, c.FactEntFechaFactura, c.FactEntPedimento, 
                    c.FactEntNofact, c.FactEntFolio, c.FactEntContenedor, a.MatDescr, a.MatNoParte
                ORDER BY FECHA DESC
                """

                cursor.execute(query)
                for r in cursor.fetchall():
                    clave = (r.MatNoParte, r.FactEntNofact)
                    if clave not in duplicados_verificados:
                        resultados_totales.append([
                            r.FECHA, r.ProvNombre, r.FactEntPedimento, r.FactEntNofact,
                            r.FactEntFolio, r.Caja_CTR, r.MatNoParte, r.MatDescr, float(r.qty),
                            mat, cantidad
                        ])
                        duplicados_verificados.add(clave)
                        self.materiales_encontrados_avanzados.add(mat)

            return True, "Búsqueda avanzada completada.", resultados_totales
        except Exception as e:
            return False, f"Error: {e}", []

    def obtener_materiales_no_encontrados(self, tipo_operacion):
        """Devuelve lista de materiales no encontrados después de consultar SQL."""
        if self.df_original is None or not self.materiales_siadal:
            return []

        df_filtrado = self.filtrar_datos(tipo_operacion)
        no_encontrados = df_filtrado[
            ~df_filtrado["Número Material"].astype(str).str.strip().isin(self.materiales_siadal)
        ]
        return no_encontrados["Número Material"].astype(str).str.strip().unique().tolist()

    def guardar_datos(self, datos, archivo):
        try:
            columnas = [
                "Fecha", "Proveedor", "Pedimento", "Factura",
                "Folio", "Caja", "Número Material", "Descripción", "Cantidad BD",
                "Material Entrada", "Cantidad Entrada"
            ]
            df_guardar = pd.DataFrame(datos, columns=columnas)
            df_guardar.to_excel(archivo, index=False)
            return True, "Datos guardados correctamente."
        except Exception as e:
            return False, f"No se pudo guardar el archivo: {e}"

    def reinyectar_coincidencias(self, resultados):
        nuevos = {str(f[6]).strip() for f in resultados if f[6]}
        if not nuevos:
            return False, "No hay coincidencias para reinyectar."

        self.materiales_siadal.update(nuevos)
        self.materiales_encontrados_avanzados.update(nuevos)
        return True, f"Se reinyectaron {len(nuevos)} materiales."

    def escribir_resultados_en_excel(self, ruta_entrada, ruta_salida, resultados_avanzados):
        try:
            wb = load_workbook(ruta_entrada)
            hoja = wb["DETALLE FAC"]

            # Encontrar la columna de "Número Material"
            encabezados = [cell.value for cell in hoja[1]]
            if "Número Material" not in encabezados:
                return False, "No se encontró la columna 'Número Material' en el archivo."

            col_material = encabezados.index("Número Material") + 1
            nueva_col = col_material + 1

            # Escribir encabezado de la nueva columna
            hoja.cell(row=1, column=nueva_col, value="Número Material SIADAL")

            # Crear diccionario con las coincidencias (clave = material y factura)
            coincidencias_dict = {(str(r[9]).strip(), r[3]): str(r[6]).strip() for r in resultados_avanzados}

            for fila in range(2, hoja.max_row + 1):
                num_mat = hoja.cell(row=fila, column=col_material).value
                factura = hoja.cell(row=fila, column=encabezados.index("Número Factura") + 1).value

                if num_mat and factura:
                    clave = (str(num_mat).strip(), factura)
                    valor_siadal = coincidencias_dict.get(clave, "")
                    hoja.cell(row=fila, column=nueva_col, value=valor_siadal)

            wb.save(ruta_salida)
            return True, "Archivo actualizado correctamente."
        except Exception as e:
            return False, f"Error al escribir en Excel: {e}"