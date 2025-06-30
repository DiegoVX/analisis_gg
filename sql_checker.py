# sql_checker.py
import pyodbc
from tkinter import messagebox

def buscar_coincidencia_siadal(numero_material, numero_factura, cantidad_umc):
    try:
        conexion = pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                'SERVER=PRACTICAS_TI\\MSSQLSERVER1;'
                'DATABASE=dbSiadalGoGlobal;'
                'UID=sa;PWD=root'
            )
        cursor = conexion.cursor()

        base_like = numero_material[:4] + "%"

        query = """
        SELECT c.FactEntFechaEnt AS FECHA, p.ProvNombre, c.FactEntPedimento, c.FactEntNofact, 
               c.FactEntFolio, c.FactEntContenedor AS Caja_CTR, a.MatNoParte, a.MatDescr, 
               SUM(b.MovExistente) as qty, c.idAlmacen
        FROM siadalgoglobaluser.tblMaterial AS a
        INNER JOIN siadalgoglobaluser.tblMovimientos AS b ON a.idMaterial = b.idMaterial
        INNER JOIN siadalgoglobaluser.tblFacturaEnt AS c ON b.idFactEnt = c.idFactEnt
        INNER JOIN siadalgoglobaluser.tblDescrFactEnt AS d 
            ON a.idMaterial = d.idMaterial AND c.idFactEnt = d.idFactEnt AND b.iddescrFERenTras = d.idDescrFactEnt
        INNER JOIN siadalgoglobaluser.tblProveedor AS p ON c.idProveedor = p.idProveedor
        WHERE (c.idAlmacen IN (1, 8, 15, 17)) 
          AND (c.FactEntFechaEnt >= 20240101)
          AND a.MatNoParte LIKE ?
          AND c.FactEntNofact = ?
        GROUP BY c.FactEntFechaEnt, p.ProvNombre, c.FactEntFechaFactura, 
                 c.FactEntPedimento, c.FactEntNofact, c.FactEntFolio, 
                 c.FactEntContenedor, a.MatDescr, a.MatNoParte, c.idAlmacen
        ORDER BY FECHA DESC
        """

        cursor.execute(query, (base_like, numero_factura))
        resultados = cursor.fetchall()

        # Coincidencia exacta
        for fila in resultados:
            matno = str(fila.MatNoParte).strip()
            qty = float(fila.qty)
            if matno == numero_material and qty == cantidad_umc:
                return {"match": "exacto", "matno": matno}

        # Coincidencia parcial (cantidad coincide)
        for fila in resultados:
            matno = str(fila.MatNoParte).strip()
            qty = float(fila.qty)
            if qty == cantidad_umc:
                return {"match": "parcial", "matno": matno}

        # Coincidencia solo por factura (aunque el material no coincida)
        if resultados:
            matno = str(resultados[0].MatNoParte).strip()
            return {"match": "ninguna", "matno": matno}

        return {"match": "ninguna", "matno": ""}

    except Exception as e:
        messagebox.showerror("Error SQL", f"No se pudo consultar la base de datos:\n{e}")
        return {"match": "ninguna", "matno": ""}