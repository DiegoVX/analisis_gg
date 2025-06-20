import pyodbc


# Configuración de la conexión
server = 'PRACTICAS_TI\\MSSQLSERVER1'  #Cambiar dependiendo del usuario
database = 'dbSiadalGoGlobal'
username = 'sa'
password = 'root'

try:
    # Crear la conexión
    conexion = pyodbc.connect(
        f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
    )
    print("Conexión exitosa a SQL Server")

    # Crear un cursor para ejecutar consultas
    cursor = conexion.cursor()

    # Ejemplo: Ejecutar una consulta
    cursor.execute("SELECT MatNoParte FROM siadalgoglobaluser.tblMaterial")
    for fila in cursor.fetchall():
        print(fila[0])

    # Cerrar la conexión
    cursor.close()
    conexion.close()

except Exception as e:
    print(f"Error al conectar con SQL Server: {e}")