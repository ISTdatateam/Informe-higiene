import os
import pyodbc
import pandas as pd
import logging

# Configuración de la base de datos utilizando variables de entorno
server = os.getenv('DB_SERVER', '170.110.40.38')
database = os.getenv('DB_DATABASE', 'ept_modprev')
username = os.getenv('DB_USERNAME', 'usr_ept_modprev')
password = os.getenv('DB_PASSWORD', 'C(Q5N:6+5sIt')
driver = '{ODBC Driver 17 for SQL Server}'

# Configuración básica de logging (ajústalo según tus necesidades)
logging.basicConfig(level=logging.INFO)

def get_db_connection():
    """
    Establece una conexión con la base de datos SQL Server.
    Retorna el objeto de conexión si es exitosa; en caso de error, lanza una excepción.
    """
    try:
        connection = pyodbc.connect(
            f'DRIVER={driver};'
            f'SERVER={server};'
            f'DATABASE={database};'
            f'UID={username};'
            f'PWD={password}'
        )
        logging.info("Conexión a la base de datos establecida exitosamente.")
        return connection
    except pyodbc.Error as e:
        logging.error(f"Error al conectar a la base de datos: {e}")
        raise

def get_centro(cuv: str) -> pd.DataFrame:
    """
    Obtiene la información del centro de trabajo (tabla higiene_Centros_Trabajo)
    para el CUV indicado.
    """
    connection = get_db_connection()
    query = "SELECT * FROM higiene_Centros_Trabajo WHERE cuv = ?"
    df = pd.read_sql(query, connection, params=[cuv])
    connection.close()
    return df

def get_visita(cuv: str) -> pd.DataFrame:
    """
    Obtiene las visitas de la tabla higiene_Visitas para el CUV indicado, ordenadas
    de forma descendente por fecha y hora de visita para seleccionar la visita más reciente.
    """
    connection = get_db_connection()
    query = """
        SELECT * FROM higiene_Visitas 
        WHERE cuv_visita = ? 
        ORDER BY fecha_visita DESC, hora_visita DESC
    """
    df = pd.read_sql(query, connection, params=[cuv])
    connection.close()
    return df

def get_mediciones(visita_id: int) -> pd.DataFrame:
    """
    Obtiene las mediciones de la tabla higiene_Mediciones asociadas a la visita indicada.
    """
    connection = get_db_connection()
    query = "SELECT * FROM higiene_Mediciones WHERE visita_id = ?"
    # Convertimos visita_id a entero (tipo int) para evitar el error de parámetro
    df = pd.read_sql(query, connection, params=[int(visita_id)])
    connection.close()
    return df

def get_equipos() -> pd.DataFrame:
    """
    Obtiene toda la información de los equipos de medición de la tabla higiene_Equipos_Medicion.
    """
    connection = get_db_connection()
    query = "SELECT * FROM higiene_Equipos_Medicion"
    df = pd.read_sql(query, connection)
    connection.close()
    return df

def get_all_cuvs_with_visits():
    """Obtiene todos los CUV únicos que tienen visitas registradas."""
    connection = get_db_connection()
    query = "SELECT DISTINCT cuv_visita FROM higiene_Visitas"
    df = pd.read_sql(query, connection)
    connection.close()
    return df['cuv_visita'].tolist()