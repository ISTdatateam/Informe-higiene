import os
import logging
import pyodbc
from pythermalcomfort.models import pmv_ppd_iso  # Asegúrate de tener instalada la librería correspondiente

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

def actualizar_pmv_ppd():
    """
    Extrae los registros de la tabla 'higiene_mediciones', calcula pmv y ppd,
    y actualiza la tabla con los nuevos valores.
    """
    try:
        connection = get_db_connection()
        cursor = connection.cursor()

        # Consulta para extraer los campos necesarios
        query = """
            SELECT id_medicion, t_bul_seco, t_globo, vel_air, hum_rel, met, clo
            FROM higiene_mediciones
        """
        cursor.execute(query)
        registros = cursor.fetchall()

        logging.info(f"Se encontraron {len(registros)} registros para actualizar.")

        for registro in registros:
            id_medicion, t_bul_seco, t_globo, vel_air, hum_rel, met, clo = registro

            # Calcular pmv y ppd usando la función pmv_ppd_iso
            resultados = pmv_ppd_iso(
                tdb=t_bul_seco,
                tr=t_globo,
                vr=vel_air,
                rh=hum_rel,
                met=met,
                clo=clo,
                model="7730-2005",
                limit_inputs=False,
                round_output=True
            )
            pmv = resultados.pmv
            ppd = resultados.ppd

            logging.info(f"Registro {id_medicion}: pmv={pmv}, ppd={ppd}")

            # Actualizar la tabla con los nuevos valores
            update_query = """
                UPDATE higiene_mediciones
                SET pmv = ?, ppd = ?
                WHERE id_medicion = ?
            """
            cursor.execute(update_query, (pmv, ppd, id_medicion))

        # Confirmar los cambios en la base de datos
        connection.commit()
        logging.info("Actualización completada exitosamente.")

    except Exception as e:
        logging.error(f"Ocurrió un error durante la actualización: {e}")
    finally:
        # Cerrar la conexión y el cursor si están abiertos
        try:
            cursor.close()
            connection.close()
        except Exception:
            pass

if __name__ == '__main__':
    actualizar_pmv_ppd()
