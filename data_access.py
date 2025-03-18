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
        SELECT * FROM higiene_Visitas_prod
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
    query = "SELECT * FROM higiene_mediciones_prod WHERE visita_id = ?"
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


def insertar_visita(cuv, fecha_visita, hora_medicion, temp_max, motivo_evaluacion,
                    nombre_personal, cargo, consultor_ist, cod_equipo_t, cod_equipo_v,
                    patron_tbs, verif_tbs_inicial, patron_tbh, verif_tbh_inicial,
                    patron_tg, verif_tg_inicial, consultor_cargo, consultor_zonal):

    connection = get_db_connection()
    cursor = connection.cursor()

    try:
        cursor.execute("SELECT id_equipo FROM higiene_Equipos_Medicion WHERE equipo_dicc = ?", (cod_equipo_t,))
        row_temp = cursor.fetchone()
        id_equipo_t = row_temp[0] if row_temp else None

        cursor.execute("SELECT id_equipo FROM higiene_Equipos_Medicion WHERE equipo_dicc = ?", (cod_equipo_v,))
        row_vel = cursor.fetchone()
        id_equipo_v = row_vel[0] if row_vel else None

        if id_equipo_t is None or id_equipo_v is None:
            logging.error(f"No se encontraron equipos en higiene_Equipos_Medicion para: "
                          f"T={cod_equipo_t}, V={cod_equipo_v}")
            return None

        insert_query = """
        INSERT INTO higiene_Visitas_prod (
            cuv_visita, fecha_visita, hora_visita, temperatura_dia, motivo_evaluacion,
            nombre_personal_visita, cargo_personal_visita, consultor_ist,
            equipo_temp, equipo_vel_air, patron_tbs, ver_tbs_ini, 
            patron_tbh, ver_tbh_ini, patron_tg, ver_tg_ini, consultor_cargo, consultor_zonal
        ) 
        OUTPUT INSERTED.id_visita
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """

        cursor.execute(insert_query, (
            cuv, fecha_visita, hora_medicion, temp_max, motivo_evaluacion,
            nombre_personal, cargo, consultor_ist, id_equipo_t, id_equipo_v,
            patron_tbs, verif_tbs_inicial, patron_tbh, verif_tbh_inicial,
            patron_tg, verif_tg_inicial, consultor_cargo, consultor_zonal
        ))

        id_visita = cursor.fetchone()[0]

        connection.commit()
        logging.info(f"Visita insertada exitosamente con id_visita: {id_visita}")

        return id_visita

    except pyodbc.Error as e:
        logging.error(f"Error al insertar la visita: {e}")
        connection.rollback()
        return None

    finally:
        cursor.close()
        connection.close()

def insert_verif_final_visita(id_visita, verif_tbs_final, verif_tbh_final, verif_tg_final, comentarios_finales):
    connection = get_db_connection()
    cursor = connection.cursor()

    try:
        sql = """
        UPDATE higiene_Visitas_prod
        SET ver_tbs_fin = ?, 
            ver_tbh_fin = ?, 
            ver_tg_fin = ?, 
            note_visita = ?
        WHERE id_visita = ?
        """

        cursor.execute(sql, (verif_tbs_final, verif_tbh_final, verif_tg_final, comentarios_finales, id_visita))
        connection.commit()

        if cursor.rowcount > 0:
            logging.info(f"Registro de id_visita {id_visita} actualizado correctamente.")
            return True
        else:
            logging.warning(f"No se encontró el registro con id_visita {id_visita} o los datos no cambiaron.")
            return False

    except Exception as e:
        logging.error(f"Error al actualizar la visita {id_visita}: {e}")
        return False

    finally:
        connection.close()


def insertar_medicion(visita_id, nombre_area, sector_especifico, puesto_trabajo,
                      posicion_trabajador, vestimenta_trabajador, t_bul_seco, t_globo,
                      hum_rel, vel_air, ppd, pmv, resultado_medicion, cond_techumbre,
                      obs_techumbre, cond_paredes, obs_paredes, cond_vantanal, obs_ventanal,
                      cond_aire_acond, obs_aire_acond, cond_ventiladores, obs_ventiladores,
                      cond_inyeccion_extraccion, obs_inyeccion_extraccion, cond_ventanas,
                      obs_ventanas, cond_puertas, obs_puertas, cond_otras, obs_otras, met, clo):

    connection = get_db_connection()
    cursor = connection.cursor()

    try:
        if not visita_id or nombre_area == "Seleccione..." or sector_especifico == "Seleccione..." or puesto_trabajo == "Seleccione...":
            logging.error("Datos de medición incompletos. No se insertará en la base de datos.")
            return None

        insert_query = """
        INSERT INTO higiene_mediciones_prod (
            visita_id, nombre_area, sector_especifico, puesto_trabajo, posicion_trabajador, vestimenta_trabajador,
            t_bul_seco, t_globo, hum_rel, vel_air, ppd, pmv, resultado_medicion, cond_techumbre, obs_techumbre,
            cond_paredes, obs_paredes, cond_vantanal, obs_ventanal, cond_aire_acond, obs_aire_acond,
            cond_ventiladores, obs_ventiladores, cond_inyeccion_extraccion, obs_inyeccion_extraccion,
            cond_ventanas, obs_ventanas, cond_puertas, obs_puertas, cond_otras, obs_otras, met, clo
        ) 
        OUTPUT INSERTED.id_medicion
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """

        cursor.execute(insert_query, (
            visita_id, nombre_area, sector_especifico, puesto_trabajo, posicion_trabajador, vestimenta_trabajador,
            t_bul_seco, t_globo, hum_rel, vel_air, ppd, pmv, resultado_medicion, cond_techumbre, obs_techumbre,
            cond_paredes, obs_paredes, cond_vantanal, obs_ventanal, cond_aire_acond, obs_aire_acond,
            cond_ventiladores, obs_ventiladores, cond_inyeccion_extraccion, obs_inyeccion_extraccion,
            cond_ventanas, obs_ventanas, cond_puertas, obs_puertas, cond_otras, obs_otras, met, clo
        ))

        id_medicion = cursor.fetchone()[0]  # Obtener el ID insertado
        connection.commit()
        logging.info(f"Medición insertada con éxito. ID: {id_medicion}")

        return id_medicion

    except pyodbc.Error as e:
        logging.error(f"Error al insertar la medición: {e}")
        connection.rollback()
        return None

    finally:
        cursor.close()
        connection.close()