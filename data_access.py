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


def insertar_visita(cuv, fecha_visita, hora_medicion, temp_max, motivo_evaluacion,
                    nombre_personal, cargo, consultor_ist, cod_equipo_t, cod_equipo_v,
                    patron_tbs, verif_tbs_inicial, patron_tbh, verif_tbh_inicial,
                    patron_tg, verif_tg_inicial):
    """
    Inserta una nueva visita en la tabla higiene_Visitas y retorna el id_visita generado.

    - Busca los valores de id_equipo en higiene_Equipos_Medicion para los códigos de equipo (cod_equipo_t y cod_equipo_v).
    - Inserta los datos de la visita en la base de datos.
    - Retorna el id_visita generado.

    Parámetros:
    - cuv: str -> Código único de visita.
    - fecha_visita: date -> Fecha de la visita.
    - hora_medicion: time -> Hora de la medición.
    - temp_max: float -> Temperatura máxima del día.
    - motivo_evaluacion: str -> Motivo de evaluación.
    - nombre_personal: str -> Nombre del personal SMU.
    - cargo: str -> Cargo del personal.
    - consultor_ist: str -> Consultor a cargo.
    - cod_equipo_t: str -> Código del equipo de temperatura (valor de equipo_dicc).
    - cod_equipo_v: str -> Código del equipo de velocidad de aire (valor de equipo_dicc).
    - patron_tbs: float -> Valor patrón de temperatura de bulbo seco.
    - verif_tbs_inicial: float -> Verificación inicial de TBS.
    - patron_tbh: float -> Valor patrón de temperatura de bulbo húmedo.
    - verif_tbh_inicial: float -> Verificación inicial de TBH.
    - patron_tg: float -> Valor patrón de temperatura globo.
    - verif_tg_inicial: float -> Verificación inicial de TG.

    Retorna:
    - id_visita generado (int) si la inserción es exitosa, None en caso de error.
    """
    connection = get_db_connection()
    cursor = connection.cursor()

    try:
        # Buscar id_equipo para cod_equipo_t
        cursor.execute("SELECT id_equipo FROM higiene_Equipos_Medicion WHERE equipo_dicc = ?", (cod_equipo_t,))
        row_temp = cursor.fetchone()
        id_equipo_t = row_temp[0] if row_temp else None

        # Buscar id_equipo para cod_equipo_v
        cursor.execute("SELECT id_equipo FROM higiene_Equipos_Medicion WHERE equipo_dicc = ?", (cod_equipo_v,))
        row_vel = cursor.fetchone()
        id_equipo_v = row_vel[0] if row_vel else None

        # Validar que ambos equipos existan en la base de datos
        if id_equipo_t is None or id_equipo_v is None:
            logging.error(f"No se encontraron equipos en higiene_Equipos_Medicion para: "
                          f"T={cod_equipo_t}, V={cod_equipo_v}")
            return None

        # Insertar la visita en higiene_Visitas
        insert_query = """
        INSERT INTO higiene_Visitas_prod (
            cuv_visita, fecha_visita, hora_visita, temperatura_dia, motivo_evaluacion,
            nombre_personal_visita, cargo_personal_visita, consultor_ist,
            equipo_temp, equipo_vel_air, patron_tbs, ver_tbs_ini, 
            patron_tbh, ver_tbh_ini, patron_tg, ver_tg_ini
        ) 
        OUTPUT INSERTED.id_visita
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """

        cursor.execute(insert_query, (
            cuv, fecha_visita, hora_medicion, temp_max, motivo_evaluacion,
            nombre_personal, cargo, consultor_ist, id_equipo_t, id_equipo_v,
            patron_tbs, verif_tbs_inicial, patron_tbh, verif_tbh_inicial,
            patron_tg, verif_tg_inicial
        ))

        # Obtener el id_visita generado
        id_visita = cursor.fetchone()[0]

        # Confirmar la transacción
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
    """
    Actualiza los valores de verificación final en la tabla higiene_Visitas_prod
    usando el id_visita.
    """

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
    """
    Inserta una medición en la tabla higiene_mediciones_prod y retorna el id_medicion generado.

    Parámetros:
    - visit_id (int): ID de la visita a la que pertenece la medición.
    - nombre_area (str): Nombre del área de medición.
    - sector_especifico (str): Sector específico de la medición.
    - puesto_trabajo (str): Puesto de trabajo medido.
    - posicion_trabajador (str): Posición del trabajador (De pie/Sentado).
    - vestimenta_trabajador (str): Tipo de vestimenta del trabajador.
    - t_bul_seco (float): Temperatura bulbo seco (°C).
    - t_globo (float): Temperatura globo (°C).
    - hum_rel (float): Humedad relativa (%).
    - vel_air (float): Velocidad del aire (m/s).
    - ppd (float): Índice de Predicción de Porcentaje de Personas Insatisfechas.
    - pmv (float): Índice de Voto Medio Predicho.
    - resultado_medicion (str): Interpretación del PMV.
    - cond_techumbre, obs_techumbre, cond_paredes, obs_paredes, cond_vantanal, obs_ventanal,
      cond_aire_acond, obs_aire_acond, cond_ventiladores, obs_ventiladores,
      cond_inyeccion_extraccion, obs_inyeccion_extraccion, cond_ventanas, obs_ventanas,
      cond_puertas, obs_puertas (str): Condiciones y observaciones constructivas del área.
    - met (float): Nivel metabólico del trabajador.
    - clo (float): Nivel de aislamiento térmico de la vestimenta.
    - caract_constructivas (str): Características constructivas del área.
    - ingreso_salida_aire (str): Descripción de ingreso/salida de aire.

    Retorna:
    - id_medicion (int): ID generado de la medición insertada, o None en caso de error.
    """
    connection = get_db_connection()
    cursor = connection.cursor()

    try:
        # Validación de parámetros
        if not visita_id or nombre_area == "Seleccione..." or sector_especifico == "Seleccione..." or puesto_trabajo == "Seleccione...":
            logging.error("Datos de medición incompletos. No se insertará en la base de datos.")
            return None

        # Query de inserción con OUTPUT para obtener el ID generado
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