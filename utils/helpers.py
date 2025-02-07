from db.mysql_utils import MySQLDatabaseManager
import streamlit as st
from streamlit.runtime.scriptrunner import get_script_run_ctx
from streamlit.source_util import get_pages
import bcrypt
import re
import time
from streamlit_cookies_controller import CookieController

cookie_controller = CookieController()

def validar_rut(rut):
    # Implementación del validador
    pass

def autenticar_usuario(username, password):
    db = MySQLDatabaseManager()
    try:
        query = "SELECT name, pass FROM usuarios WHERE email = %s"
        user = db.fetch_one(query, (username,))
        if user and bcrypt.checkpw(password.encode(), user["pass"].encode()):
            return {"nombre": user["name"]}
        return None
    finally:
        db.close()

def get_current_page_name():
    ctx = get_script_run_ctx()
    if ctx is None:
        raise RuntimeError("Couldn't get script context")

    pages = get_pages("")

    return pages[ctx.page_script_hash]["page_name"]


def make_sidebar():
    with st.sidebar:
        if st.session_state.get("data_user", None):
            st.write(f"👋 Hola, {st.session_state['data_user']['name']}!")

            if st.button("Cerrar sesión"):
                cookie_controller.set("user_data", "", max_age=0)
                # cookie_controller.remove("user_data")
                st.session_state["data_user"] = None
                st.success("Has cerrado sesión correctamente.")
                time.sleep(2)
                st.rerun()


        #elif get_current_page_name() != "app":
            # If anyone tries to access a secret page without being logged in,
            # redirect them to the login page


def logout():
    cookie_controller.set("user_data", "", max_age=0)
    #cookie_controller.remove("user_data")
    st.session_state["user_data"] = None
    st.success("Has cerrado sesión correctamente.")
    time.sleep(2)
    st.rerun()

def login(username, password):
    user_data = autenticar_usuario(username, password)
    if user_data:
        # Guardar datos en cookies
        #cookie_controller.set("user_data", user_data, max_age=7200)
        #st.session_state["user_data"] = user_data
        st.success("Login exitoso")
        time.sleep(1)
        #st.switch_page("pages/home.py")
        st.rerun()
    else:
        st.error("Usuario o contraseña incorrectos.")

# Checkear sesión al inicio
def check_session():
    #st.write("check de cookie")
    user_data = cookie_controller.get("user_data")
    #st.write(cookie_controller.get("user_data"))
    if user_data:
        st.session_state["data_user"] = user_data
        #st.write(st.session_state["user_data"])
    return user_data

def get_ct(cuv):
    db = MySQLDatabaseManager()
    try:
        query = """
            SELECT cuv, rut, razon_social, rut2, nombre_ct, direccion_ct, comuna_ct, region_ct, region_num_ct
            FROM centros_trabajo 
            WHERE cuv = %s
        """
        db.cursor.execute(query, (cuv,))
        resultados = db.cursor.fetchall()
        return resultados
    finally:
        db.close()


def precompletar_campos_ct(data_ct):
    razon_social = st.text_input("Razón Social",
                                 value=str(data_ct.get("razon_social", "")),
                                 disabled=True)
    rut = st.text_input("RUT",
                        value=str(data_ct.get("rut", "")),
                        disabled=True)
    nombre_local = st.text_input("Nombre de Local",
                                 value=str(data_ct.get("nombre_ct", "")),
                                 disabled=True)
    direccion = st.text_input("Dirección",
                              value=str(data_ct.get("direccion_ct", "")),
                              disabled=True)
    comuna = st.text_input("Comuna",
                           value=str(data_ct.get("comuna_ct", "")),
                           disabled=True)
    region = st.text_input("Región",
                           value=str(data_ct.get("region_ct", "")),
                           disabled=True)
    cuv_val = st.text_input("CUV",
                            value=str(data_ct.get("cuv", "")),
                            disabled=True)

    return {
        "razon_social": razon_social,
        "rut": rut,
        "nombre_local": nombre_local,
        "direccion": direccion,
        "comuna": comuna,
        "region": region,
        "cuv": cuv_val
    }

def interpreter_pmv(pmv):
    if pmv >= 2.5:
        interpretacion = "Calurosa"
    elif pmv >= 1.5:
        interpretacion = "Cálida"
    elif pmv >= 0.5:
        interpretacion = "Ligeramente cálida"
    elif pmv > -0.5:
        interpretacion = "Neutra - Confortable"
    elif pmv > -1.5:
        interpretacion = "Ligeramente fresca"
    elif pmv > -2.5:
        interpretacion = "Fresca"
    else:
        interpretacion = "Fría"
    return interpretacion

@st.dialog("Confirmación de guardado")
def dialogo_confirmacion():
    st.write("¿Estás segur@ que quieres guardar los datos de la medición?")
    with st.container():
        col1, col2 = st.columns(2)
        aceptar = col1.button("Aceptar", key="dialog_aceptar")
        cancelar = col2.button("Cancelar", key="dialog_cancelar")
    if aceptar:
        st.session_state["confirm_save"] = True
        st.rerun()
    if cancelar:
        st.session_state["confirm_save"] = False
        st.rerun()


def confirmar_guardado():
    if "confirm_save" not in st.session_state:
        st.session_state["confirm_save"] = None
    dialogo_confirmacion()
    return st.session_state.get("confirm_save")