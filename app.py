import streamlit as st
import urllib.parse
import pandas as pd
from datetime import datetime, timedelta
import auth, gastos, ventas, costos, validacion  
from utils import conectar_hoja, obtener_dataframe

st.set_page_config(page_title="Auditoría DTE Pro", layout="wide")

ocultar_elementos = """
    <style>
    .stAppDeployButton {display: none !important;}
    footer {visibility: hidden !important;}
    </style>
    """
st.markdown(ocultar_elementos, unsafe_allow_html=True)

usuarios_validos = auth.obtener_usuarios_db()

# --- Gestión de Sesión ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "last_activity" not in st.session_state:
    st.session_state.last_activity = None
if "usuario_actual" not in st.session_state:
    st.session_state.usuario_actual = ""
if "rol_actual" not in st.session_state:
    st.session_state.rol_actual = ""

usuario_url = st.query_params.get("auth")
if usuario_url:
    usuario_url = urllib.parse.unquote(usuario_url).replace("+", " ").strip().upper()
    if usuario_url in usuarios_validos:
        st.session_state.logged_in = True
        st.session_state.usuario_actual = usuario_url
        st.session_state.rol_actual = usuarios_validos[usuario_url].get("Rol", "USUARIO")
        st.session_state.anexo_actual = usuarios_validos[usuario_url].get("Anexo", "General")
        
        modulos_permitidos = usuarios_validos[usuario_url].get("Modulos", "PEDIDOS CAFETERÍA")
        st.session_state.modulos_permitidos = [m.strip() for m in modulos_permitidos.split(",")]
        
        if st.session_state.last_activity is None:
            st.session_state.last_activity = datetime.now()

if st.session_state.logged_in:
    tiempo_inactivo = datetime.now() - st.session_state.last_activity
    if tiempo_inactivo > timedelta(hours=1):
        st.session_state.logged_in = False
        st.query_params.clear() 
        st.session_state.usuario_actual = ""
        st.session_state.rol_actual = ""
        st.session_state.mensaje_login = "Tu sesión ha expirado. Por favor, inicia sesión nuevamente."
        st.rerun()
    else:
        st.session_state.last_activity = datetime.now()

if not st.session_state.logged_in:
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        st.title("🔒 Acceso al Sistema")
        if "mensaje_login" in st.session_state:
            st.warning(st.session_state.mensaje_login)
            del st.session_state.mensaje_login

        with st.form("login_form"):
            usuario = st.text_input("Usuario")
            contrasena = st.text_input("Contraseña", type="password")
            btn_ingresar = st.form_submit_button("Ingresar", use_container_width=True)

            # --- Sidebar y Menú ---
with st.sidebar:
    st.title("Menú Principal")
    
    todos_los_modulos = ["KPI DE REGISTROS", "KPI DE VENTAS", "CONTABILIDAD DE COSTOS", "VALIDACIÓN DE COSTOS", "PRODUCCIÓN", "AUDITORÍA DE CUENTAS", "PEDIDOS CAFETERÍA"]
    
    # Lógica de asignación de módulos mejorada
    if st.session_state.rol_actual == "ADMIN":
        # El Administrador tiene acceso total por defecto
        opciones_menu = todos_los_modulos + ["CONFIGURACIÓN"]
    else:
        # Los demás usuarios se filtran estrictamente por sus permisos en la base
        opciones_menu = [mod for mod in todos_los_modulos if mod in st.session_state.get("modulos_permitidos", [])]
            
    if "menu_opcion" not in st.session_state:
        st.session_state.menu_opcion = st.query_params.get("modulo", opciones_menu[0] if opciones_menu else "")
        
    if st.session_state.menu_opcion not in opciones_menu and opciones_menu:
        st.session_state.menu_opcion = opciones_menu[0]
        
    def actualizar_url_modulo():
        st.query_params["modulo"] = st.session_state.menu_radio
        st.session_state.menu_opcion = st.session_state.menu_radio
        
    if opciones_menu:
        opcion = st.radio(
            "Módulos disponibles:", 
            opciones_menu,
            index=opciones_menu.index(st.session_state.menu_opcion) if st.session_state.menu_opcion in opciones_menu else 0,
            key="menu_radio",
            on_change=actualizar_url_modulo
        )
    else:
        st.warning("No tienes módulos asignados.")
        opcion = None
    
    st.divider()
    st.markdown(f"👤 Usuario: **{st.session_state.usuario_actual}**")
    st.markdown(f"📍 Anexo: **{st.session_state.get('anexo_actual', 'No asignado')}**")
    st.markdown(f"🛡️ Rol: **{st.session_state.rol_actual}**")
    
    if st.button("Cerrar Sesión", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.usuario_actual = ""
        st.session_state.rol_actual = ""
        st.query_params.clear() 
        if "menu_opcion" in st.session_state:
            del st.session_state["menu_opcion"]
        st.rerun()