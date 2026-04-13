import streamlit as st
import urllib.parse
from datetime import datetime, timedelta
import auth, gastos, ventas, costos, validacion  # <--- Importamos el nuevo módulo
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
        if st.session_state.last_activity is None:
            st.session_state.last_activity = datetime.now()

if st.session_state.logged_in:
    tiempo_inactivo = datetime.now() - st.session_state.last_activity
    if tiempo_inactivo > timedelta(hours=1):
        st.session_state.logged_in = False
        st.query_params.clear() 
        st.session_state.usuario_actual = ""
        st.session_state.rol_actual = ""
        st.session_state.mensaje_login = "⏳ Tu sesión ha expirado. Por favor, inicia sesión nuevamente."
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

            if btn_ingresar:
                usuario_limpio = usuario.strip().upper()
                if usuario_limpio in usuarios_validos and contrasena.strip() == usuarios_validos[usuario_limpio]["Contrasena"]:
                    st.session_state.logged_in = True
                    st.session_state.usuario_actual = usuario_limpio
                    st.session_state.rol_actual = usuarios_validos[usuario_limpio].get("Rol", "USUARIO")
                    st.session_state.last_activity = datetime.now()
                    st.query_params["auth"] = urllib.parse.quote(usuario_limpio)
                    st.rerun()
                else:
                    st.error("❌ Usuario o contraseña incorrectos.")
    st.stop() 

# --- Sidebar y Menú ---
with st.sidebar:
    st.title("Menú Principal")
    # Agregamos la opción de VALIDACIÓN DE COSTOS al menú
    opciones_menu = ["KPI DE REGISTROS", "KPI DE VENTAS", "CONTABILIDAD DE COSTOS", "VALIDACIÓN DE COSTOS"]
    
    if st.session_state.rol_actual == "ADMIN":
        opciones_menu.append("CONFIGURACIÓN")
        
    if "menu_opcion" not in st.session_state:
        st.session_state.menu_opcion = st.query_params.get("modulo", opciones_menu[0])
        
    if st.session_state.menu_opcion not in opciones_menu:
        st.session_state.menu_opcion = opciones_menu[0]
        
    def actualizar_url_modulo():
        st.query_params["modulo"] = st.session_state.menu_radio
        st.session_state.menu_opcion = st.session_state.menu_radio
        
    opcion = st.radio(
        "Módulos disponibles:", 
        opciones_menu,
        index=opciones_menu.index(st.session_state.menu_opcion),
        key="menu_radio",
        on_change=actualizar_url_modulo
    )
    
    st.divider()
    st.markdown(f"👤 Usuario: **{st.session_state.usuario_actual}**")
    st.markdown(f"🛡️ Rol: **{st.session_state.rol_actual}**")
    
    if st.button("Cerrar Sesión", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.usuario_actual = ""
        st.session_state.rol_actual = ""
        st.query_params.clear() 
        if "menu_opcion" in st.session_state:
            del st.session_state["menu_opcion"]
        st.rerun()

# --- Lógica de Navegación ---
if opcion == "KPI DE REGISTROS":
    gastos.mostrar_modulo_gastos()
elif opcion == "KPI DE VENTAS":
    ventas.mostrar_modulo_ventas()
elif opcion == "CONTABILIDAD DE COSTOS":
    costos.mostrar_modulo_costos()
elif opcion == "VALIDACIÓN DE COSTOS":
    validacion.mostrar_modulo_validacion()  # <--- Llamada al nuevo módulo independiente
elif opcion == "CONFIGURACIÓN":
    st.title("⚙️ Configuración del Sistema")
    st.markdown("Desde aquí puedes administrar los accesos a la plataforma de forma rápida.")
    
    st.subheader("👥 Gestión de Usuarios")
    with st.expander("➕ Agregar nuevo usuario", expanded=False):
        with st.form("form_nuevo_usuario"):
            st.write("Completa los datos para registrar a una nueva persona:")
            col_n1, col_n2 = st.columns(2)
            with col_n1: nuevo_nombre = st.text_input("Nombre de Usuario").upper()
            with col_n2: nuevo_rol = st.selectbox("Nivel de Acceso (Rol)", ["USUARIO", "ADMIN"])
            nueva_pass = st.text_input("Contraseña")
            btn_crear = st.form_submit_button("Guardar Usuario")
            
            if btn_crear:
                if nuevo_nombre and nueva_pass:
                    if nuevo_nombre in usuarios_validos:
                        st.error("⚠️ Ese nombre de usuario ya existe en el sistema.")
                    else:
                        with st.spinner("Guardando en la base de datos..."):
                            ws_usu = conectar_hoja("Usuarios")
                            if ws_usu:
                                try:
                                    ws_usu.append_row([nuevo_nombre, nueva_pass, nuevo_rol])
                                    auth.obtener_usuarios_db.clear()
                                    st.success(f"✅ Usuario '{nuevo_nombre}' creado con éxito.")
                                except Exception as e:
                                    st.error(f"Error al guardar: {e}")
                else:
                    st.warning("⚠️ Debes completar el nombre y la contraseña.")

    st.write("---")
    st.write("**Lista de usuarios activos:**")
    df_usuarios_ver = obtener_dataframe("Usuarios")
    if not df_usuarios_ver.empty:
        st.dataframe(df_usuarios_ver[['Usuario', 'Rol']], hide_index=True, use_container_width=True)