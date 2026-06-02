import streamlit as st
import urllib.parse
import pandas as pd
from datetime import datetime, timedelta
import auth, gastos, ventas, costos, validacion  
from utils import conectar_hoja, obtener_dataframe

st.set_page_config(page_title="Auditoría DTE Pro", layout="wide")

# --- INYECCIÓN DE DISEÑO (CSS) ---
estilo_personalizado = """
    <style>
    /* Importar fuente de Google Fonts (Nunito) */
    @import url('https://fonts.googleapis.com/css2?family=Nunito:wght@300;400;600;700&display=swap');

    /* Aplicar la fuente a la interfaz */
    html, body, [class*="css"], [class*="st-"] {
        font-family: 'Nunito', sans-serif !important;
    }

    /* Fondo animado (Gradiente en movimiento) */
    .stApp {
        background: linear-gradient(-45deg, #f0f2f6, #e2e8f0, #cbd5e1, #f8fafc);
        background-size: 400% 400%;
        animation: moverFondo 15s ease infinite;
    }

    @keyframes moverFondo {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }

    /* Estilización de botones */
    .stButton>button {
        border-radius: 20px !important;
        font-weight: 600 !important;
        transition: all 0.3s ease;
    }
    
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
    }

    /* Suavizar los bordes de los formularios */
    div[data-testid="stForm"] {
        border-radius: 15px !important;
        border: 1px solid rgba(0, 0, 0, 0.05) !important;
        padding: 20px !important;
        background-color: rgba(255, 255, 255, 0.8) !important; /* Fondo semitransparente para resaltar sobre el fondo animado */
    }

    /* Estilizar las pestañas */
    button[data-baseweb="tab"] {
        font-family: 'Nunito', sans-serif !important;
        font-weight: 600 !important;
    }

    /* Ocultar elementos predeterminados de Streamlit */
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
"""
st.markdown(estilo_personalizado, unsafe_allow_html=True)
# ---------------------------------

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

            if btn_ingresar:
                usuario_limpio = usuario.strip().upper()
                if usuario_limpio in usuarios_validos and contrasena.strip() == usuarios_validos[usuario_limpio]["Contrasena"]:
                    st.session_state.logged_in = True
                    st.session_state.usuario_actual = usuario_limpio
                    st.session_state.rol_actual = usuarios_validos[usuario_limpio].get("Rol", "USUARIO")
                    st.session_state.anexo_actual = usuarios_validos[usuario_limpio].get("Anexo", "General")
                    
                    modulos_permitidos = usuarios_validos[usuario_limpio].get("Modulos", "PEDIDOS CAFETERÍA")
                    st.session_state.modulos_permitidos = [m.strip() for m in modulos_permitidos.split(",")]
                    
                    st.session_state.last_activity = datetime.now()
                    st.query_params["auth"] = urllib.parse.quote(usuario_limpio)
                    st.rerun()
                else:
                    st.error("Usuario o contraseña incorrectos.")
    st.stop() 

# --- Sidebar y Menú ---
with st.sidebar:
    st.title("Menú Principal")
    
    todos_los_modulos = ["KPI DE REGISTROS", "KPI DE VENTAS", "CONTABILIDAD DE COSTOS", "VALIDACIÓN DE COSTOS", "PRODUCCIÓN", "AUDITORÍA DE CUENTAS", "PEDIDOS CAFETERÍA"]
    
    if st.session_state.rol_actual == "ADMIN":
        opciones_menu = todos_los_modulos + ["CONFIGURACIÓN"]
    else:
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

# --- Lógica de Navegación ---
if opcion == "KPI DE REGISTROS":
    gastos.mostrar_modulo_gastos()
elif opcion == "KPI DE VENTAS":
    ventas.mostrar_modulo_ventas()
elif opcion == "CONTABILIDAD DE COSTOS":
    costos.mostrar_modulo_costos()
elif opcion == "VALIDACIÓN DE COSTOS":
    validacion.mostrar_modulo_validacion()  
elif opcion == "PRODUCCIÓN":
    try:
        from costs_produccion import mostrar_modulo_produccion
        mostrar_modulo_produccion()
    except ImportError:
        st.warning("El archivo 'costs_produccion.py' aún no ha sido creado.")
elif opcion == "AUDITORÍA DE CUENTAS":
    try:
        from audit_cuentas import mostrar_modulo_auditoria
        mostrar_modulo_auditoria()
    except ImportError:
        st.warning("El archivo 'audit_cuentas.py' aún no ha sido creado.")
elif opcion == "PEDIDOS CAFETERÍA":
    try:
        from pedidos_cafeteria import mostrar_modulo_pedidos
        mostrar_modulo_pedidos()
    except ImportError:
        st.warning("El archivo 'pedidos_cafeteria.py' aún no ha sido creado.")
elif opcion == "CONFIGURACIÓN":
    try:
        from configuracion import mostrar_modulo_configuracion
        mostrar_modulo_configuracion()
    except ImportError:
        st.warning("El archivo 'configuracion.py' aún no ha sido creado.")