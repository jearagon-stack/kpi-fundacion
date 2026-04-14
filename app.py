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
            
            # --- NUEVA SECCIÓN: PARÁMETROS DE AUDITORÍA ---
        st.divider()
        st.subheader("⚙️ Configuración de Parámetros de Auditoría")
        
        # Intentamos leer si ya hay datos guardados
        try:
            df_params = obtener_dataframe("Parametros_Auditoria")
        except:
            df_params = pd.DataFrame() # Si no existe la hoja, no da error
        
        # Valores por defecto
        def_var, def_limp, def_emp, def_mat = 0.01, 3000.0, 5000.0, 10000.0
        
        # Si hay datos en la hoja, los usamos para llenar las cajitas
        if not df_params.empty:
            try:
                def_var = float(df_params[df_params['Criterio'] == 'VARIACION_MAX_PERMITIDA']['Valor_Tope'].iloc[0])
                def_limp = float(df_params[df_params['Criterio'] == 'LIMPIEZA']['Valor_Tope'].iloc[0])
                def_emp = float(df_params[df_params['Criterio'] == 'EMPAQUE']['Valor_Tope'].iloc[0])
                def_mat = float(df_params[df_params['Criterio'] == 'MATERIA_PRIMA']['Valor_Tope'].iloc[0])
            except:
                pass # Si hay algún error leyendo, usa los valores por defecto

        with st.form("form_parametros_auditoria"):
            c1, c2 = st.columns(2)
            new_var = c1.number_input("Variación Máx. Costo (Ej: 0.01 para 1%)", value=def_var, format="%.4f")
            new_limp = c2.number_input("Tope Inventario Limpieza ($)", value=def_limp)
            new_emp = c1.number_input("Tope Inventario Empaque ($)", value=def_emp)
            new_mat = c2.number_input("Tope Inventario Materia Prima ($)", value=def_mat)
            
            if st.form_submit_button("💾 Guardar Configuración de Auditoría"):
                ws_p = conectar_hoja("Parametros_Auditoria")
                ws_p.clear() 
                headers = ["Variable", "Criterio", "Valor_Tope", "Descripcion"]
                filas = [
                    headers,
                    ["AUDIT_COSTO", "VARIACION_MAX_PERMITIDA", new_var, "Variación costo inicial vs final"],
                    ["TOPE_CATEGORIA", "LIMPIEZA", new_limp, "Límite inversión limpieza"],
                    ["TOPE_CATEGORIA", "EMPAQUE", new_emp, "Límite inversión empaque"],
                    ["TOPE_CATEGORIA", "MATERIA_PRIMA", new_mat, "Límite inversión materia prima"]
                ]
                ws_p.update("A1", filas)
                st.cache_data.clear()
                st.success("✅ Parámetros de auditoría actualizados correctamente.")