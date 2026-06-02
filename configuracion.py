import streamlit as st
import pandas as pd
from utils import conectar_hoja, obtener_dataframe

def mostrar_modulo_configuracion():
    st.title("⚙️ Configuración del Sistema")
    st.markdown("Administración de accesos y parámetros de la plataforma.")
    
    st.subheader("👥 Gestión de Usuarios")
    with st.expander("➕ Agregar nuevo usuario", expanded=True):
        with st.form("form_nuevo_usuario"):
            st.write("Datos para registrar a una nueva persona:")
            col_n1, col_n2 = st.columns(2)
            with col_n1: 
                nuevo_nombre = st.text_input("Nombre de Usuario").upper()
                nueva_pass = st.text_input("Contraseña", type="password")
            with col_n2: 
                nuevo_rol = st.selectbox("Nivel de Acceso (Rol)", ["CAJERA", "BODEGUERO", "USUARIO", "ADMIN"])
                nuevo_anexo = st.text_input("Anexo (Ej: Cafeteria ICAS)", value="General")
                
            modulos_seleccionados = st.multiselect(
                "Módulos Permitidos",
                ["KPI DE REGISTROS", "KPI DE VENTAS", "CONTABILIDAD DE COSTOS", "VALIDACIÓN DE COSTOS", "PRODUCCIÓN", "AUDITORÍA DE CUENTAS", "PEDIDOS CAFETERÍA"],
                default=["PEDIDOS CAFETERÍA"]
            )
            
            btn_crear = st.form_submit_button("💾 Guardar Usuario")
            
            if btn_crear:
                if nuevo_nombre == "" or nueva_pass == "":
                    st.warning("⚠️ El nombre de usuario y la contraseña son obligatorios.")
                else:
                    try:
                        # OJO: Asumo que tu pestaña de accesos se llama "Usuarios". 
                        # Si se llama diferente, cambia el nombre aquí abajo.
                        ws_usuarios = conectar_hoja("Usuarios") 
                        
                        # Convertimos la lista de módulos en texto separado por comas
                        modulos_str = ", ".join(modulos_seleccionados)
                        
                        # Creamos la fila que se insertará en Google Sheets
                        nueva_fila = [nuevo_nombre, nueva_pass, nuevo_rol, nuevo_anexo, modulos_str]
                        
                        # Guardamos en la nube
                        ws_usuarios.append_row(nueva_fila)
                        st.cache_data.clear() # Limpiamos caché para que el sistema reconozca al nuevo usuario de inmediato
                        
                        st.success(f"✅ ¡Usuario '{nuevo_nombre}' creado y guardado en la base de datos con éxito!")
                    except Exception as e:
                        st.error(f"❌ Error al conectar con Google Sheets: {e}")
                        st.info("Verifica que la pestaña se llame 'Usuarios' en tu Google Sheets.")
                
    st.divider()
    st.subheader("⚙️ Configuración de Parámetros de Auditoría")
    
    try:
        df_params = obtener_dataframe("Parametros_Auditoria")
    except Exception:
        df_params = pd.DataFrame() 
    
    def_var, def_limp, def_emp, def_mat, def_prod = 0.01, 3000.0, 5000.0, 10000.0, 8000.0
    
    if not df_params.empty:
        try:
            def_var = float(df_params[df_params['Criterio'] == 'VARIACION_MAX_PERMITIDA']['Valor_Tope'].iloc[0])
            def_limp = float(df_params[df_params['Criterio'] == 'LIMPIEZA']['Valor_Tope'].iloc[0])
            def_emp = float(df_params[df_params['Criterio'] == 'EMPAQUE']['Valor_Tope'].iloc[0])
            def_mat = float(df_params[df_params['Criterio'] == 'MATERIA_PRIMA']['Valor_Tope'].iloc[0])
            val_prod = df_params[df_params['Criterio'].isin(['PRODUCTO_TERMINADO', 'PRODUCTO TERMINADO'])]['Valor_Tope']
            if not val_prod.empty: def_prod = float(val_prod.iloc[0])
        except Exception:
            pass 

    with st.form("form_parametros_auditoria"):
        c1, c2 = st.columns(2)
        new_var = c1.number_input("Variación Máx. Costo (Ej: 0.01 para 1%)", value=def_var, format="%.4f")
        new_limp = c2.number_input("Tope Inventario Limpieza ($)", value=def_limp)
        new_emp = c1.number_input("Tope Inventario Empaque ($)", value=def_emp)
        new_mat = c2.number_input("Tope Inventario Materia Prima ($)", value=def_mat)
        new_prod = c1.number_input("Tope Inventario Producto Terminado ($)", value=def_prod)
        
        if st.form_submit_button("💾 Guardar Configuración de Auditoría"):
            ws_p = conectar_hoja("Parametros_Auditoria")
            ws_p.clear() 
            headers = ["Variable", "Criterio", "Valor_Tope", "Descripcion"]
            filas = [
                headers,
                ["AUDIT_COSTO", "VARIACION_MAX_PERMITIDA", new_var, "Variación costo inicial vs final"],
                ["TOPE_CATEGORIA", "LIMPIEZA", new_limp, "Límite inversión limpieza"],
                ["TOPE_CATEGORIA", "EMPAQUE", new_emp, "Límite inversión empaque"],
                ["TOPE_CATEGORIA", "MATERIA_PRIMA", new_mat, "Límite inversión materia prima"],
                ["TOPE_CATEGORIA", "PRODUCTO_TERMINADO", new_prod, "Límite inversión producto terminado"]
            ]
            ws_p.update("A1", filas)
            st.cache_data.clear()
            st.success("✅ Parámetros de auditoría actualizados correctamente.")