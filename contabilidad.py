import streamlit as st
import pandas as pd
import plotly.express as px
from utils import obtener_dataframe

# --- FUNCIONES AUXILIARES ---
def limpiar_texto(texto):
    return str(texto).strip().upper()

def buscar_columna(df, palabras_clave):
    """Busca una columna en el DataFrame basada en palabras clave."""
    for col in df.columns:
        if any(p in str(col).upper() for p in palabras_clave):
            return col
    return None

def mostrar_modulo_contabilidad():
    st.title("📊 Análisis Financiero y Punto de Equilibrio")
    
    tab_carga, tab_pe, tab_dash = st.tabs([
        "📥 1. Carga y Sincronización", 
        "⚖️ 2. Punto de Equilibrio", 
        "📊 3. Dashboard Analítico"
    ])

    # ==========================================
    # PESTAÑA 1: CARGA DE ARCHIVOS
    # ==========================================
    with tab_carga:
        st.subheader("Configuración del Periodo y Unidad")
        
        # 1. Periodo Desde / Hasta
        col1, col2, col3, col4 = st.columns(4)
        meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        with col1:
            mes_desde = st.selectbox("Mes Desde:", meses, index=0)
        with col2:
            anio_desde = st.number_input("Año Desde:", min_value=2024, max_value=2030, value=2025)
        with col3:
            mes_hasta = st.selectbox("Mes Hasta:", meses, index=11)
        with col4:
            anio_hasta = st.number_input("Año Hasta:", min_value=2024, max_value=2030, value=2025)
            
        # 2. Selector de Unidades
        st.markdown("---")
        unidades_seleccionadas = st.multiselect(
            "Selecciona la(s) unidad(es) a analizar (Consolidado):", 
            ["Fundación Global", "Cafetería", "Talleres", "Librería", "Proyectos Especiales", "Otras"],
            default=["Fundación Global"]
        )
        st.session_state['unidades_analisis'] = unidades_seleccionadas

        st.markdown("---")
        archivos_subidos = st.file_uploader("Sube los archivos Excel de las unidades seleccionadas:", type=["xlsx", "xls"], accept_multiple_files=True)

        if st.button("🔄 Sincronizar con Diccionario y Calcular", type="primary", use_container_width=True):
            if not archivos_subidos:
                st.warning("⚠️ Debes subir al menos un archivo de Excel para proceder.")
            elif not unidades_seleccionadas:
                st.warning("⚠️ Debes seleccionar al menos una unidad de negocio.")
            else:
                with st.spinner("Procesando datos y sincronizando con Google Sheets..."):
                    try:
                        # 1. Leer Diccionario Maestro
                        df_map = obtener_dataframe("Balance_Mapeado")
                        if df_map is None:
                            st.error("Error: No se pudo conectar a la hoja 'Balance_Mapeado'.")
                            st.stop()
                        
                        df_map.columns = df_map.columns.str.strip().str.upper()
                        
                        # Columnas esperadas
                        c_map_cta = buscar_columna(df_map, ["CUENTA", "ID"])
                        c_map_tipo = buscar_columna(df_map, ["TIPO"])
                        c_map_est = buscar_columna(df_map, ["ESTADO"])
                        c_map_cat = buscar_columna(df_map, ["CATEGOR"])
                        
                        # 2. Procesar Archivos Subidos
                        dfs = []
                        for arch in archivos_subidos:
                            df_temp = pd.read_excel(arch, dtype=str)
                            df_temp.columns = df_temp.columns.str.strip().str.upper()
                            dfs.append(df_temp)
                        
                        df_cons = pd.concat(dfs, ignore_index=True)
                        
                        # Búsqueda dinámica de columnas en los archivos subidos
                        c_arch_cta = buscar_columna(df_cons, ["CUENTA", "ID"])
                        c_arch_sld = buscar_columna(df_cons, ["SALDO", "FINAL"])

                        # Limpieza numérica
                        df_cons[c_arch_sld] = pd.to_numeric(df_cons[c_arch_sld].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                        df_agrupado = df_cons.groupby(c_arch_cta, as_index=False)[c_arch_sld].sum()

                        # 3. Cruce Maestro (Inner Join)
                        df_final = pd.merge(df_agrupado, df_map, left_on=c_arch_cta, right_on=c_map_cta, how="inner")
                        
                        # Limpiar tipo de dato y descartar lo que no aplica
                        df_final[c_map_tipo] = df_final[c_map_tipo].apply(limpiar_texto)
                        df_final = df_final[~df_final[c_map_tipo].isin(["NO APLICA", "CUENTA DE MAYOR", ""])]
                        
                        st.session_state['cont_df'] = df_final
                        st.session_state['c_tipo'] = c_map_tipo
                        st.session_state['c_est'] = c_map_est
                        st.session_state['c_cat'] = c_map_cat
                        st.session_state['c_sld'] = c_arch_sld
                        
                        st.success(f"✅ ¡Datos procesados! Se analizarán las cuentas para: {', '.join(unidades_seleccionadas)}")
                    except Exception as e:
                        st.error(f"Error técnico en el proceso: {e}")

    # ==========================================
    # PESTAÑA 2: PUNTO DE EQUILIBRIO
    # ==========================================
    with tab_pe:
        if 'cont_df' in st.session_state:
            df = st.session_state['cont_df']
            c_t = st.session_state['c_tipo']
            c_s = st.session_state['c_sld']
            unidades = st.session_state.get('unidades_analisis', [])
            
            st.subheader(f"⚖️ Punto de Equilibrio: {', '.join(unidades)}")
            
            # Cálculo usando expresiones regulares para detectar "Ingresos", "Ventas", "Fijos", "Variables"
            ventas = df[df[c_t].str.contains("INGRESO|VENTA", na=False)][c_s].abs().sum()
            cf = df[df[c_t].str.contains("FIJO", na=False)][c_s].abs().sum()
            cv = df[df[c_t].str.contains("VARIABLE", na=False)][c_s].abs().sum()
            
            # Fórmulas
            margen_pct = (1 - (cv / ventas)) if ventas > 0 else 0
            pe = (cf / margen_pct) if margen_pct > 0 else 0
            
            c1, c2, c3 = st.columns(3)
            c1.metric("Ingresos Totales", f"${ventas:,.2f}")
            c2.metric("Costos Variables (CV)", f"${cv:,.2f}")
            c3.metric("Costos Fijos (CF)", f"${cf:,.2f}")
            
            st.divider()
            
            col_margen, col_pe = st.columns(2)
            with col_margen:
                st.info("### Margen de Contribución")
                st.markdown(f"<h1 style='text-align: center; color: #4CAF50;'>{margen_pct * 100:.2f}%</h1>", unsafe_allow_html=True)
                st.caption("Porcentaje de cada dólar que queda para cubrir costos fijos.")
                
            with col_pe:
                if margen_pct > 0:
                    st.success("### Punto de Equilibrio ($)")
                    st.markdown(f"<h1 style='text-align: center;'>${pe:,.2f}</h1>", unsafe_allow_html=True)
                    st.caption("Monto de ingresos necesario para no tener pérdidas.")
                else:
                    st.error("### Punto de Equilibrio ($)")
                    st.markdown("<h3 style='text-align: center;'>Incalculable</h3>", unsafe_allow_html=True)
                    st.caption("Los costos variables superan a los ingresos, o no hay ingresos registrados.")
        else:
            st.info("👈 Por favor, realiza la sincronización en la Pestaña 1.")

    # ==========================================
    # PESTAÑA 3: DASHBOARD ANALÍTICO
    # ==========================================
    with tab_dash:
        if 'cont_df' in st.session_state:
            df = st.session_state['cont_df']
            c_t = st.session_state['c_tipo']
            c_est = st.session_state['c_est']
            c_cat = st.session_state['c_cat']
            c_s = st.session_state['c_sld']
            unidades = st.session_state.get('unidades_analisis', [])
            
            st.subheader(f"📊 Desglose Analítico: {', '.join(unidades)}")
            
            df_gastos = df[df[c_t].str.contains("FIJO|VARIABLE", na=False)]
            
            if not df_gastos.empty:
                col_g1, col_g2 = st.columns(2)
                
                with col_g1:
                    st.markdown("**Proporción de Gastos (Fijo vs Variable)**")
                    fig_tipo = px.pie(df_gastos, values=c_s, names=c_t, hole=0.4)
                    st.plotly_chart(fig_tipo, use_container_width=True)
                    
                with col_g2:
                    st.markdown("**Top Gastos por Categoría**")
                    df_cat = df_gastos.groupby(c_cat, as_index=False)[c_s].sum().sort_values(by=c_s, ascending=False)
                    fig_cat = px.bar(df_cat, x=c_cat, y=c_s, text_auto='.2s')
                    st.plotly_chart(fig_cat, use_container_width=True)
                    
                st.divider()
                st.markdown("**Análisis Profundo por Estado de Cuenta**")
                
                # Usar selectbox en lugar de radio para evitar errores si no hay algún tipo
                tipos_existentes = df_gastos[c_t].unique().tolist()
                tipo_filtro = st.selectbox("Filtrar desglose por:", tipos_existentes)
                
                df_filtro = df_gastos[df_gastos[c_t] == tipo_filtro]
                if not df_filtro.empty:
                    df_estado = df_filtro.groupby(c_est, as_index=False)[c_s].sum().sort_values(by=c_s, ascending=False)
                    fig_est = px.bar(df_estado, y=c_est, x=c_s, orientation='h', text_auto='.2s')
                    st.plotly_chart(fig_est, use_container_width=True)
            else:
                st.write("No hay gastos registrados en el periodo para analizar.")
        else:
            st.info("👈 Por favor, realiza la sincronización en la Pestaña 1.")