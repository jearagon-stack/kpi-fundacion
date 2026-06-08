import streamlit as st
st.write("¡El archivo contabilidad.py SÍ está siendo detectado!")
import streamlit as st
import pandas as pd
import plotly.express as px
from utils import obtener_dataframe

# --- FUNCIONES AUXILIARES ---
def limpiar_texto(texto):
    return str(texto).strip().upper()

def buscar_columna(df, palabras_clave):
    """Busca una columna en el DataFrame basándose en palabras clave."""
    for col in df.columns:
        if any(p in str(col).upper() for p in palabras_clave):
            return col
    return None

# --- MÓDULO PRINCIPAL ---
def mostrar_modulo_contabilidad():
    st.title("📊 Análisis Financiero y Punto de Equilibrio")
    st.info("Módulo de consolidación automática basado en el diccionario contable.")

    tab_carga, tab_pe, tab_dash = st.tabs([
        "📥 1. Carga y Sincronización", 
        "⚖️ 2. Punto de Equilibrio", 
        "📊 3. Dashboard Analítico"
    ])

    # ==========================================
    # PESTAÑA 1: CARGA DE ARCHIVOS Y DICCIONARIO
    # ==========================================
    with tab_carga:
        st.subheader("Configuración del Análisis")
        
        col1, col2 = st.columns(2)
        with col1:
            mes = st.selectbox("Mes de análisis:", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"])
        with col2:
            anio = st.number_input("Año:", min_value=2024, max_value=2030, value=2026)
            
        st.markdown("---")
        st.write("**Archivos de Balances Mensuales (Puedes subir 1 o varios para consolidar):**")
        archivos_subidos = st.file_uploader("Selecciona los archivos Excel de las unidades", type=["xlsx", "xls"], accept_multiple_files=True)

        if st.button("🔄 Sincronizar con Diccionario y Calcular", type="primary", use_container_width=True):
            if not archivos_subidos:
                st.warning("⚠️ Debes subir al menos un archivo de Excel para analizar.")
            else:
                with st.spinner("Conectando con Google Sheets y cruzando datos..."):
                    try:
                        # 1. Leer Diccionario de Google Sheets
                        df_map = obtener_dataframe("Balance_Mapeado")
                        if df_map is None or df_map.empty:
                            st.error("❌ No se pudo leer la pestaña 'Balance_Mapeado' o está vacía.")
                            st.stop()
                            
                        # Limpiar nombres de columnas del diccionario
                        df_map.columns = df_map.columns.str.strip().str.upper()
                        
                        col_map_cuenta = buscar_columna(df_map, ["CUENTA", "ID"])
                        col_map_tipo = buscar_columna(df_map, ["TIPO"])
                        col_map_estado = buscar_columna(df_map, ["ESTADO"])
                        col_map_cat = buscar_columna(df_map, ["CATEGOR"])
                        
                        if not all([col_map_cuenta, col_map_tipo, col_map_estado, col_map_cat]):
                            st.error("❌ El diccionario en Google Sheets no tiene las columnas requeridas (Cuenta, Tipo, Estado, Categoria).")
                            st.stop()

                        # Filtrar solo cuentas válidas en el diccionario
                        tipos_validos = ["COSTO FIJO", "COSTO VARIABLE", "VENTAS"]
                        df_map[col_map_tipo] = df_map[col_map_tipo].apply(limpiar_texto)
                        df_map_valido = df_map[df_map[col_map_tipo].isin(tipos_validos)].copy()
                        df_map_valido[col_map_cuenta] = df_map_valido[col_map_cuenta].astype(str).str.strip()

                        # 2. Leer y Consolidar Archivos Subidos
                        dfs_archivos = []
                        for arch in archivos_subidos:
                            df_temp = pd.read_excel(arch, dtype=str)
                            df_temp.columns = df_temp.columns.str.strip().str.upper()
                            dfs_archivos.append(df_temp)
                            
                        df_consolidado = pd.concat(dfs_archivos, ignore_index=True)

                        # Encontrar columnas clave en los archivos subidos
                        col_arch_cuenta = buscar_columna(df_consolidado, ["CUENTA", "ID"])
                        col_arch_saldo = buscar_columna(df_consolidado, ["SALDO", "FINAL"])

                        if not col_arch_cuenta or not col_arch_saldo:
                            st.error("❌ No se encontraron las columnas 'Cuenta' o 'Saldo' en los archivos de Excel.")
                            st.stop()

                        # Limpiar datos consolidados
                        df_consolidado[col_arch_cuenta] = df_consolidado[col_arch_cuenta].astype(str).str.strip()
                        df_consolidado[col_arch_saldo] = pd.to_numeric(df_consolidado[col_arch_saldo].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                        
                        # Sumar saldos si hay cuentas repetidas (por consolidación)
                        df_agrupado = df_consolidado.groupby(col_arch_cuenta, as_index=False)[col_arch_saldo].sum()

                        # 3. Cruce de Datos (Match)
                        df_final = pd.merge(
                            df_agrupado, 
                            df_map_valido[[col_map_cuenta, col_map_tipo, col_map_estado, col_map_cat]], 
                            left_on=col_arch_cuenta, 
                            right_on=col_map_cuenta, 
                            how="inner"
                        )

                        # Trabajar con valores absolutos (para evitar problemas con cuentas de naturaleza acreedora/deudora)
                        df_final["Saldo_Absoluto"] = df_final[col_arch_saldo].abs()

                        # Guardar en memoria
                        st.session_state['cont_df_final'] = df_final
                        st.session_state['cont_col_tipo'] = col_map_tipo
                        st.session_state['cont_col_estado'] = col_map_estado
                        st.session_state['cont_col_cat'] = col_map_cat
                        
                        st.success(f"✅ ¡Datos procesados con éxito! Se consolidaron {len(archivos_subidos)} archivo(s). Avanza a la pestaña 2.")

                    except Exception as e:
                        st.error(f"Ocurrió un error en el procesamiento: {e}")

    # ==========================================
    # PESTAÑA 2: PUNTO DE EQUILIBRIO
    # ==========================================
    with tab_pe:
        st.subheader("⚖️ Cálculo de Punto de Equilibrio Financiero")
        
        if 'cont_df_final' in st.session_state:
            df = st.session_state['cont_df_final']
            c_tipo = st.session_state['cont_col_tipo']
            
            # Separar por tipos
            ventas = df[df[c_tipo] == "VENTAS"]["Saldo_Absoluto"].sum()
            costos_fijos = df[df[c_tipo] == "COSTO FIJO"]["Saldo_Absoluto"].sum()
            costos_variables = df[df[c_tipo] == "COSTO VARIABLE"]["Saldo_Absoluto"].sum()
            
            # Cálculos Matemáticos
            margen_contribucion_dlrs = ventas - costos_variables
            margen_contribucion_pct = (margen_contribucion_dlrs / ventas) if ventas > 0 else 0
            
            punto_equilibrio = (costos_fijos / margen_contribucion_pct) if margen_contribucion_pct > 0 else 0
            
            # Visualización (Tarjetas)
            col_v, col_cv, col_cf = st.columns(3)
            col_v.metric("Ventas Totales", f"${ventas:,.2f}")
            col_cv.metric("Costos Variables (CV)", f"${costos_variables:,.2f}")
            col_cf.metric("Costos Fijos (CF)", f"${costos_fijos:,.2f}")
            
            st.divider()
            
            col_margen, col_pe = st.columns(2)
            with col_margen:
                st.info("### Margen de Contribución")
                st.markdown(f"<h1 style='text-align: center; color: #4CAF50;'>{margen_contribucion_pct * 100:.2f}%</h1>", unsafe_allow_html=True)
                st.caption("Porcentaje de cada dólar de venta que queda para cubrir costos fijos.")
                
            with col_pe:
                if margen_contribucion_pct > 0:
                    st.success("### Punto de Equilibrio ($)")
                    st.markdown(f"<h1 style='text-align: center;'>${punto_equilibrio:,.2f}</h1>", unsafe_allow_html=True)
                    st.caption("Monto de ventas necesario para no tener pérdidas ni ganancias.")
                else:
                    st.error("### Punto de Equilibrio ($)")
                    st.markdown("<h3 style='text-align: center;'>Incalculable</h3>", unsafe_allow_html=True)
                    st.caption("Los costos variables superan o igualan a las ventas.")

            # Indicador visual simple de alcance
            if ventas > 0 and punto_equilibrio > 0:
                progreso = min(ventas / punto_equilibrio, 1.0)
                st.progress(progreso)
                if ventas >= punto_equilibrio:
                    st.write(f"🎉 ¡Meta superada! Las ventas están **${(ventas - punto_equilibrio):,.2f}** por encima del punto de equilibrio.")
                else:
                    st.write(f"⚠️ Faltan **${(punto_equilibrio - ventas):,.2f}** en ventas para alcanzar el punto de equilibrio.")
                    
        else:
            st.info("👈 Por favor, carga los archivos en la Pestaña 1 primero.")

    # ==========================================
    # PESTAÑA 3: DASHBOARD ANALÍTICO
    # ==========================================
    with tab_dash:
        st.subheader("📊 Desglose de Gastos y Categorías")
        
        if 'cont_df_final' in st.session_state:
            df = st.session_state['cont_df_final']
            c_tipo = st.session_state['cont_col_tipo']
            c_est = st.session_state['cont_col_estado']
            c_cat = st.session_state['cont_col_cat']
            
            df_gastos = df[df[c_tipo].isin(["COSTO FIJO", "COSTO VARIABLE"])]
            
            if not df_gastos.empty:
                col_g1, col_g2 = st.columns(2)
                
                with col_g1:
                    st.markdown("**Gastos por Tipo (Fijo vs Variable)**")
                    fig_tipo = px.pie(df_gastos, values='Saldo_Absoluto', names=c_tipo, hole=0.4)
                    st.plotly_chart(fig_tipo, use_container_width=True)
                    
                with col_g2:
                    st.markdown("**Top Gastos por Categoría (General)**")
                    df_cat = df_gastos.groupby(c_cat, as_index=False)['Saldo_Absoluto'].sum().sort_values(by='Saldo_Absoluto', ascending=False)
                    fig_cat = px.bar(df_cat, x=c_cat, y='Saldo_Absoluto', text_auto='.2s')
                    st.plotly_chart(fig_cat, use_container_width=True)
                    
                st.divider()
                st.markdown("**Desglose Específico por Estado (Ej. Mano de Obra, Provisiones)**")
                
                tipo_filtro = st.radio("Selecciona qué analizar:", ["COSTO FIJO", "COSTO VARIABLE"], horizontal=True)
                df_filtro = df_gastos[df_gastos[c_tipo] == tipo_filtro]
                
                if not df_filtro.empty:
                    df_estado = df_filtro.groupby(c_est, as_index=False)['Saldo_Absoluto'].sum().sort_values(by='Saldo_Absoluto', ascending=False)
                    fig_est = px.bar(df_estado, y=c_est, x='Saldo_Absoluto', orientation='h', text_auto='.2s')
                    st.plotly_chart(fig_est, use_container_width=True)
                    
                    with st.expander(f"Ver tabla de detalles - {tipo_filtro}"):
                        st.dataframe(df_filtro[[buscar_columna(df_filtro, ["CUENTA", "ID"]), c_est, c_cat, 'Saldo_Absoluto']], use_container_width=True)
                else:
                    st.write(f"No hay registros clasificados como {tipo_filtro}.")
            else:
                st.write("No hay gastos registrados para analizar en el dashboard.")
        else:
            st.info("👈 Por favor, carga los archivos en la Pestaña 1 primero.")