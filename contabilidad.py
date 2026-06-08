import streamlit as st
import pandas as pd
import plotly.express as px
from utils import obtener_dataframe

# --- FUNCIONES AUXILIARES ---
def limpiar_texto(texto):
    return str(texto).strip().upper()

def buscar_columna(df, palabras_clave, todas=False):
    for col in df.columns:
        if todas:
            if all(p in str(col).upper() for p in palabras_clave):
                return col
        else:
            if any(p in str(col).upper() for p in palabras_clave):
                return col
    return None

def mostrar_modulo_contabilidad():
    st.title("📊 Análisis Financiero y Punto de Equilibrio")
    
    tab_carga, tab_pe, tab_dash, tab_matriz = st.tabs([
        "📥 1. Carga", 
        "🎛️ 2. Simulador y PE", 
        "📊 3. Dashboard",
        "📑 4. Matriz (Junta)"
    ])

    # ==========================================
    # PESTAÑA 1: CARGA DE ARCHIVOS
    # ==========================================
    with tab_carga:
        st.subheader("Configuración del Periodo y Archivos")
        
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
            
        st.markdown("---")
        archivos_subidos = st.file_uploader("Sube los archivos Excel mensuales:", type=["xlsx", "xls"], accept_multiple_files=True)

        mapeo_archivos = {}
        if archivos_subidos:
            st.write("**Clasificación de Archivos**")
            st.write("Asigna la unidad de negocio correspondiente a cada archivo subido para construir la matriz:")
            for arch in archivos_subidos:
                unidad = st.selectbox(
                    f"Unidad para el archivo '{arch.name}':", 
                    ["Cafetería", "Talleres", "Librería", "Proyectos", "Global", "Otros"], 
                    key=arch.name
                )
                mapeo_archivos[arch.name] = {"file": arch, "unidad": unidad}

        if st.button("🔄 Procesar y Sincronizar", type="primary", use_container_width=True):
            if not archivos_subidos:
                st.warning("⚠️ Sube al menos un archivo.")
            else:
                with st.spinner("Procesando archivos y estructurando datos..."):
                    try:
                        df_map = obtener_dataframe("Balance_Mapeado")
                        if df_map is None:
                            st.error("Error al conectar con 'Balance_Mapeado'.")
                            st.stop()
                        
                        df_map.columns = df_map.columns.str.strip().str.upper()
                        
                        c_map_cta = buscar_columna(df_map, ["CUENTA", "ID"])
                        c_map_tipo = buscar_columna(df_map, ["TIPO"])
                        c_map_est = buscar_columna(df_map, ["ESTADO"])
                        c_map_cat = buscar_columna(df_map, ["CATEGOR"])
                        c_map_nom = buscar_columna(df_map, ["NOMBRE", "DESCRIP"])
                        
                        df_map[c_map_cta] = df_map[c_map_cta].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                        df_map[c_map_tipo] = df_map[c_map_tipo].apply(limpiar_texto)
                        
                        # Filtrar mapa base para la matriz
                        df_map_valido = df_map[~df_map[c_map_tipo].isin(["NO APLICA", "CUENTA DE MAYOR", ""])].copy()
                        df_matriz = df_map_valido.copy()
                        
                        dfs_cons = []
                        unidades_procesadas = []

                        for nombre_arch, info in mapeo_archivos.items():
                            df_temp = pd.read_excel(info["file"], dtype=str)
                            df_temp.columns = df_temp.columns.str.strip().str.upper()
                            
                            c_arch_cta = buscar_columna(df_temp, ["CUENTA", "ID"])
                            c_arch_sld = buscar_columna(df_temp, ["SALDO", "FINAL"], todas=True)
                            
                            if not c_arch_sld or not c_arch_cta:
                                st.error(f"Error en {nombre_arch}: No se encontraron las columnas de Cuenta o Saldo Final.")
                                st.stop()

                            df_temp[c_arch_cta] = df_temp[c_arch_cta].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                            df_temp[c_arch_sld] = df_temp[c_arch_sld].astype(str).str.replace(r'[^\d\.\-]', '', regex=True)
                            df_temp[c_arch_sld] = pd.to_numeric(df_temp[c_arch_sld], errors='coerce').fillna(0)
                            
                            df_agrup = df_temp.groupby(c_arch_cta, as_index=False)[c_arch_sld].sum()
                            
                            # Para el consolidado total (Pestañas 2 y 3)
                            df_temp_cons = df_agrup.copy()
                            df_temp_cons.rename(columns={c_arch_sld: "SALDO_FINAL_GLOBAL"}, inplace=True)
                            dfs_cons.append(df_temp_cons)
                            
                            # Para la matriz separada (Pestaña 4)
                            unidad_col = info["unidad"]
                            # Si se suben varios de la misma unidad, agrupar nombres
                            if unidad_col in unidades_procesadas:
                                unidad_col = f"{unidad_col} ({nombre_arch})"
                            unidades_procesadas.append(unidad_col)
                            
                            df_agrup.rename(columns={c_arch_sld: unidad_col}, inplace=True)
                            df_matriz = pd.merge(df_matriz, df_agrup[[c_arch_cta, unidad_col]], left_on=c_map_cta, right_on=c_arch_cta, how="left")
                            df_matriz[unidad_col] = df_matriz[unidad_col].fillna(0)
                            if c_arch_cta in df_matriz.columns and c_arch_cta != c_map_cta:
                                df_matriz.drop(columns=[c_arch_cta], inplace=True)

                        # Procesamiento Consolidado
                        df_total = pd.concat(dfs_cons, ignore_index=True)
                        df_total_agrup = df_total.groupby(c_arch_cta, as_index=False)["SALDO_FINAL_GLOBAL"].sum()
                        
                        df_final = pd.merge(df_map_valido, df_total_agrup, left_on=c_map_cta, right_on=c_arch_cta, how="inner")
                        
                        # Almacenamiento en sesión
                        st.session_state['cont_df'] = df_final
                        st.session_state['cont_matriz'] = df_matriz
                        st.session_state['unidades_procesadas'] = unidades_procesadas
                        st.session_state['cols_dict'] = {
                            'cta': c_map_cta, 'tipo': c_map_tipo, 'est': c_map_est, 
                            'cat': c_map_cat, 'nom': c_map_nom, 'sld': "SALDO_FINAL_GLOBAL"
                        }
                        
                        st.success("✅ Procesamiento completado. Revisa las pestañas de Simulador y Matriz.")
                    except Exception as e:
                        st.error(f"Error técnico: {e}")

    # ==========================================
    # PESTAÑA 2: SIMULADOR Y PE
    # ==========================================
    with tab_pe:
        if 'cont_df' in st.session_state:
            df = st.session_state['cont_df']
            cd = st.session_state['cols_dict']
            
            st.subheader("🎛️ Simulador Financiero en Tiempo Real")
            st.write("Modifica los saldos o reclasifica el tipo de cuenta en la tabla inferior. El Punto de Equilibrio y el Dashboard se recalcularán automáticamente basándose en estos cambios.")
            
            df_editable = df[[cd['cta'], cd['nom'], cd['est'], cd['cat'], cd['tipo'], cd['sld']]].copy()
            
            edited_df = st.data_editor(
                df_editable,
                column_config={
                    cd['tipo']: st.column_config.SelectboxColumn("Tipo de Costo", options=["COSTO FIJO", "COSTO VARIABLE", "INGRESOS", "INGRESO"]),
                    cd['sld']: st.column_config.NumberColumn("Monto ($)", format="$ %.2f", min_value=0)
                },
                use_container_width=True,
                hide_index=True,
                key="editor_escenarios"
            )
            
            st.session_state['edited_df'] = edited_df
            
            # Cálculos usando los datos editados
            ventas = edited_df[edited_df[cd['tipo']].isin(["INGRESOS", "INGRESO"])][cd['sld']].abs().sum()
            cf = edited_df[edited_df[cd['tipo']] == "COSTO FIJO"][cd['sld']].abs().sum()
            cv = edited_df[edited_df[cd['tipo']] == "COSTO VARIABLE"][cd['sld']].abs().sum()
            
            margen_pct = (1 - (cv / ventas)) if ventas > 0 else 0
            pe = (cf / margen_pct) if margen_pct > 0 else 0
            
            st.divider()
            c1, c2, c3 = st.columns(3)
            c1.metric("Ingresos Proyectados", f"${ventas:,.2f}")
            c2.metric("Costos Variables (CV)", f"${cv:,.2f}")
            c3.metric("Costos Fijos (CF)", f"${cf:,.2f}")
            
            col_m, col_p = st.columns(2)
            with col_m:
                st.info("### Margen de Contribución")
                st.markdown(f"<h1 style='text-align: center; color: #4CAF50;'>{margen_pct * 100:.2f}%</h1>", unsafe_allow_html=True)
                
            with col_p:
                if margen_pct > 0:
                    st.success("### Punto de Equilibrio ($)")
                    st.markdown(f"<h1 style='text-align: center;'>${pe:,.2f}</h1>", unsafe_allow_html=True)
                else:
                    st.error("### Punto de Equilibrio ($)")
                    st.markdown("<h3 style='text-align: center;'>Incalculable</h3>", unsafe_allow_html=True)
        else:
            st.info("👈 Realiza la sincronización en la Pestaña 1.")

    # ==========================================
    # PESTAÑA 3: DASHBOARD ANALÍTICO
    # ==========================================
    with tab_dash:
        if 'edited_df' in st.session_state:
            df_dash = st.session_state['edited_df']
            cd = st.session_state['cols_dict']
            
            st.subheader("📊 Desglose basado en la Simulación")
            
            df_gastos = df_dash[df_dash[cd['tipo']].isin(["COSTO FIJO", "COSTO VARIABLE"])]
            
            if not df_gastos.empty:
                col_g1, col_g2 = st.columns(2)
                with col_g1:
                    fig_tipo = px.pie(df_gastos, values=cd['sld'], names=cd['tipo'], hole=0.4, title="Proporción de Gastos")
                    st.plotly_chart(fig_tipo, use_container_width=True)
                    
                with col_g2:
                    df_cat = df_gastos.groupby(cd['cat'], as_index=False)[cd['sld']].sum().sort_values(by=cd['sld'], ascending=False)
                    fig_cat = px.bar(df_cat, x=cd['cat'], y=cd['sld'], text_auto='.2s', title="Top Gastos por Categoría")
                    st.plotly_chart(fig_cat, use_container_width=True)
                    
                st.divider()
                tipos_existentes = df_gastos[cd['tipo']].unique().tolist()
                tipo_filtro = st.selectbox("Filtrar estado por:", tipos_existentes)
                
                df_filtro = df_gastos[df_gastos[cd['tipo']] == tipo_filtro]
                if not df_filtro.empty:
                    df_estado = df_filtro.groupby(cd['est'], as_index=False)[cd['sld']].sum().sort_values(by=cd['sld'], ascending=False)
                    fig_est = px.bar(df_estado, y=cd['est'], x=cd['sld'], orientation='h', text_auto='.2s', title=f"Detalle de {tipo_filtro}")
                    st.plotly_chart(fig_est, use_container_width=True)
            else:
                st.write("No hay gastos registrados en la simulación.")
        else:
            st.info("👈 Realiza la sincronización en la Pestaña 1.")

    # ==========================================
    # PESTAÑA 4: MATRIZ DE UNIDADES
    # ==========================================
    with tab_matriz:
        if 'cont_matriz' in st.session_state:
            df_matriz = st.session_state['cont_matriz']
            unidades = st.session_state['unidades_procesadas']
            
            st.subheader("📑 Reporte de Junta: Matriz por Unidad")
            
            # Calcular columna Total
            df_matriz['TOTAL CONSOLIDADO'] = df_matriz[unidades].sum(axis=1)
            
            # Formato moneda para visualización
            columnas_numericas = unidades + ['TOTAL CONSOLIDADO']
            df_display = df_matriz.copy()
            for col in columnas_numericas:
                df_display[col] = df_display[col].apply(lambda x: f"${x:,.2f}")
                
            st.dataframe(df_display, use_container_width=True, hide_index=True)
            
            csv_matriz = df_matriz.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 Descargar Matriz (CSV para Excel)",
                data=csv_matriz,
                file_name='matriz_unidades.csv',
                mime='text/csv',
            )
        else:
            st.info("👈 Realiza la sincronización en la Pestaña 1.")