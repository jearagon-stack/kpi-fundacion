import streamlit as st
import pandas as pd
import plotly.express as px
import io
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

def generar_excel(df):
    """Convierte un DataFrame a un archivo Excel en memoria listo para descargar."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Reporte')
    return output.getvalue()

def mostrar_modulo_contabilidad():
    st.title("📊 Análisis Financiero y Punto de Equilibrio")
    
    tab_carga, tab_pe, tab_dash, tab_matriz = st.tabs([
        "📥 1. Carga General", 
        "🎛️ 2. Simulador y PE", 
        "📊 3. Dashboard",
        "📑 4. Matriz (Junta)"
    ])

    # ==========================================
    # PESTAÑA 1: CARGA GENERAL (SIMULADOR)
    # ==========================================
    with tab_carga:
        st.subheader("Configuración para el Simulador Global")
        st.write("Sube aquí los archivos para tu análisis de escenarios. (La matriz de Junta se arma en la pestaña 4).")
        
        col1, col2, col3, col4 = st.columns(4)
        meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        with col1:
            mes_desde = st.selectbox("Mes Desde:", meses, index=0)
        with col2:
            anio_desde = st.number_input("Año Desde:", min_value=2024, max_value=2030, value=2026)
        with col3:
            mes_hasta = st.selectbox("Mes Hasta:", meses, index=11)
        with col4:
            anio_hasta = st.number_input("Año Hasta:", min_value=2024, max_value=2030, value=2026)
            
        st.markdown("---")
        archivos_subidos = st.file_uploader("Sube los archivos Excel para el análisis global:", type=["xlsx", "xls"], accept_multiple_files=True, key="up_general")

        if st.button("🔄 Procesar Consolidado Global", type="primary", use_container_width=True):
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
                        
                        df_map[c_map_cta] = df_map[c_map_cta].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                        df_map[c_map_tipo] = df_map[c_map_tipo].apply(limpiar_texto)
                        
                        df_map_valido = df_map[~df_map[c_map_tipo].isin(["NO APLICA", "CUENTA DE MAYOR", ""])].copy()
                        
                        dfs_cons = []
                        for arch in archivos_subidos:
                            df_temp = pd.read_excel(arch, dtype=str)
                            df_temp.columns = df_temp.columns.str.strip().str.upper()
                            
                            c_arch_cta = buscar_columna(df_temp, ["CUENTA", "ID"])
                            c_arch_sld = buscar_columna(df_temp, ["SALDO", "FINAL"], todas=True)
                            
                            if not c_arch_sld or not c_arch_cta:
                                continue

                            df_temp[c_arch_cta] = df_temp[c_arch_cta].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                            df_temp[c_arch_sld] = df_temp[c_arch_sld].astype(str).str.replace(r'[^\d\.\-]', '', regex=True)
                            df_temp[c_arch_sld] = pd.to_numeric(df_temp[c_arch_sld], errors='coerce').fillna(0)
                            
                            df_agrup = df_temp.groupby(c_arch_cta, as_index=False)[c_arch_sld].sum()
                            df_agrup.rename(columns={c_arch_sld: "SALDO_FINAL_GLOBAL"}, inplace=True)
                            dfs_cons.append(df_agrup)

                        df_total = pd.concat(dfs_cons, ignore_index=True)
                        df_total_agrup = df_total.groupby(c_arch_cta, as_index=False)["SALDO_FINAL_GLOBAL"].sum()
                        
                        df_final = pd.merge(df_map_valido, df_total_agrup, left_on=c_map_cta, right_on=c_arch_cta, how="inner")
                        
                        st.session_state['cont_df'] = df_final
                        st.session_state['cols_dict'] = {
                            'cta': c_map_cta, 'tipo': c_map_tipo, 'est': c_map_est, 
                            'cat': c_map_cat, 'sld': "SALDO_FINAL_GLOBAL"
                        }
                        st.session_state['sync_flag'] = True 
                        
                        st.success("✅ Procesamiento completado. Revisa el Simulador (Pestaña 2).")
                    except Exception as e:
                        st.error(f"Error técnico: {e}")

    # ==========================================
    # PESTAÑA 2: SIMULADOR Y PE
    # ==========================================
    with tab_pe:
        if 'cont_df' in st.session_state:
            df = st.session_state['cont_df']
            cd = st.session_state['cols_dict']
            
            if 'df_master' not in st.session_state or st.session_state.get('sync_flag', False):
                st.session_state['df_master'] = df[[cd['cta'], cd['est'], cd['cat'], cd['tipo'], cd['sld']]].copy()
                st.session_state['sync_flag'] = False
                
            st.subheader("🎛️ Simulador Financiero")
            
            st.markdown("#### 1. Buscar Cuenta")
            col_f1, col_f2, col_f3 = st.columns(3)
            
            with col_f1:
                tipos_disp = st.session_state['df_master'][cd['tipo']].unique().tolist()
                filtro_tipo = st.selectbox("A. Filtrar por Tipo:", options=["Todos"] + tipos_disp)
                
            df_filtro = st.session_state['df_master'].copy()
            if filtro_tipo != "Todos":
                df_filtro = df_filtro[df_filtro[cd['tipo']] == filtro_tipo]
                
            with col_f2:
                cat_disp = df_filtro[cd['cat']].unique().tolist()
                filtro_cat = st.selectbox("B. Filtrar por Categoría:", options=["Todas"] + cat_disp)
                
            if filtro_cat != "Todas":
                df_filtro = df_filtro[df_filtro[cd['cat']] == filtro_cat]
                
            with col_f3:
                df_filtro['display_name'] = df_filtro[cd['cta']].astype(str) + " - " + df_filtro[cd['est']]
                opciones_cuentas = df_filtro['display_name'].tolist()
                cuenta_seleccionada = st.selectbox("C. Seleccionar Cuenta:", options=["Seleccione una cuenta..."] + opciones_cuentas)

            if cuenta_seleccionada != "Seleccione una cuenta...":
                st.markdown("#### 2. Configurar Escenario")
                id_cuenta_sel = cuenta_seleccionada.split(" - ")[0]
                idx = st.session_state['df_master'].index[st.session_state['df_master'][cd['cta']].astype(str) == id_cuenta_sel].tolist()[0]
                
                tipo_actual = st.session_state['df_master'].at[idx, cd['tipo']]
                monto_actual = float(st.session_state['df_master'].at[idx, cd['sld']])
                nombre_cuenta = st.session_state['df_master'].at[idx, cd['est']]
                
                col_ed1, col_ed2, col_ed3 = st.columns([2, 2, 1])
                with col_ed1:
                    opciones_tipo = ["COSTO FIJO", "COSTO VARIABLE", "INGRESOS", "INGRESO"]
                    if tipo_actual not in opciones_tipo:
                        opciones_tipo.append(tipo_actual)
                    nuevo_tipo = st.selectbox("Reclasificar Tipo:", options=opciones_tipo, index=opciones_tipo.index(tipo_actual))
                with col_ed2:
                    nuevo_monto = st.number_input("Proyectar Nuevo Monto ($):", value=monto_actual, min_value=0.0, format="%.2f")
                with col_ed3:
                    st.write("")
                    st.write("")
                    if st.button("✅ Aplicar Cambio", type="primary", use_container_width=True):
                        st.session_state['df_master'].at[idx, cd['tipo']] = nuevo_tipo
                        st.session_state['df_master'].at[idx, cd['sld']] = nuevo_monto
                        st.rerun()

            st.divider()
            
            df_math = st.session_state['df_master']
            ventas = df_math[df_math[cd['tipo']].isin(["INGRESOS", "INGRESO"])][cd['sld']].abs().sum()
            cf = df_math[df_math[cd['tipo']] == "COSTO FIJO"][cd['sld']].abs().sum()
            cv = df_math[df_math[cd['tipo']] == "COSTO VARIABLE"][cd['sld']].abs().sum()
            
            margen_pct = (1 - (cv / ventas)) if ventas > 0 else 0
            pe = (cf / margen_pct) if margen_pct > 0 else 0
            
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
                    
                    # --- INDICADOR DE DISTANCIA A LA META ---
                    diferencia_monto = ventas - pe
                    diferencia_pct = (diferencia_monto / pe) * 100 if pe > 0 else 0
                    
                    if diferencia_monto >= 0:
                        st.markdown(f"<div style='text-align: center; padding-top: 10px;'><span style='background-color: #d4edda; color: #155724; padding: 5px 10px; border-radius: 5px; font-weight: bold;'>▲ Superávit: +${diferencia_monto:,.2f} (+{diferencia_pct:.2f}%)</span></div>", unsafe_allow_html=True)
                    else:
                        st.markdown(f"<div style='text-align: center; padding-top: 10px;'><span style='background-color: #f8d7da; color: #721c24; padding: 5px 10px; border-radius: 5px; font-weight: bold;'>▼ Déficit: -${abs(diferencia_monto):,.2f} ({diferencia_pct:.2f}%)</span></div>", unsafe_allow_html=True)
                else:
                    st.error("### Punto de Equilibrio ($)")
                    st.markdown("<h3 style='text-align: center;'>Incalculable</h3>", unsafe_allow_html=True)
                    
            st.divider()
            st.subheader("🕵️ Exportar Escenario Actual")
            excel_sim = generar_excel(df_math)
            st.download_button(
                label="📥 Descargar Detalle en Excel (.xlsx)",
                data=excel_sim,
                file_name='auditoria_simulador.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
        else:
            st.info("👈 Realiza la sincronización en la Pestaña 1 primero.")

    # ==========================================
    # PESTAÑA 3: DASHBOARD ANALÍTICO
    # ==========================================
    with tab_dash:
        if 'df_master' in st.session_state:
            df_dash = st.session_state['df_master']
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
                tipo_filtro_dash = st.selectbox("Filtrar estado por:", tipos_existentes)
                df_filtro_dash = df_gastos[df_gastos[cd['tipo']] == tipo_filtro_dash]
                if not df_filtro_dash.empty:
                    df_estado = df_filtro_dash.groupby(cd['est'], as_index=False)[cd['sld']].sum().sort_values(by=cd['sld'], ascending=False)
                    fig_est = px.bar(df_estado, y=cd['est'], x=cd['sld'], orientation='h', text_auto='.2s', title=f"Detalle de {tipo_filtro_dash}")
                    st.plotly_chart(fig_est, use_container_width=True)
            else:
                st.write("No hay gastos registrados en la simulación.")
        else:
            st.info("👈 Realiza la sincronización en la Pestaña 1 primero.")

    # ==========================================
    # PESTAÑA 4: MATRIZ DE UNIDADES (INDEPENDIENTE)
    # ==========================================
    with tab_matriz:
        st.subheader("📑 Reporte de Junta: Matriz Independiente")
        st.write("Sube los archivos correspondientes y clasifícalos en las 9 unidades oficiales. Esta sección no mezcla datos con el simulador.")
        
        archivos_matriz = st.file_uploader("Sube los archivos Excel para armar la matriz:", type=["xlsx", "xls"], accept_multiple_files=True, key="up_matriz")
        
        unidades_oficiales = ["Cafetería", "Librería", "Centro Soho", "CID Campus", "Talleres Gráfico", "Despensa", "Terraza", "Servicios Generales", "Gerencias"]
        mapeo_matriz = {}
        
        if archivos_matriz:
            st.write("**Clasifica los archivos subidos:**")
            for arch in archivos_matriz:
                unidad_sel = st.selectbox(f"Unidad para '{arch.name}':", unidades_oficiales, key=f"sel_{arch.name}")
                mapeo_matriz[arch.name] = {"file": arch, "unidad": unidad_sel}
                
            if st.button("🔄 Armar Matriz de Junta", type="primary"):
                with st.spinner("Construyendo matriz por unidades..."):
                    try:
                        df_map_m = obtener_dataframe("Balance_Mapeado")
                        if df_map_m is None:
                            st.error("Error al conectar con 'Balance_Mapeado'.")
                            st.stop()
                            
                        df_map_m.columns = df_map_m.columns.str.strip().str.upper()
                        cm_cta = buscar_columna(df_map_m, ["CUENTA", "ID"])
                        cm_tipo = buscar_columna(df_map_m, ["TIPO"])
                        
                        df_map_m[cm_cta] = df_map_m[cm_cta].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                        df_map_m[cm_tipo] = df_map_m[cm_tipo].apply(limpiar_texto)
                        
                        df_base_matriz = df_map_m[~df_map_m[cm_tipo].isin(["NO APLICA", "CUENTA DE MAYOR", ""])].copy()
                        
                        for u in unidades_oficiales:
                            df_base_matriz[u] = 0.0

                        for nombre_arch, info in mapeo_matriz.items():
                            df_temp_m = pd.read_excel(info["file"], dtype=str)
                            df_temp_m.columns = df_temp_m.columns.str.strip().str.upper()
                            
                            c_arch_cta_m = buscar_columna(df_temp_m, ["CUENTA", "ID"])
                            c_arch_sld_m = buscar_columna(df_temp_m, ["SALDO", "FINAL"], todas=True)
                            
                            if not c_arch_sld_m or not c_arch_cta_m:
                                continue
                                
                            df_temp_m[c_arch_cta_m] = df_temp_m[c_arch_cta_m].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                            df_temp_m[c_arch_sld_m] = df_temp_m[c_arch_sld_m].astype(str).str.replace(r'[^\d\.\-]', '', regex=True)
                            df_temp_m[c_arch_sld_m] = pd.to_numeric(df_temp_m[c_arch_sld_m], errors='coerce').fillna(0)
                            
                            df_agrup_m = df_temp_m.groupby(c_arch_cta_m, as_index=False)[c_arch_sld_m].sum()
                            
                            unidad_obj = info["unidad"]
                            for _, row in df_agrup_m.iterrows():
                                cuenta = row[c_arch_cta_m]
                                monto = row[c_arch_sld_m]
                                idx_cuenta = df_base_matriz[df_base_matriz[cm_cta] == cuenta].index
                                if not idx_cuenta.empty:
                                    df_base_matriz.loc[idx_cuenta, unidad_obj] += monto

                        df_base_matriz['TOTAL CONSOLIDADO'] = df_base_matriz[unidades_oficiales].sum(axis=1)
                        df_base_matriz = df_base_matriz[df_base_matriz['TOTAL CONSOLIDADO'] != 0].copy()
                        
                        st.session_state['matriz_final'] = df_base_matriz
                        st.success("✅ Matriz construida con éxito.")
                    except Exception as e:
                        st.error(f"Error técnico en la matriz: {e}")

        if 'matriz_final' in st.session_state:
            df_mostrar = st.session_state['matriz_final'].copy()
            
            cols_num = unidades_oficiales + ['TOTAL CONSOLIDADO']
            for col in cols_num:
                df_mostrar[col] = df_mostrar[col].apply(lambda x: f"${x:,.2f}")
                
            st.dataframe(df_mostrar, use_container_width=True, hide_index=True)
            
            excel_matriz = generar_excel(st.session_state['matriz_final'])
            st.download_button(
                label="📥 Descargar Matriz en Excel (.xlsx)",
                data=excel_matriz,
                file_name='matriz_junta.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )