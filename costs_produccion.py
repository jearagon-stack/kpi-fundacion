import streamlit as st
import pandas as pd
import io
from datetime import date

# ==========================================
# FUNCIONES AUXILIARES
# ==========================================
def limpiar_codigo(c):
    """Respeta ceros a la izquierda, solo quita '.0' de los excels"""
    if pd.isna(c) or str(c).strip() == "": return "SIN_CODIGO"
    val = str(c).strip().upper()
    if val.endswith('.0'):
        return val[:-2]
    return val

def generar_excel_proyeccion(df, nombre_hoja="Proyeccion_Compras"):
    """Genera el archivo Excel final para descarga nativa"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=nombre_hoja)
    return output.getvalue()

# ==========================================
# MÓDULO PRINCIPAL DE PRODUCCIÓN
# ==========================================
def mostrar_modulo_produccion():
    st.title("🏭 Módulo de Gestión de Producción")
    st.info("Planificación Estratégica, Proyección de Compras y Estructura de Recetas.")

    tab_proyeccion, tab_recetas = st.tabs([
        "📊 1. Proyección de Compras", 
        "🥣 2. Recetas"
    ])

    # ==========================================
    # PESTAÑA 1: PROYECCIÓN DE COMPRAS 
    # ==========================================
    with tab_proyeccion:
        st.subheader("Simulador y Proyección de Compras de Materia Prima")
        
        st.markdown("##### ⏱️ Configuración a Proyectar")
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            dias_proyectar = st.number_input("Días a Proyectar (Ej: 7 para 1 semana):", min_value=1, value=7, key="prod_dias_proj")
        with col_t2:
            dias_operativos = st.number_input("Días laborados por semana (Informativo):", min_value=1.0, max_value=7.0, value=6.5, step=0.5, key="prod_dias_op")

        st.markdown("---")
        st.markdown("##### 📥 Carga de Archivos Maestros")
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            arch_stock = st.file_uploader("1. Inventario / Stock Actual Saneado", type=["xlsx", "xls"], key="prod_f_stock")
            arch_salidas = st.file_uploader("2. Historial de Movimientos de Bodega (Consumos)", type=["xlsx", "xls"], key="prod_f_salidas")
        with col_f2:
            arch_compras = st.file_uploader("3. Historial de Compras (Facturas de Proveedores)", type=["xlsx", "xls"], key="prod_f_compras")

        if arch_stock and arch_salidas and arch_compras:
            # Lectura preliminar rápida para obtener encabezados reales
            try:
                df_s_pre = pd.read_excel(arch_stock, nrows=1, dtype=str)
                df_m_pre = pd.read_excel(arch_salidas, nrows=1, dtype=str)
                df_c_pre = pd.read_excel(arch_compras, nrows=1, dtype=str)
                
                # --- NUEVO: MAPEO DINÁMICO DE COLUMNAS ---
                st.markdown("### 🔍 Mapeo de Columnas de Origen")
                st.info("Selecciona la columna exacta para cada dato. Esto garantiza un cruce perfecto y evita la pérdida de filas.")
                
                col_sel1, col_sel2, col_sel3 = st.columns(3)
                
                with col_sel1:
                    st.caption("📦 1. Stock Actual")
                    col_cod_s = st.selectbox("Columna Código:", df_s_pre.columns, index=0)
                    col_desc_s = st.selectbox("Columna Descripción:", df_s_pre.columns, index=1 if len(df_s_pre.columns) > 1 else 0)
                    col_cant_s = st.selectbox("Columna Cantidad:", df_s_pre.columns, index=2 if len(df_s_pre.columns) > 2 else 0)

                with col_sel2:
                    st.caption("🔄 2. Salidas / Ajustes")
                    col_cod_m = st.selectbox("Columna Código:", df_m_pre.columns, index=0, key='cod_m')
                    col_cant_m = st.selectbox("Columna Cantidad:", df_m_pre.columns, index=2 if len(df_m_pre.columns) > 2 else 0, key='cant_m')
                    col_tipo_m = st.selectbox("Columna Tipo Movimiento (Opcional):", ["No utilizar"] + list(df_m_pre.columns), index=0)
                    col_fec_m = st.selectbox("Columna Fecha (Opcional):", ["No utilizar"] + list(df_m_pre.columns), index=0)

                with col_sel3:
                    st.caption("💰 3. Compras Históricas")
                    col_cod_c = st.selectbox("Columna Código:", df_c_pre.columns, index=0, key='cod_c')
                    col_cant_c = st.selectbox("Columna Cantidad:", df_c_pre.columns, index=2 if len(df_c_pre.columns) > 2 else 0, key='cant_c')
                    col_fec_c = st.selectbox("Columna Fecha (Opcional):", ["No utilizar"] + list(df_c_pre.columns), index=0, key='fec_c')

                if st.button("🚀 Ejecutar Análisis de Demanda y Variación", type="primary", use_container_width=True):
                    with st.spinner("Procesando, cruzando y normalizando información basada en su selección..."):
                        try:
                            # 1. Leer Archivos completos
                            df_s_raw = pd.read_excel(arch_stock, dtype=str)
                            df_out_raw = pd.read_excel(arch_salidas, dtype=str)
                            df_in_raw = pd.read_excel(arch_compras, dtype=str)

                            # 2. Limpieza de Datos
                            df_s_raw['Cod_Clean'] = df_s_raw[col_cod_s].apply(limpiar_codigo)
                            df_s_raw['Stock_Num'] = pd.to_numeric(df_s_raw[col_cant_s], errors='coerce').fillna(0)
                            
                            df_out_raw['Cod_Clean'] = df_out_raw[col_cod_m].apply(limpiar_codigo)
                            df_out_raw['Cant_Abs'] = pd.to_numeric(df_out_raw[col_cant_m], errors='coerce').fillna(0).abs()
                            
                            if col_tipo_m != "No utilizar":
                                def calcular_consumo_neto(row):
                                    tipo = str(row[col_tipo_m]).upper()
                                    if 'ENTRADA' in tipo or 'INGRESO' in tipo:
                                        return -row['Cant_Abs']
                                    return row['Cant_Abs']
                                df_out_raw['Cant_Num'] = df_out_raw.apply(calcular_consumo_neto, axis=1)
                            else:
                                df_out_raw['Cant_Num'] = df_out_raw['Cant_Abs']
                                
                            # Limpiar fechas
                            if col_fec_m != "No utilizar": 
                                df_out_raw['Fecha_Clean'] = pd.to_datetime(df_out_raw[col_fec_m], errors='coerce')
                                df_out_raw.loc[(df_out_raw['Fecha_Clean'].dt.year < 2020) | (df_out_raw['Fecha_Clean'].dt.year > 2030), 'Fecha_Clean'] = pd.NaT

                            df_in_raw['Cod_Clean'] = df_in_raw[col_cod_c].apply(limpiar_codigo)
                            df_in_raw['Cant_Num'] = pd.to_numeric(df_in_raw[col_cant_c], errors='coerce').fillna(0).abs()
                            
                            if col_fec_c != "No utilizar": 
                                df_in_raw['Fecha_Clean'] = pd.to_datetime(df_in_raw[col_fec_c], errors='coerce')
                                df_in_raw.loc[(df_in_raw['Fecha_Clean'].dt.year < 2020) | (df_in_raw['Fecha_Clean'].dt.year > 2030), 'Fecha_Clean'] = pd.NaT

                            # Cálculo de Días Históricos
                            min_dates = []
                            max_dates = []
                            
                            if col_fec_m != "No utilizar" and 'Fecha_Clean' in df_out_raw.columns and not df_out_raw['Fecha_Clean'].isna().all():
                                min_dates.append(df_out_raw['Fecha_Clean'].min())
                                max_dates.append(df_out_raw['Fecha_Clean'].max())
                            
                            if col_fec_c != "No utilizar" and 'Fecha_Clean' in df_in_raw.columns and not df_in_raw['Fecha_Clean'].isna().all():
                                min_dates.append(df_in_raw['Fecha_Clean'].min())
                                max_dates.append(df_in_raw['Fecha_Clean'].max())

                            if min_dates and max_dates:
                                fecha_min_global = min(min_dates)
                                fecha_max_global = max(max_dates)
                                dias_historial_calculado = (fecha_max_global - fecha_min_global).days + 1
                                rango_fechas_str = f"desde {fecha_min_global.strftime('%d/%m/%Y')} hasta {fecha_max_global.strftime('%d/%m/%Y')}"
                            else:
                                dias_historial_calculado = 180 # Valor por defecto seguro
                                rango_fechas_str = "No detectado (Usando 180 días por defecto)"

                            if dias_historial_calculado < 1: dias_historial_calculado = 1

                            st.session_state['prod_dias_hist'] = dias_historial_calculado
                            st.session_state['prod_rango_fechas'] = rango_fechas_str

                            # 4. Agrupaciones seguras (Manteniendo el stock como maestro para no perder filas)
                            df_stock_group = df_s_raw.groupby('Cod_Clean').agg({
                                col_desc_s: 'first',
                                'Stock_Num': 'sum'
                            }).reset_index().rename(columns={'Stock_Num': 'Stock_Actual', col_desc_s: 'Descripción'})
                            
                            df_salidas_group = df_out_raw.groupby('Cod_Clean')['Cant_Num'].sum().reset_index().rename(columns={'Cant_Num': 'Consumo_Historico'})
                            df_compras_group = df_in_raw.groupby('Cod_Clean')['Cant_Num'].sum().reset_index().rename(columns={'Cant_Num': 'Compras_Historicas'})
                            df_min_lot = df_in_raw[df_in_raw['Cant_Num'] > 0].groupby('Cod_Clean')['Cant_Num'].min().reset_index().rename(columns={'Cant_Num': 'Lot_Min'})
                            df_max_lot = df_in_raw.groupby('Cod_Clean')['Cant_Num'].max().reset_index().rename(columns={'Cant_Num': 'Lot_Max'})

                            # 5. Consolidación: Left Join sobre la base de Stock para conservar los 284 registros intactos
                            df_maestro = df_stock_group
                            df_maestro = df_maestro.merge(df_salidas_group, on='Cod_Clean', how='left')
                            df_maestro = df_maestro.merge(df_compras_group, on='Cod_Clean', how='left')
                            df_maestro = df_maestro.merge(df_min_lot, on='Cod_Clean', how='left')
                            df_maestro = df_maestro.merge(df_max_lot, on='Cod_Clean', how='left')

                            df_maestro = df_maestro.rename(columns={'Cod_Clean': 'Código'})
                            df_maestro['Descripción'] = df_maestro['Descripción'].fillna('PRODUCTO SIN NOMBRE')
                            df_maestro['Stock_Actual'] = df_maestro['Stock_Actual'].fillna(0.0)
                            df_maestro['Consumo_Historico'] = df_maestro['Consumo_Historico'].fillna(0.0)
                            df_maestro['Compras_Historicas'] = df_maestro['Compras_Historicas'].fillna(0.0)
                            df_maestro['Stock_Seguridad_Pct'] = 5.0 

                            st.session_state['prod_df_calculo_base'] = df_maestro
                            st.session_state['prod_ejecutado'] = True

                        except Exception as e:
                            st.error(f"Error procesando la información: {e}")

            except Exception as e:
                st.warning("Asegúrate de que los archivos tengan el formato correcto de Excel.")

        # Sección del Simulador Interactivo
        if st.session_state.get('prod_ejecutado', False):
            dias_hist_calc = st.session_state.get('prod_dias_hist', 180)
            rango_fechas = st.session_state.get('prod_rango_fechas', '')
            
            st.success(f"✅ Análisis completado. Historial detectado: **{dias_hist_calc} días** ({rango_fechas}).")
            st.markdown("---")
            st.subheader("🎛️ Panel de Simulación")

            df_base = st.session_state['prod_df_calculo_base']
            df_calculado = df_base.copy()

            # Cálculos Semanales base
            semanas_historial = dias_hist_calc / 7.0
            semanas_proyectar = dias_proyectar / 7.0

            df_calculado['Consumo_Semanal'] = df_calculado['Consumo_Historico'] / semanas_historial
            df_calculado['Compra_Semanal'] = df_calculado['Compras_Historicas'] / semanas_historial

            df_interactivo = st.data_editor(
                df_calculado[['Código', 'Descripción', 'Stock_Actual', 'Consumo_Semanal', 'Compra_Semanal', 'Stock_Seguridad_Pct', 'Lot_Min', 'Lot_Max']],
                column_config={
                    "Código": st.column_config.TextColumn("Código", disabled=True),
                    "Descripción": st.column_config.TextColumn("Descripción", disabled=True),
                    "Stock_Actual": st.column_config.NumberColumn("Stock", disabled=True, format="%.2f"),
                    "Consumo_Semanal": st.column_config.NumberColumn("Consumo/Semana", disabled=True, format="%.2f"),
                    "Compra_Semanal": st.column_config.NumberColumn("Compra/Semana", disabled=True, format="%.2f"),
                    "Stock_Seguridad_Pct": st.column_config.NumberColumn("Seguridad (%)", min_value=0.0, max_value=100.0, format="%.1f"),
                    "Lot_Min": st.column_config.NumberColumn("Min Lot", disabled=True, format="%.2f"),
                    "Lot_Max": st.column_config.NumberColumn("Max Lot", disabled=True, format="%.2f")
                },
                hide_index=True,
                use_container_width=True,
                key="editor_interactivo_prod"
            )

            # CÁLCULOS MATEMÁTICOS FINALES
            df_final = df_interactivo.copy()
            
            df_final['Consumo_Proyectado'] = df_final['Consumo_Semanal'] * semanas_proyectar
            df_final['Monto_Seguridad'] = df_final['Consumo_Proyectado'] * (df_final['Stock_Seguridad_Pct'] / 100.0)
            
            df_final['Proyección 1 (Consumo Real)'] = (df_final['Consumo_Proyectado'] + df_final['Monto_Seguridad']) - df_final['Stock_Actual']
            df_final['Proyección 1 (Consumo Real)'] = df_final['Proyección 1 (Consumo Real)'].apply(lambda x: max(0.0, round(x, 2)))

            df_final['Proyección 2 (Ritmo Compras)'] = (df_final['Compra_Semanal'] * semanas_proyectar).round(2)

            def calcular_var(row):
                p1 = row['Proyección 1 (Consumo Real)']
                p2 = row['Proyección 2 (Ritmo Compras)']
                if p2 == 0:
                    return 100.0 if p1 > 0 else 0.0
                return round(((p1 - p2) / p2) * 100.0, 2)

            df_final['Variación (%)'] = df_final.apply(calcular_var, axis=1)

            def evaluar_semaforo(row):
                p1 = row['Proyección 1 (Consumo Real)']
                l_min = pd.to_numeric(row['Lot_Min'], errors='coerce')
                l_max = pd.to_numeric(row['Lot_Max'], errors='coerce')
                
                if p1 == 0: return "🟢 Stock Suficiente"
                if pd.notna(l_min) and p1 < l_min: return f"🔴 Menor a Mín. ({l_min:,.1f})"
                if pd.notna(l_max) and p1 > l_max: return f"🟡 Mayor a Máx. ({l_max:,.1f})"
                return "🟢 Rango Tradicional"

            df_final['Semáforo / Alerta'] = df_final.apply(evaluar_semaforo, axis=1)

            columnas_finales = [
                'Código', 'Descripción', 'Stock_Actual', 'Consumo_Semanal', 'Compra_Semanal',
                'Proyección 1 (Consumo Real)', 'Proyección 2 (Ritmo Compras)', 'Variación (%)', 'Semáforo / Alerta'
            ]
            df_vista_final = df_final[columnas_finales]

            st.markdown("##### 📋 Resultados Finales Calculados")
            st.dataframe(df_vista_final, use_container_width=True, hide_index=True)

            st.download_button(
                label="📥 Descargar Reporte de Proyecciones en Excel (.xlsx)",
                data=generar_excel_proyeccion(df_vista_final),
                file_name=f"Proyeccion_Compras_semana_{date.today().strftime('%d_%m_%Y')}.xlsx",
                type="primary",
                use_container_width=True
            )

    # ==========================================
    # PESTAÑA 2: RECETAS
    # ==========================================
    with tab_recetas:
        st.subheader("Estructura de Recetas (BOM - Bill of Materials)")
        st.write("Gestión de explosión de ingredientes y costos unitarios de fabricación.")
        st.info("Pestaña lista para recibir la lógica de recetas en la siguiente etapa.")