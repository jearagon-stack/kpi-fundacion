import streamlit as st
import pandas as pd
import io
import math
from datetime import date

# ==========================================
# FUNCIONES AUXILIARES
# ==========================================
def limpiar_codigo(c):
    """Mantiene el código puro y respeta ceros iniciales"""
    if pd.isna(c) or str(c).strip() == "": return "SIN_CODIGO"
    val = str(c).strip()
    if val.endswith('.0'): return val[:-2]
    return val

def generar_excel_proyeccion(df, nombre_hoja="Proyeccion_Compras"):
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
            arch_stock = st.file_uploader("1. Inventario / Stock Actual", type=["xlsx", "xls"], key="prod_f_stock")
            arch_salidas = st.file_uploader("2. Historial de Movimientos de Bodega", type=["xlsx", "xls"], key="prod_f_salidas")
        with col_f2:
            arch_compras = st.file_uploader("3. Historial de Compras", type=["xlsx", "xls"], key="prod_f_compras")

        if arch_stock and arch_salidas and arch_compras:
            if st.button("🚀 Ejecutar Análisis de Demanda y Variación", type="primary", use_container_width=True):
                with st.spinner("Procesando archivos y calculando proyecciones..."):
                    try:
                        df_s = pd.read_excel(arch_stock, dtype=str)
                        df_m = pd.read_excel(arch_salidas, dtype=str)
                        df_c = pd.read_excel(arch_compras, dtype=str)

                        df_s.columns = df_s.columns.str.strip()
                        df_m.columns = df_m.columns.str.strip()
                        df_c.columns = df_c.columns.str.strip()

                        # --- STOCK ---
                        df_s['Cod_Clean'] = df_s['IdProducto'].apply(limpiar_codigo)
                        df_s['Stock_Num'] = pd.to_numeric(df_s['Existencias'], errors='coerce').fillna(0)
                        df_stock_group = df_s.groupby('Cod_Clean').agg({'Nombre': 'first', 'Stock_Num': 'sum'}).reset_index().rename(columns={'Nombre': 'Descripción', 'Stock_Num': 'Stock_Actual'})

                        # --- MOVIMIENTOS ---
                        df_m['Cod_Clean'] = df_m['IdProducto'].apply(limpiar_codigo)
                        col_cant_mov = 'Cantidad' if 'Cantidad' in df_m.columns else 'Cantida'
                        df_m['Cant_Abs'] = pd.to_numeric(df_m[col_cant_mov], errors='coerce').fillna(0).abs()

                        def calcular_consumo_neto(row):
                            tipo = str(row['Tipo']).upper()
                            if 'ENTRADA' in tipo:
                                return -row['Cant_Abs']
                            return row['Cant_Abs']
                        
                        df_m['Consumo_Neto'] = df_m.apply(calcular_consumo_neto, axis=1)
                        df_salidas_group = df_m.groupby('Cod_Clean')['Consumo_Neto'].sum().reset_index().rename(columns={'Consumo_Neto': 'Consumo_Historico'})

                        # --- COMPRAS ---
                        df_c['Cod_Clean'] = df_c['IdProducto'].apply(limpiar_codigo)
                        df_c['Cant_Num'] = pd.to_numeric(df_c['Cantidad'], errors='coerce').fillna(0).abs()
                        df_compras_group = df_c.groupby('Cod_Clean')['Cant_Num'].sum().reset_index().rename(columns={'Cant_Num': 'Compras_Historicas'})
                        df_min_lot = df_c[df_c['Cant_Num'] > 0].groupby('Cod_Clean')['Cant_Num'].min().reset_index().rename(columns={'Cant_Num': 'Lot_Min'})
                        df_max_lot = df_c.groupby('Cod_Clean')['Cant_Num'].max().reset_index().rename(columns={'Cant_Num': 'Lot_Max'})

                        # --- FECHAS ---
                        min_dates, max_dates = [], []
                        
                        if 'Fecha' in df_m.columns:
                            fechas_m = pd.to_datetime(df_m['Fecha'], dayfirst=True, errors='coerce').dropna()
                            fechas_m = fechas_m[(fechas_m.dt.year >= 2020) & (fechas_m.dt.year <= 2030)]
                            if not fechas_m.empty:
                                min_dates.append(fechas_m.min())
                                max_dates.append(fechas_m.max())
                        
                        if 'Fecha' in df_c.columns:
                            fechas_c = pd.to_datetime(df_c['Fecha'], dayfirst=True, errors='coerce').dropna()
                            fechas_c = fechas_c[(fechas_c.dt.year >= 2020) & (fechas_c.dt.year <= 2030)]
                            if not fechas_c.empty:
                                min_dates.append(fechas_c.min())
                                max_dates.append(fechas_c.max())

                        if min_dates and max_dates:
                            fecha_min_global = min(min_dates)
                            fecha_max_global = max(max_dates)
                            dias_historial_calculado = (fecha_max_global - fecha_min_global).days + 1
                            rango_fechas_str = f"desde {fecha_min_global.strftime('%d/%m/%Y')} hasta {fecha_max_global.strftime('%d/%m/%Y')}"
                        else:
                            dias_historial_calculado = 180
                            rango_fechas_str = "No detectado (Usando 180 días por defecto)"
                            
                        if dias_historial_calculado < 1: dias_historial_calculado = 1

                        st.session_state['prod_dias_hist'] = dias_historial_calculado
                        st.session_state['prod_rango_fechas'] = rango_fechas_str

                        # Consolidación
                        df_maestro = df_stock_group.copy()
                        df_maestro = df_maestro.merge(df_salidas_group, on='Cod_Clean', how='left')
                        df_maestro = df_maestro.merge(df_compras_group, on='Cod_Clean', how='left')
                        df_maestro = df_maestro.merge(df_min_lot, on='Cod_Clean', how='left')
                        df_maestro = df_maestro.merge(df_max_lot, on='Cod_Clean', how='left')

                        df_maestro = df_maestro.rename(columns={'Cod_Clean': 'Código'})
                        df_maestro['Descripción'] = df_maestro['Descripción'].fillna('SIN NOMBRE')
                        df_maestro['Stock_Actual'] = df_maestro['Stock_Actual'].fillna(0.0)
                        df_maestro['Consumo_Historico'] = df_maestro['Consumo_Historico'].fillna(0.0)
                        df_maestro['Compras_Historicas'] = df_maestro['Compras_Historicas'].fillna(0.0)
                        df_maestro['Stock_Seguridad_Pct'] = 5.0 

                        st.session_state['prod_df_calculo_base'] = df_maestro
                        st.session_state['prod_ejecutado'] = True

                    except Exception as e:
                        st.error(f"Error procesando la información. Verifica el formato de tus archivos. Error: {e}")

        if st.session_state.get('prod_ejecutado', False):
            dias_hist_calc = st.session_state.get('prod_dias_hist', 180)
            rango_fechas = st.session_state.get('prod_rango_fechas', '')
            
            st.success(f"✅ Análisis completado. Historial detectado: **{dias_hist_calc} días** ({rango_fechas}).")
            st.markdown("---")
            st.subheader("🎛️ Panel de Simulación y Decisión de Compra")

            df_base = st.session_state['prod_df_calculo_base']
            df_calculado = df_base.copy()

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
                    "Consumo_Semanal": st.column_config.NumberColumn("Consumo/Sem", disabled=True, format="%.2f"),
                    "Compra_Semanal": st.column_config.NumberColumn("Compra/Sem", disabled=True, format="%.2f"),
                    "Stock_Seguridad_Pct": st.column_config.NumberColumn("Seguridad (%)", min_value=0.0, max_value=100.0, format="%.1f"),
                    "Lot_Min": st.column_config.NumberColumn("Min Lot", disabled=True, format="%.2f"),
                    "Lot_Max": st.column_config.NumberColumn("Max Lot", disabled=True, format="%.2f")
                },
                hide_index=True,
                use_container_width=True,
                key="editor_interactivo_prod"
            )

            # --- CÁLCULOS FINALES ---
            df_final = df_interactivo.copy()
            
            df_final['Consumo_Proyectado'] = df_final['Consumo_Semanal'] * semanas_proyectar
            df_final['Monto_Seguridad'] = df_final['Consumo_Proyectado'] * (df_final['Stock_Seguridad_Pct'] / 100.0)
            
            df_final['Calculo Exacto'] = (df_final['Consumo_Proyectado'] + df_final['Monto_Seguridad']) - df_final['Stock_Actual']
            df_final['Calculo Exacto'] = df_final['Calculo Exacto'].apply(lambda x: max(0.0, round(x, 2)))

            df_final['🛒 A COMPRAR (Unidades)'] = df_final['Calculo Exacto'].apply(lambda x: math.ceil(x))

            def calcular_var(row):
                c_sem = row['Consumo_Semanal']
                comp_sem = row['Compra_Semanal']
                
                if c_sem == 0:
                    return 100.0 if comp_sem > 0 else 0.0
                
                variacion = ((comp_sem - c_sem) / c_sem) * 100.0
                return round(variacion, 2)

            df_final['Var. Compra vs Consumo (%)'] = df_final.apply(calcular_var, axis=1)

            def evaluar_semaforo(row):
                p1 = row['Calculo Exacto']
                l_min = pd.to_numeric(row['Lot_Min'], errors='coerce')
                l_max = pd.to_numeric(row['Lot_Max'], errors='coerce')
                
                if p1 == 0: return "🟢 Stock Suficiente"
                if pd.notna(l_min) and p1 < l_min: return f"🔴 Menor a Mín. ({l_min:,.1f})"
                if pd.notna(l_max) and p1 > l_max: return f"🟡 Mayor a Máx. ({l_max:,.1f})"
                return "🟢 Rango Tradicional"

            df_final['Semáforo / Alerta'] = df_final.apply(evaluar_semaforo, axis=1)

            # SE INCLUYÓ LA COLUMNA 'Compra_Semanal' EN LA LISTA DE COLUMNAS FINALES
            columnas_finales = [
                'Código', 'Descripción', 'Stock_Actual', 'Consumo_Semanal', 'Compra_Semanal',
                '🛒 A COMPRAR (Unidades)', 'Calculo Exacto', 'Var. Compra vs Consumo (%)', 'Semáforo / Alerta'
            ]
            df_vista_final = df_final[columnas_finales]

            st.markdown("##### 📋 Sugerencia de Compra Final")
            st.dataframe(df_vista_final, use_container_width=True, hide_index=True)

            st.download_button(
                label="📥 Descargar Reporte en Excel (.xlsx)",
                data=generar_excel_proyeccion(df_vista_final),
                file_name=f"Sugerencia_Compras_{date.today().strftime('%d_%m_%Y')}.xlsx",
                type="primary",
                use_container_width=True
            )

    with tab_recetas:
        st.subheader("Estructura de Recetas (BOM - Bill of Materials)")
        st.write("Gestión de explosión de ingredientes y costos unitarios de fabricación.")
        st.info("Pestaña lista para recibir la lógica de recetas en la siguiente etapa.")