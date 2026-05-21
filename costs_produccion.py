import streamlit as st
import pandas as pd
import io
import re
from datetime import date

# ==========================================
# FUNCIONES AUXILIARES DE INTELIGENCIA DE DATOS
# ==========================================
def detectar_columna(df, palabras_clave):
    """Busca de forma inteligente una columna basada en palabras clave"""
    for col in df.columns:
        col_str = str(col).upper().strip()
        if any(p in col_str for p in palabras_clave):
            return col
    return df.columns[0] if len(df.columns) > 0 else None

def limpiar_codigo(c):
    """Estandariza los códigos de producto para cruces exactos"""
    if pd.isna(c): return ""
    val = str(c).strip().upper()
    if val.replace('.', '', 1).isdigit():
        return str(int(float(val)))
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
    # PESTAÑA 1: PROYECCIÓN DE COMPRAS (ETAPA 1)
    # ==========================================
    with tab_proyeccion:
        st.subheader("Simulador y Proyección de Compras de Materia Prima")
        st.write("Carga los archivos base del sistema para calcular las necesidades reales frente al comportamiento de compras.")
        
        # Bloque de Parámetros de Tiempo Ajustables Manualmente
        st.markdown("##### ⏱️ Configuración de Horizontes de Tiempo")
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            dias_historial = st.number_input("Días cubiertos por el Historial Cargado (Ej: 90 para 3 meses):", min_value=1, value=90, key="zip_dias_hist")
        with col_t2:
            dias_proyectar = st.number_input("Días a Proyectar para la Compra (Ej: 30 para 1 mes):", min_value=1, value=30, key="zip_dias_proj")

        st.markdown("---")
        st.markdown("##### 📥 Carga de Archivos Maestros")
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            arch_stock = st.file_uploader("1. Inventario / Stock Actual Saneado", type=["xlsx", "xls"], key="zip_f_stock")
            arch_salidas = st.file_uploader("2. Historial de Salidas de Bodega (Consumos)", type=["xlsx", "xls"], key="zip_f_salidas")
        with col_f2:
            arch_compras = st.file_uploader("3. Historial de Compras (Facturas de Proveedores)", type=["xlsx", "xls"], key="zip_f_compras")

        if arch_stock and arch_salidas and arch_compras:
            if st.button("🚀 Ejecutar Análisis de Demanda y Variación", type="primary", use_container_width=True):
                with st.spinner("Procesando, cruzando y normalizando información..."):
                    try:
                        # 1. Leer Archivos
                        df_s_raw = pd.read_excel(arch_stock, dtype=str)
                        df_out_raw = pd.read_excel(arch_salidas, dtype=str)
                        df_in_raw = pd.read_excel(arch_compras, dtype=str)
                        
                        df_s_raw.columns = df_s_raw.columns.str.strip()
                        df_out_raw.columns = df_out_raw.columns.str.strip()
                        df_in_raw.columns = df_in_raw.columns.str.strip()

                        # 2. Detección Inteligente de Columnas Clave
                        col_cod_s = detectar_columna(df_s_raw, ['COD', 'ID', 'ARTICULO', 'PRODUCTO', 'ITEM'])
                        col_stk_s = detectar_columna(df_s_raw, ['STOCK', 'EXIST', 'ACTUAL', 'CANTIDAD', 'DISPO'])
                        col_nom_s = detectar_columna(df_s_raw, ['NOM', 'DESC', 'DETALLE', 'PRODUCTO'])

                        col_cod_out = detectar_columna(df_out_raw, ['COD', 'ID', 'ARTICULO', 'PRODUCTO', 'ITEM'])
                        col_cant_out = detectar_columna(df_out_raw, ['CANT', 'TOTAL', 'MOV', 'SALIDA', 'VOLUMEN'])

                        col_cod_in = detectar_columna(df_in_raw, ['COD', 'ID', 'ARTICULO', 'PRODUCTO', 'ITEM'])
                        col_cant_in = detectar_columna(df_in_raw, ['CANT', 'TOTAL', 'COMPRA', 'VOLUMEN'])

                        # 3. Limpieza y Homologación de Datos
                        df_s_raw['Cod_Clean'] = df_s_raw[col_cod_s].apply(limpiar_codigo)
                        df_s_raw['Stock_Num'] = pd.to_numeric(df_s_raw[col_stk_s], errors='coerce').fillna(0)
                        
                        df_out_raw['Cod_Clean'] = df_out_raw[col_cod_out].apply(limpiar_codigo)
                        df_out_raw['Cant_Num'] = pd.to_numeric(df_out_raw[col_cant_out], errors='coerce').fillna(0)
                        
                        df_in_raw['Cod_Clean'] = df_in_raw[col_cod_in].apply(limpiar_codigo)
                        df_in_raw['Cant_Num'] = pd.to_numeric(df_in_raw[col_cant_in], errors='coerce').fillna(0)

                        # 4. Agrupaciones por Código de Producto
                        df_stock_group = df_s_raw.groupby('Cod_Clean').agg({
                            'Stock_Num': 'sum',
                            col_nom_s: 'first'
                        }).reset_index().rename(columns={col_nom_s: 'Descripción', 'Stock_Num': 'Stock_Actual'})

                        df_salidas_group = df_out_raw.groupby('Cod_Clean')['Cant_Num'].sum().reset_index().rename(columns={'Cant_Num': 'Consumo_Historico'})
                        df_compras_group = df_in_raw.groupby('Cod_Clean')['Cant_Num'].sum().reset_index().rename(columns={'Cant_Num': 'Compras_Historicas'})

                        # 5. Cálculo de Lotes Mínimos y Máximos desde el Histórico de Compras (Dato sólido)
                        df_min_lot = df_in_raw[df_in_raw['Cant_Num'] > 0].groupby('Cod_Clean')['Cant_Num'].min().reset_index().rename(columns={'Cant_Num': 'Lot_Min'})
                        df_max_lot = df_in_raw.groupby('Cod_Clean')['Cant_Num'].max().reset_index().rename(columns={'Cant_Num': 'Lot_Max'})

                        # 6. Consolidación en un Universo Maestro de Códigos
                        todos_los_codigos = set(df_stock_group['Cod_Clean']).union(set(df_salidas_group['Cod_Clean'])).union(set(df_compras_group['Cod_Clean']))
                        df_maestro = pd.DataFrame({'Código': list(todos_los_codigos)})
                        df_maestro = df_maestro[df_maestro['Código'] != ""]

                        df_maestro = df_maestro.merge(df_stock_group, left_on='Código', right_on='Cod_Clean', how='left').drop(columns=['Cod_Clean'])
                        df_maestro = df_maestro.merge(df_salidas_group, left_on='Código', right_on='Cod_Clean', how='left').drop(columns=['Cod_Clean'])
                        df_maestro = df_maestro.merge(df_compras_group, left_on='Código', right_on='Cod_Clean', how='left').drop(columns=['Cod_Clean'])
                        df_maestro = df_maestro.merge(df_min_lot, left_on='Código', right_on='Cod_Clean', how='left').drop(columns=['Cod_Clean'])
                        df_maestro = df_maestro.merge(df_max_lot, left_on='Código', right_on='Cod_Clean', how='left').drop(columns=['Cod_Clean'])

                        df_maestro['Descripción'] = df_maestro['Descripción'].fillna('PRODUCTO NUEVOS SIN NOMBRE')
                        df_maestro['Stock_Actual'] = df_maestro['Stock_Actual'].fillna(0.0)
                        df_maestro['Consumo_Historico'] = df_maestro['Consumo_Historico'].fillna(0.0)
                        df_maestro['Compras_Historicas'] = df_maestro['Compras_Historicas'].fillna(0.0)
                        df_maestro['Stock_Seguridad_Pct'] = 5.0  # Parámetro base por defecto

                        st.session_state['zip_df_calculo_base'] = df_maestro
                        st.session_state['zip_ejecutado'] = True
                        st.success("✅ Datos unificados y listos para la simulación interactiva.")

                    except Exception as e:
                        st.error(f"Error procesando la información de compras: {e}")

        # Sección del Simulador Interactivo si ya fue ejecutado
        if st.session_state.get('zip_ejecutado', False):
            st.markdown("---")
            st.subheader("🎛️ Panel de Simulación y Selección de Stock de Seguridad")
            st.write("Modifica el porcentaje de seguridad de cada producto directamente en la tabla si lo consideras necesario:")

            df_base = st.session_state['zip_df_calculo_base']

            df_interactivo = st.data_editor(
                df_base[['Código', 'Descripción', 'Stock_Actual', 'Consumo_Historico', 'Compras_Historicas', 'Stock_Seguridad_Pct', 'Lot_Min', 'Lot_Max']],
                column_config={
                    "Código": st.column_config.TextColumn("Código", disabled=True),
                    "Descripción": st.column_config.TextColumn("Descripción", disabled=True),
                    "Stock_Actual": st.column_config.NumberColumn("Stock Actual", disabled=True, format="%.2f"),
                    "Consumo_Historico": st.column_config.NumberColumn("Consumo Hist.", disabled=True, format="%.2f"),
                    "Compras_Historicas": st.column_config.NumberColumn("Compras Hist.", disabled=True, format="%.2f"),
                    "Stock_Seguridad_Pct": st.column_config.NumberColumn("Seguridad (%)", min_value=0.0, max_value=100.0, format="%.1f"),
                    "Lot_Min": st.column_config.NumberColumn("Min Compra Lot", disabled=True, format="%.2f"),
                    "Lot_Max": st.column_config.NumberColumn("Max Compra Lot", disabled=True, format="%.2f")
                },
                hide_index=True,
                use_container_width=True,
                key="editor_interactivo_zip"
            )

            # CÁLCULOS MATEMÁTICOS DE OPCIONES EN TIEMPO REAL
            df_calculado = df_interactivo.copy()
            
            # Proyección 1: Consumo Real Proyectado
            df_calculado['Consumo_Diario'] = df_calculado['Consumo_Historico'] / dias_historial
            df_calculado['Consumo_Proyectado'] = df_calculado['Consumo_Diario'] * dias_proyectar
            df_calculado['Monto_Seguridad'] = df_calculado['Consumo_Proyectado'] * (df_calculado['Stock_Seguridad_Pct'] / 100.0)
            
            df_calculado['Proyección 1 (Consumo Real)'] = (df_calculado['Consumo_Proyectado'] + df_calculado['Monto_Seguridad']) - df_calculado['Stock_Actual']
            df_calculado['Proyección 1 (Consumo Real)'] = df_calculado['Proyección 1 (Consumo Real)'].apply(lambda x: max(0.0, round(x, 2)))

            # Proyección 2: Ritmo de Compras Histórico
            df_calculado['Compra_Diaria'] = df_calculado['Compras_Historicas'] / dias_historial
            df_calculado['Proyección 2 (Ritmo Compras)'] = (df_calculado['Compra_Diaria'] * dias_proyectar).round(2)

            # Variación Porcentual entre opciones
            def calcular_var(row):
                p1 = row['Proyección 1 (Consumo Real)']
                p2 = row['Proyección 2 (Ritmo Compras)']
                if p2 == 0:
                    return 100.0 if p1 > 0 else 0.0
                return round(((p1 - p2) / p2) * 100.0, 2)

            df_calculado['Variación (%)'] = df_calculado.apply(calcular_var, axis=1)

            # Lógica de Semáforo basada en el historial de lotes sólidos de compras
            def evaluar_semaforo(row):
                p1 = row['Proyección 1 (Consumo Real)']
                l_min = pd.to_numeric(row['Lot_Min'], errors='coerce')
                l_max = pd.to_numeric(row['Lot_Max'], errors='coerce')
                
                if p1 == 0:
                    return "🟢 Stock Suficiente"
                if pd.notna(l_min) and p1 < l_min:
                    return f"🔴 Menor a Mínimo Comprado ({l_min:,.1f})"
                if pd.notna(l_max) and p1 > l_max:
                    return f"🟡 Mayor a Máximo Comprado ({l_max:,.1f})"
                return "🟢 Rango Tradicional"

            df_calculado['Semáforo / Alerta'] = df_calculado.apply(evaluador_semaforo, axis=1)

            # Columnas limpias finales para mostrar y descargar
            columnas_finales = [
                'Código', 'Descripción', 'Stock_Actual', 'Proyección 1 (Consumo Real)', 
                'Proyección 2 (Ritmo Compras)', 'Variación (%)', 'Semáforo / Alerta'
            ]
            df_vista_final = df_calculado[columnas_finales]

            st.markdown("##### 📋 Resultados Finales Calculados")
            st.dataframe(df_vista_final, use_container_width=True, hide_index=True)

            # Botón de Descarga del Reporte en Excel Nativo (.xlsx) con columnas divididas
            st.download_button(
                label="📥 Descargar Reporte de Proyecciones en Excel (.xlsx)",
                data=generar_excel_proyeccion(df_vista_final),
                file_name=f"Proyeccion_Compras_{date.today().strftime('%d_%m_%Y')}.xlsx",
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
