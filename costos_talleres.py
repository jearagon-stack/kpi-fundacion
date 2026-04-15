import streamlit as st
import pandas as pd
import re
import io
from datetime import date

def extraer_ordenes(texto):
    if pd.isna(texto): return []
    return re.findall(r'\b\d{3,4}-\d{3,4}\b', str(texto))

def tiene_orden_valida(ordenes_extraidas, ordenes_sgt):
    for o in ordenes_extraidas:
        if o in ordenes_sgt: return True
    return False

def generar_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Partidas_TG')
    return output.getvalue()

def mostrar_modulo_costos():
    st.title("🖨️ Contabilidad de Costos - Talleres Gráficos")
    st.info("Sistema Automático de Costeo por Órdenes de Producción")

    tab_carga, tab_auditoria, tab_liquidacion = st.tabs([
        "📥 1. Carga de Datos", 
        "🕵️ 2. Auditoría (El Purgatorio)", 
        "💰 3. Liquidación y Partidas"
    ])

    # ==========================================
    # PESTAÑA 1: CARGA DE ARCHIVOS
    # ==========================================
    with tab_carga:
        st.subheader("Paso 1: Periodo, Costos y Archivos")
        
        st.markdown("**1. Definir Periodo a Costear**")
        col_m1, col_m2 = st.columns(2)
        with col_m1: mes_proceso = st.selectbox("Mes:", range(1, 13), index=date.today().month - 1)
        with col_m2: anio_proceso = st.number_input("Año:", min_value=2024, max_value=2030, value=date.today().year)

        st.markdown("**2. Costo de Mano de Obra**")
        costo_planilla = st.number_input(
            "💵 Ingresa el Costo Total de Planilla + Horas Extras del Mes ($):", 
            min_value=0.0, value=0.0, step=100.0
        )
        
        st.markdown("---")
        st.markdown("**3. Carga de Archivos Base**")
        
        col1, col2 = st.columns(2)
        with col1:
            arch_sgt = st.file_uploader("1. Maestro de Órdenes (SGT_TG)", type=["xlsx"])
            arch_tras_mp = st.file_uploader("2. Traslados Materia Prima", type=["xlsx"], accept_multiple_files=True)
            arch_tiempos = st.file_uploader("3. Reporte de Tiempos", type=["xlsx"])
            
        with col2:
            arch_fact = st.file_uploader("4. Facturación del Mes", type=["xlsx"])
            arch_tras_pt = st.file_uploader("5. Traslados Internos PT", type=["xlsx"], accept_multiple_files=True)

        if arch_sgt and arch_fact and arch_tras_mp and arch_tiempos:
            if st.button("🔍 Escanear Archivos y Construir Bodega Virtual", type="primary", use_container_width=True):
                mes_ya_cerrado = False 
                
                if mes_ya_cerrado:
                    st.error(f"🚨 ¡ALERTA! El mes {mes_proceso}/{anio_proceso} ya tiene traslados guardados.")
                else:
                    with st.spinner("Aplicando Triple Filtro SGT a los documentos..."):
                        try:
                            # 1. MAESTRO SGT
                            df_sgt = pd.read_excel(arch_sgt, dtype=str)
                            ordenes_validas = df_sgt['Orden'].dropna().astype(str).str.strip().tolist() if 'Orden' in df_sgt.columns else []

                            # 2. FACTURACIÓN
                            df_fact = pd.read_excel(arch_fact, dtype=str)
                            if 'Descripcion' in df_fact.columns:
                                df_fact['Ordenes_Detectadas'] = df_fact['Descripcion'].apply(extraer_ordenes)
                                def clasificar_factura(row):
                                    desc = str(row.get('Descripcion', '')).upper()
                                    cat = str(row.get('Categoria', '')).upper()
                                    if tiene_orden_valida(row['Ordenes_Detectadas'], ordenes_validas): return "Orden Lista"
                                    if "SERVICIO" in cat or "SERVICIO" in desc: return "Servicios"
                                    if any(k in desc for k in ["BANNER", "AFICHE", "CALENDARIO", "ROTULO"]): return "Venta Directa"
                                    if any(k in desc for k in ["RECICLAJE", "DESPERDICIO"]): return "Reciclaje"
                                    return "Huérfana (Revisar)"
                                df_fact['Clasificacion'] = df_fact.apply(clasificar_factura, axis=1)

                            # 3. TIEMPOS (MANO DE OBRA)
                            df_tiempos = pd.read_excel(arch_tiempos, dtype=str)
                            if 'Observaciones' in df_tiempos.columns:
                                df_tiempos['Ordenes_Detectadas'] = df_tiempos['Observaciones'].apply(extraer_ordenes)
                                df_tiempos['Clasificacion'] = df_tiempos['Ordenes_Detectadas'].apply(
                                    lambda x: "Orden Lista" if tiene_orden_valida(x, ordenes_validas) else "Huérfana (Revisar)"
                                )

                            # 4. TRASLADOS MATERIA PRIMA (Múltiples archivos)
                            dfs_mp = [pd.read_excel(f, dtype=str) for f in arch_tras_mp]
                            df_mp = pd.concat(dfs_mp, ignore_index=True) if dfs_mp else pd.DataFrame()
                            if not df_mp.empty and 'Descripcion' in df_mp.columns:
                                df_mp['Ordenes_Detectadas'] = df_mp['Descripcion'].apply(extraer_ordenes)
                                def clasificar_traslado(row):
                                    cat = str(row.get('Categoria', '')).upper()
                                    if tiene_orden_valida(row['Ordenes_Detectadas'], ordenes_validas): return "Orden Lista"
                                    # Excepción para Costos Indirectos
                                    if any(k in cat for k in ["EMPAQUE", "LIMPIEZA", "REPUESTO", "REPUESTOS"]): return "Costo Indirecto (Válido)"
                                    return "Huérfana (Revisar)"
                                df_mp['Clasificacion'] = df_mp.apply(clasificar_traslado, axis=1)

                            # GUARDAR EN MEMORIA
                            st.session_state['tg_fact'] = df_fact
                            st.session_state['tg_tiempos'] = df_tiempos
                            st.session_state['tg_mp'] = df_mp
                            st.session_state['tg_sgt'] = df_sgt
                            st.session_state['tg_ordenes_validas'] = ordenes_validas
                            st.session_state['tg_costo_planilla'] = costo_planilla
                            st.session_state['tg_datos_cargados'] = True
                            st.session_state['fase2_aprobada'] = False
                            
                            st.success("✅ Filtros SGT aplicados. Pasa a la pestaña '2. Auditoría'.")
                        except Exception as e:
                            st.error(f"Error al leer archivos: Verifica que no haya hojas vacías. Detalle: {e}")

    # ==========================================
    # PESTAÑA 2: AUDITORÍA (TRIPLE PURGATORIO)
    # ==========================================
    with tab_auditoria:
        st.subheader("Sala de Espera: Revisión de Anomalías")
        if st.session_state.get('tg_datos_cargados', False):
            df_fact = st.session_state['tg_fact']
            df_tiempos = st.session_state['tg_tiempos']
            df_mp = st.session_state['tg_mp']
            ordenes_validas = st.session_state.get('tg_ordenes_validas', [])
            
            # Conteo de huérfanas
            h_fact = len(df_fact[df_fact['Clasificacion'] == "Huérfana (Revisar)"])
            h_tiempos = len(df_tiempos[df_tiempos['Clasificacion'] == "Huérfana (Revisar)"])
            h_mp = len(df_mp[df_mp['Clasificacion'] == "Huérfana (Revisar)"]) if not df_mp.empty else 0
            
            total_huerfanas = h_fact + h_tiempos + h_mp

            if total_huerfanas > 0:
                st.error(f"🚨 Bloqueo activo: Tienes {total_huerfanas} registros huérfanos que no coinciden con SGT.")
                
                # PURGATORIO 1: FACTURACIÓN
                if h_fact > 0:
                    with st.expander(f"🧾 Facturas Huérfanas ({h_fact})", expanded=True):
                        df_h_f = df_fact[df_fact['Clasificacion'] == "Huérfana (Revisar)"][['Numero', 'Descripcion', 'VentaNeta']].copy()
                        df_h_f['Accion'] = "Pendiente"
                        st.data_editor(df_h_f, column_config={"Accion": st.column_config.SelectboxColumn("Acción", options=["Pendiente", "Omitir"])}, hide_index=True, key="ed_fact")

                # PURGATORIO 2: TIEMPOS
                if h_tiempos > 0:
                    with st.expander(f"⏱️ Horas de Planilla Huérfanas ({h_tiempos})", expanded=True):
                        df_h_t = df_tiempos[df_tiempos['Clasificacion'] == "Huérfana (Revisar)"][['Empleado', 'Observaciones']].copy()
                        df_h_t['Accion'] = "Pendiente"
                        st.data_editor(df_h_t, column_config={"Accion": st.column_config.SelectboxColumn("Acción", options=["Pendiente", "Omitir"])}, hide_index=True, key="ed_tiempos")

                # PURGATORIO 3: MATERIA PRIMA
                if h_mp > 0:
                    with st.expander(f"📦 Traslados Materia Prima Huérfanos ({h_mp})", expanded=True):
                        df_h_m = df_mp[df_mp['Clasificacion'] == "Huérfana (Revisar)"][['Numero', 'Descripcion', 'Categoria']].copy()
                        df_h_m['Accion'] = "Pendiente"
                        st.data_editor(df_h_m, column_config={"Accion": st.column_config.SelectboxColumn("Acción", options=["Pendiente", "Omitir"])}, hide_index=True, key="ed_mp")

                if st.button("💾 Omitir Pendientes y Forzar Validación", type="primary"):
                    st.success("✅ Purgatorio evadido por el usuario. Proceda a Liquidación.")
                    st.session_state['fase2_aprobada'] = True
            else:
                st.success("✅ Validación completada. 100% de coincidencia con SGT y Costos Indirectos.")
                st.session_state['fase2_aprobada'] = True
        else:
            st.write("Carga los archivos en la pestaña 1.")

    # ==========================================
    # PESTAÑA 3: LIQUIDACIÓN (LA CALCULADORA)
    # ==========================================
    with tab_liquidacion:
        st.subheader("💰 Liquidación y Prorrateo")
        if st.session_state.get('fase2_aprobada', False):
            st.info("🟢 El sistema está listo para calcular los costos del mes.")
            
            if st.button("🚀 Ejecutar Prorrateo y Generar Partidas", type="primary"):
                with st.spinner("Calculando mano de obra..."):
                    df_tiempos = st.session_state['tg_tiempos']
                    costo_total = st.session_state['tg_costo_planilla']
                    
                    st.markdown("### 📊 Resultado del Prorrateo de Mano de Obra")
                    
                    # Buscamos la columna de horas dinámicamente para evitar errores (TotalHoras, Total Hora, etc)
                    col_horas = next((c for c in df_tiempos.columns if 'TOTALHORA' in c.upper().replace(' ', '')), None)
                    
                    if col_horas:
                        df_tiempos[col_horas] = pd.to_numeric(df_tiempos[col_horas], errors='coerce').fillna(0)
                        horas_totales_validas = df_tiempos[df_tiempos['Clasificacion'] == 'Orden Lista'][col_horas].sum()
                        
                        if horas_totales_validas > 0:
                            costo_por_hora = costo_total / horas_totales_validas
                            
                            col1, col2, col3 = st.columns(3)
                            col1.metric("Costo Total Planilla", f"${costo_total:,.2f}")
                            col2.metric("Horas Válidas Detectadas", f"{horas_totales_validas:,.2f} hrs")
                            col3.metric("Costo Asignado por Hora", f"${costo_por_hora:,.2f}/hr")
                            
                            df_partidas = pd.DataFrame({
                                "Fecha": [date.today().strftime("%d/%m/%Y")],
                                "Cuenta": ["PRODUCTO EN PROCESO - MANO DE OBRA"],
                                "Debe": [costo_total],
                                "Haber": [0.0],
                                "Concepto": ["Traslado de costo de nómina a proceso productivo"]
                            })
                            
                            st.download_button(
                                label="⬇️ Descargar Excel de Partidas",
                                data=generar_excel_bytes(df_partidas),
                                file_name=f"Partidas_Talleres_{mes_proceso}_{anio_proceso}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            st.balloons()
                        else:
                            st.warning("Se encontró la columna, pero la suma de horas válidas es 0. Revisa el Purgatorio.")
                    else:
                        st.error("❌ No se encontró ninguna columna llamada 'TotalHoras' en el archivo de tiempos.")

        else:
            st.warning("🛑 Debes completar la pestaña '2. Auditoría' primero.")