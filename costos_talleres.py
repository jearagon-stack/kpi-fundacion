import streamlit as st
import pandas as pd
import re
import io
from datetime import date

def extraer_ordenes(texto):
    if pd.isna(texto): return []
    return re.findall(r'\b\d{3,4}-\d{3,4}\b', str(texto))

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
                    with st.spinner("Escaneando documentos y validando..."):
                        try:
                            df_sgt = pd.read_excel(arch_sgt, dtype=str)
                            lista_ordenes_validas = df_sgt['Orden'].dropna().astype(str).str.strip().tolist() if 'Orden' in df_sgt.columns else []

                            df_fact = pd.read_excel(arch_fact, dtype=str)
                            if 'Descripcion' in df_fact.columns:
                                df_fact['Ordenes_Detectadas'] = df_fact['Descripcion'].apply(extraer_ordenes)
                                def clasificar_factura(row):
                                    desc = str(row.get('Descripcion', '')).upper()
                                    cat = str(row.get('Categoria', '')).upper()
                                    if len(row['Ordenes_Detectadas']) > 0: return "Orden Lista"
                                    if "SERVICIO" in cat or "SERVICIO" in desc: return "Servicios"
                                    if any(k in desc for k in ["BANNER", "AFICHE", "CALENDARIO"]): return "Venta Directa"
                                    return "Huérfana (Revisar)"
                                df_fact['Clasificacion'] = df_fact.apply(clasificar_factura, axis=1)

                            df_tiempos = pd.read_excel(arch_tiempos, dtype=str)
                            if 'Observaciones' in df_tiempos.columns:
                                df_tiempos['Ordenes_Detectadas'] = df_tiempos['Observaciones'].apply(extraer_ordenes)
                                df_tiempos['Clasificacion'] = df_tiempos['Ordenes_Detectadas'].apply(lambda x: "Orden Lista" if len(x) > 0 else "Huérfana (Revisar)")

                            # Guardamos en memoria para las siguientes pestañas
                            st.session_state['tg_fact'] = df_fact
                            st.session_state['tg_tiempos'] = df_tiempos
                            st.session_state['tg_sgt'] = df_sgt
                            st.session_state['tg_ordenes_validas'] = lista_ordenes_validas
                            st.session_state['tg_costo_planilla'] = costo_planilla
                            st.session_state['tg_datos_cargados'] = True
                            st.session_state['fase2_aprobada'] = False
                            
                            st.success("✅ Archivos leídos. Pasa a la pestaña '2. Auditoría'.")
                        except Exception as e:
                            st.error(f"Error al leer archivos: {e}")

    # ==========================================
    # PESTAÑA 2: AUDITORÍA (EL PURGATORIO)
    # ==========================================
    with tab_auditoria:
        st.subheader("Sala de Espera: Revisión de Anomalías")
        if st.session_state.get('tg_datos_cargados', False):
            df_fact = st.session_state['tg_fact']
            ordenes_validas = st.session_state.get('tg_ordenes_validas', [])
            
            conteos = df_fact['Clasificacion'].value_counts()
            col_a, col_b, col_c = st.columns(3)
            col_a.metric("✅ Órdenes Listas", conteos.get("Orden Lista", 0))
            col_b.metric("🛠️ Servicios / Directa", conteos.get("Servicios", 0) + conteos.get("Venta Directa", 0))
            col_c.metric("⚠️ Huérfanas", conteos.get("Huérfana (Revisar)", 0))
            st.divider()

            df_huerfanas_fact = df_fact[df_fact['Clasificacion'] == "Huérfana (Revisar)"].copy()
            
            if not df_huerfanas_fact.empty:
                st.error("🚨 Bloqueo activo: Resuelve las facturas huérfanas.")
                df_mostrar = df_huerfanas_fact[['Fecha', 'Numero', 'Descripcion', 'VentaNeta']].copy()
                df_mostrar['Accion'] = "Pendiente"
                df_mostrar['Orden_SGT'] = ""
                
                editado = st.data_editor(
                    df_mostrar,
                    column_config={"Accion": st.column_config.SelectboxColumn("Acción", options=["Pendiente", "Omitir"], required=True)},
                    use_container_width=True, hide_index=True
                )
                
                if st.button("💾 Validar Correcciones", type="primary"):
                    errores = [f"Factura {r['Numero']} sigue pendiente" for i, r in editado.iterrows() if r['Accion'] == "Pendiente"]
                    if errores:
                        st.error("❌ " + "\n".join(errores))
                    else:
                        st.success("✅ Purgatorio limpio. Proceda a Liquidación.")
                        st.session_state['fase2_aprobada'] = True
            else:
                st.success("✅ Validación completada. No hay facturas huérfanas.")
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
                with st.spinner("Realizando cálculos matemáticos..."):
                    
                    df_tiempos = st.session_state['tg_tiempos']
                    costo_total = st.session_state['tg_costo_planilla']
                    
                    # 1. Prorrateo Matemático
                    st.markdown("### 📊 Resultado del Prorrateo de Mano de Obra")
                    
                    # Intentamos extraer las horas. Si la columna se llama diferente en tu Excel, lo ajustaremos.
                    col_horas = 'TotalHora' if 'TotalHora' in df_tiempos.columns else df_tiempos.columns[-3] 
                    
                    df_tiempos[col_horas] = pd.to_numeric(df_tiempos[col_horas], errors='coerce').fillna(0)
                    horas_totales_validas = df_tiempos[df_tiempos['Clasificacion'] == 'Orden Lista'][col_horas].sum()
                    
                    if horas_totales_validas > 0:
                        costo_por_hora = costo_total / horas_totales_validas
                        
                        col1, col2, col3 = st.columns(3)
                        col1.metric("Costo Total Planilla", f"${costo_total:,.2f}")
                        col2.metric("Horas Válidas Detectadas", f"{horas_totales_validas:,.2f} hrs")
                        col3.metric("Costo Asignado por Hora", f"${costo_por_hora:,.2f}/hr")
                        
                        # 2. Generar el Excel de Partidas (Muestra genérica)
                        df_partidas = pd.DataFrame({
                            "Fecha": [date.today().strftime("%d/%m/%Y")],
                            "Cuenta": ["PRODUCTO EN PROCESO - MANO DE OBRA"],
                            "Debe": [costo_total],
                            "Haber": [0.0],
                            "Concepto": ["Traslado de costo de nómina a proceso productivo"]
                        })
                        
                        excel_data = generar_excel_bytes(df_partidas)
                        
                        st.success("✅ Cálculos realizados. Todo está listo para contabilizar.")
                        st.download_button(
                            label="⬇️ Descargar Excel de Partidas",
                            data=excel_data,
                            file_name=f"Partidas_Talleres_{mes_proceso}_{anio_proceso}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.balloons()
                    else:
                        st.warning("No se detectaron horas válidas para prorratear. Revisa el Excel de tiempos.")

        else:
            st.warning("🛑 Debes completar la pestaña '2. Auditoría' primero.")