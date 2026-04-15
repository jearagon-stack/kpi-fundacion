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
        
        col_m1, col_m2 = st.columns(2)
        with col_m1: mes_proceso = st.selectbox("Mes:", range(1, 13), index=date.today().month - 1)
        with col_m2: anio_proceso = st.number_input("Año:", min_value=2024, max_value=2030, value=date.today().year)

        costo_planilla = st.number_input(
            "💵 Ingresa el Costo Total de Planilla + Horas Extras del Mes ($):", 
            min_value=0.0, value=0.0, step=100.0
        )
        
        st.markdown("---")
        
        col1, col2 = st.columns(2)
        with col1:
            arch_sgt = st.file_uploader("1. Maestro de Órdenes (SGT_TG)", type=["xlsx"])
            arch_tras_mp = st.file_uploader("2. Traslados Materia Prima", type=["xlsx"], accept_multiple_files=True)
            arch_tiempos = st.file_uploader("3. Reporte de Tiempos", type=["xlsx"])
            
        with col2:
            arch_fact = st.file_uploader("4. Facturación del Mes", type=["xlsx"])
            arch_tras_pt = st.file_uploader("5. Traslados Internos PT", type=["xlsx"], accept_multiple_files=True)

        if arch_sgt and arch_fact and arch_tras_mp and arch_tiempos:
            if st.button("🔍 Escanear Archivos y Aplicar Filtros", type="primary", use_container_width=True):
                with st.spinner("Leyendo y cruzando datos..."):
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

                        # 3. TIEMPOS (MANO DE OBRA) - IGNORANDO VACÍOS
                        df_tiempos = pd.read_excel(arch_tiempos, dtype=str)
                        if 'Observaciones' in df_tiempos.columns:
                            df_tiempos['Ordenes_Detectadas'] = df_tiempos['Observaciones'].apply(extraer_ordenes)
                            def clasificar_tiempos(row):
                                obs = str(row.get('Observaciones', '')).strip()
                                if obs in ['', 'nan', 'None', '--', '-'] or pd.isna(row.get('Observaciones')): 
                                    return "Omitido Automático"
                                if tiene_orden_valida(row['Ordenes_Detectadas'], ordenes_validas): 
                                    return "Orden Lista"
                                return "Huérfana (Revisar)"
                            df_tiempos['Clasificacion'] = df_tiempos.apply(clasificar_tiempos, axis=1)

                        # 4. TRASLADOS MATERIA PRIMA (Sin filtro automático de indirectos)
                        dfs_mp = [pd.read_excel(f, dtype=str) for f in arch_tras_mp]
                        df_mp = pd.concat(dfs_mp, ignore_index=True) if dfs_mp else pd.DataFrame()
                        
                        col_texto_mp = 'Concepto' if 'Concepto' in df_mp.columns else 'Descripcion'
                        
                        if not df_mp.empty and col_texto_mp in df_mp.columns:
                            df_mp['Ordenes_Detectadas'] = df_mp[col_texto_mp].apply(extraer_ordenes)
                            def clasificar_traslado(row):
                                if tiene_orden_valida(row['Ordenes_Detectadas'], ordenes_validas): return "Orden Lista"
                                return "Huérfana (Revisar)"
                            df_mp['Clasificacion'] = df_mp.apply(clasificar_traslado, axis=1)

                        st.session_state['tg_fact'] = df_fact
                        st.session_state['tg_tiempos'] = df_tiempos
                        st.session_state['tg_mp'] = df_mp
                        st.session_state['tg_ordenes_validas'] = ordenes_validas
                        st.session_state['tg_costo_planilla'] = costo_planilla
                        st.session_state['tg_datos_cargados'] = True
                        st.session_state['fase2_aprobada'] = False
                        
                        st.success("✅ Datos cruzados correctamente. Avanza al Purgatorio.")
                    except Exception as e:
                        st.error(f"Error al leer archivos: {e}")

    # ==========================================
    # PESTAÑA 2: AUDITORÍA (EL PURGATORIO INTERACTIVO)
    # ==========================================
    with tab_auditoria:
        st.subheader("🕵️ Sala de Espera: Revisión Manual")
        if st.session_state.get('tg_datos_cargados', False):
            df_fact = st.session_state['tg_fact']
            df_tiempos = st.session_state['tg_tiempos']
            df_mp = st.session_state['tg_mp']
            ordenes_validas = st.session_state.get('tg_ordenes_validas', [])
            
            h_fact = len(df_fact[df_fact['Clasificacion'] == "Huérfana (Revisar)"])
            h_tiempos = len(df_tiempos[df_tiempos['Clasificacion'] == "Huérfana (Revisar)"])
            h_mp = len(df_mp[df_mp['Clasificacion'] == "Huérfana (Revisar)"]) if not df_mp.empty else 0
            total_huerfanas = h_fact + h_tiempos + h_mp

            if total_huerfanas > 0:
                st.error(f"🚨 Tienes {total_huerfanas} registros que necesitan tu decisión.")
                
                opciones_accion = ["Pendiente", "Asignar Orden", "Costo Indirecto", "Omitir"]
                config_columnas = {
                    "Accion": st.column_config.SelectboxColumn("Acción", options=opciones_accion, required=True),
                    "Orden_SGT": st.column_config.TextColumn("Código Orden (Si aplica)")
                }

                ed_fact = ed_tiempos = ed_mp = pd.DataFrame()

                if h_fact > 0:
                    with st.expander(f"🧾 Facturas Huérfanas ({h_fact})", expanded=True):
                        df_h_f = df_fact[df_fact['Clasificacion'] == "Huérfana (Revisar)"][['Numero', 'Descripcion', 'VentaNeta']].copy()
                        df_h_f['Accion'] = "Pendiente"
                        df_h_f['Orden_SGT'] = ""
                        ed_fact = st.data_editor(df_h_f, column_config=config_columnas, hide_index=True, key="ed_fact")

                if h_tiempos > 0:
                    with st.expander(f"⏱️ Horas Huérfanas ({h_tiempos})", expanded=True):
                        df_h_t = df_tiempos[df_tiempos['Clasificacion'] == "Huérfana (Revisar)"][['Empleado', 'Observaciones']].copy()
                        df_h_t['Accion'] = "Pendiente"
                        df_h_t['Orden_SGT'] = ""
                        ed_tiempos = st.data_editor(df_h_t, column_config=config_columnas, hide_index=True, key="ed_tiempos")

                if h_mp > 0:
                    with st.expander(f"📦 Traslados MP Huérfanos ({h_mp})", expanded=True):
                        col_mostrar_mp = 'Concepto' if 'Concepto' in df_mp.columns else 'Descripcion'
                        df_h_m = df_mp[df_mp['Clasificacion'] == "Huérfana (Revisar)"][['Numero', col_mostrar_mp, 'Categoria']].copy()
                        df_h_m['Accion'] = "Pendiente"
                        df_h_m['Orden_SGT'] = ""
                        ed_mp = st.data_editor(df_h_m, column_config=config_columnas, hide_index=True, key="ed_mp")

                if st.button("💾 Guardar Decisiones y Validar", type="primary"):
                    errores = []
                    
                    def revisar_tabla(df_editado, nombre_tabla):
                        for i, row in df_editado.iterrows():
                            if row['Accion'] == "Pendiente":
                                errores.append(f"Queda un registro pendiente en {nombre_tabla}.")
                            elif row['Accion'] == "Asignar Orden":
                                orden = str(row['Orden_SGT']).strip()
                                if orden not in ordenes_validas:
                                    errores.append(f"La orden '{orden}' asignada en {nombre_tabla} NO existe en SGT.")

                    if not ed_fact.empty: revisar_tabla(ed_fact, "Facturas")
                    if not ed_tiempos.empty: revisar_tabla(ed_tiempos, "Tiempos")
                    if not ed_mp.empty: revisar_tabla(ed_mp, "Traslados MP")

                    if errores:
                        st.error("❌ Corrige los siguientes errores antes de pasar a liquidar:")
                        for e in errores: st.write(f"- {e}")
                    else:
                        st.success("✅ ¡Purgatorio limpio y validado! Procede a la Liquidación.")
                        st.session_state['fase2_aprobada'] = True
            else:
                st.success("✅ Todo está perfecto. No hay huérfanas.")
                st.session_state['fase2_aprobada'] = True
        else:
            st.write("Carga los archivos en la Pestaña 1.")

    # ==========================================
    # PESTAÑA 3: LIQUIDACIÓN (LA CALCULADORA)
    # ==========================================
    with tab_liquidacion:
        st.subheader("💰 Liquidación y Prorrateo")
        if st.session_state.get('fase2_aprobada', False):
            st.info("🟢 Listo. Nota: Si una fila tiene múltiples órdenes, el sistema dividirá el costo/horas en partes iguales automáticamente.")
            
            if st.button("🚀 Ejecutar Prorrateo y Generar Partidas", type="primary"):
                df_tiempos = st.session_state['tg_tiempos']
                costo_total = st.session_state['tg_costo_planilla']
                
                col_horas = next((c for c in df_tiempos.columns if 'TOTALHORA' in c.upper().replace(' ', '')), None)
                
                if col_horas:
                    df_tiempos[col_horas] = pd.to_numeric(df_tiempos[col_horas], errors='coerce').fillna(0)
                    horas_totales_validas = df_tiempos[df_tiempos['Clasificacion'] == 'Orden Lista'][col_horas].sum()
                    
                    if horas_totales_validas > 0:
                        costo_por_hora = costo_total / horas_totales_validas
                        col1, col2, col3 = st.columns(3)
                        col1.metric("Costo Planilla", f"${costo_total:,.2f}")
                        col2.metric("Horas Válidas", f"{horas_totales_validas:,.2f} hrs")
                        col3.metric("Costo por Hora", f"${costo_por_hora:,.2f}/hr")
                        
                        df_partidas = pd.DataFrame({
                            "Cuenta": ["PRODUCTO EN PROCESO - MANO DE OBRA"],
                            "Debe": [costo_total], "Haber": [0.0],
                            "Concepto": ["Traslado de costo de nómina"]
                        })
                        
                        st.download_button("⬇️ Descargar Partidas", data=generar_excel_bytes(df_partidas), file_name="Partidas_TG.xlsx")
                    else:
                        st.warning("No hay horas válidas. Verifica el reporte.")
                else:
                    st.error("No se encontró la columna de horas.")
        else:
            st.warning("🛑 Debes completar el Purgatorio primero.")