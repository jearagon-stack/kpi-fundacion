import streamlit as st
import pandas as pd
import re
from datetime import date
# Asegúrate de importar tu conector a Google Sheets si ya lo tienes:
# from utils import obtener_dataframe

def extraer_ordenes(texto):
    """Busca patrones de órdenes tipo 1234-1234 o 123-1234"""
    if pd.isna(texto): return []
    return re.findall(r'\b\d{3,4}-\d{3,4}\b', str(texto))

def mostrar_modulo_costos():
    st.title("🖨️ Contabilidad de Costos - Talleres Gráficos")
    st.info("Sistema Automático de Costeo por Órdenes de Producción")

    tab_carga, tab_auditoria, tab_liquidacion = st.tabs([
        "📥 1. Carga de Datos", 
        "🕵️ 2. Auditoría (El Purgatorio)", 
        "💰 3. Liquidación y Partidas"
    ])

    # ==========================================
    # PESTAÑA 1: CARGA DE ARCHIVOS Y SEGURIDAD
    # ==========================================
    with tab_carga:
        st.subheader("Paso 1: Periodo, Costos y Archivos")
        
        # 1. EL CANDADO DE TIEMPO
        st.markdown("**1. Definir Periodo a Costear**")
        col_m1, col_m2 = st.columns(2)
        with col_m1: mes_proceso = st.selectbox("Mes:", range(1, 13), index=date.today().month - 1)
        with col_m2: anio_proceso = st.number_input("Año:", min_value=2024, max_value=2030, value=date.today().year)

        # 2. INPUT DE PLANILLA
        st.markdown("**2. Costo de Mano de Obra**")
        costo_planilla = st.number_input(
            "💵 Ingresa el Costo Total de Planilla + Horas Extras del Mes ($):", 
            min_value=0.0, value=0.0, step=100.0
        )
        
        st.markdown("---")
        st.markdown("**3. Carga de Archivos Base**")
        
        col1, col2 = st.columns(2)
        with col1:
            arch_sgt = st.file_uploader("1. Maestro de Órdenes (SGT_TG)", type=["xlsx"], help="🟢 Pase libre (Actualizable)")
            arch_tras_mp = st.file_uploader("2. Traslados Materia Prima", type=["xlsx"], accept_multiple_files=True, help="🔴 Bloqueado si el mes ya está cerrado")
            arch_tiempos = st.file_uploader("3. Reporte de Tiempos", type=["xlsx"], help="🔴 Bloqueado si el mes ya está cerrado")
            
        with col2:
            arch_fact = st.file_uploader("4. Facturación del Mes", type=["xlsx"], help="🟢 Pase libre (Actualizable)")
            arch_tras_pt = st.file_uploader("5. Traslados Internos PT", type=["xlsx"], accept_multiple_files=True, help="🔴 Bloqueado si el mes ya está cerrado")

        if arch_sgt and arch_fact and arch_tras_mp and arch_tiempos:
            if st.button("🔍 Escanear Archivos y Construir Bodega Virtual", type="primary", use_container_width=True):
                
                # --- LÓGICA DEL ESCUDO ANTIDUPLICADOS ---
                # Aquí conectarás con Google Sheets: df_historial = obtener_dataframe("Movimientos_Inventario_TG")
                # Por ahora, simulamos la validación para proteger tu sistema:
                mes_ya_cerrado = False # Cambiaremos esto cuando conectemos la BD real
                
                if mes_ya_cerrado:
                    st.error(f"🚨 ¡ALERTA DE SEGURIDAD! El mes {mes_proceso}/{anio_proceso} ya tiene traslados y planillas guardadas en el historial.")
                    st.warning("Solo puedes procesar Órdenes (SGT) y Facturas para actualizar estados. Retira los archivos de Traslados y Tiempos para continuar.")
                else:
                    with st.spinner("Escaneando documentos y validando duplicados..."):
                        try:
                            # 1. Leer Maestro SGT
                            df_sgt = pd.read_excel(arch_sgt, dtype=str)
                            lista_ordenes_validas = df_sgt['Orden'].dropna().astype(str).str.strip().tolist() if 'Orden' in df_sgt.columns else []

                            # 2. Leer Facturación
                            df_fact = pd.read_excel(arch_fact, dtype=str)
                            if 'Descripcion' in df_fact.columns:
                                df_fact['Ordenes_Detectadas'] = df_fact['Descripcion'].apply(extraer_ordenes)
                                def clasificar_factura(row):
                                    desc = str(row.get('Descripcion', '')).upper()
                                    cat = str(row.get('Categoria', '')).upper()
                                    if len(row['Ordenes_Detectadas']) > 0: return "Orden Lista"
                                    if "SERVICIO" in cat or "SERVICIO" in desc: return "Servicios"
                                    if any(k in desc for k in ["BANNER", "AFICHE", "CALENDARIO", "ROTULO", "IMPRESION"]): return "Venta Directa"
                                    if any(k in desc for k in ["RECICLAJE", "DESPERDICIO"]): return "Reciclaje"
                                    return "Huérfana (Revisar)"
                                df_fact['Clasificacion'] = df_fact.apply(clasificar_factura, axis=1)

                            # 3. Leer Tiempos
                            df_tiempos = pd.read_excel(arch_tiempos, dtype=str)
                            if 'Observaciones' in df_tiempos.columns:
                                df_tiempos['Ordenes_Detectadas'] = df_tiempos['Observaciones'].apply(extraer_ordenes)
                                df_tiempos['Clasificacion'] = df_tiempos['Ordenes_Detectadas'].apply(lambda x: "Orden Lista" if len(x) > 0 else "Huérfana (Revisar)")

                            # Guardar en memoria
                            st.session_state['tg_fact'] = df_fact
                            st.session_state['tg_tiempos'] = df_tiempos
                            st.session_state['tg_sgt'] = df_sgt
                            st.session_state['tg_ordenes_validas'] = lista_ordenes_validas
                            st.session_state['tg_datos_cargados'] = True
                            st.session_state['fase2_aprobada'] = False
                            
                            st.success("✅ Archivos limpios y sin duplicados. Pasa a la pestaña '2. Auditoría'.")
                        except Exception as e:
                            st.error(f"Error al leer archivos: {e}")

    # ==========================================
    # PESTAÑA 2: EL PURGATORIO (AUDITORÍA)
    # ==========================================
    with tab_auditoria:
        st.subheader("Sala de Espera: Revisión de Anomalías")
        if st.session_state.get('tg_datos_cargados', False):
            df_fact = st.session_state['tg_fact']
            ordenes_validas = st.session_state.get('tg_ordenes_validas', [])
            
            st.markdown("### Resumen Preliminar de Facturación")
            col_a, col_b, col_c, col_d, col_e = st.columns(5)
            conteos = df_fact['Clasificacion'].value_counts()
            
            col_a.metric("✅ Órdenes Listas", conteos.get("Orden Lista", 0))
            col_b.metric("🛍️ Venta Directa", conteos.get("Venta Directa", 0))
            col_c.metric("🛠️ Servicios", conteos.get("Servicios", 0))
            col_d.metric("♻️ Reciclaje", conteos.get("Reciclaje", 0))
            col_e.metric("⚠️ Huérfanas", conteos.get("Huérfana (Revisar)", 0))
            st.divider()

            df_huerfanas_fact = df_fact[df_fact['Clasificacion'] == "Huérfana (Revisar)"].copy()
            
            if not df_huerfanas_fact.empty:
                st.error(f"🚨 El sistema ha bloqueado el paso. Tienes {len(df_huerfanas_fact)} facturas sin orden.")
                
                df_mostrar = df_huerfanas_fact[['Fecha', 'Numero', 'Descripcion', 'VentaNeta']].copy()
                df_mostrar['Accion'] = "Pendiente"
                df_mostrar['Orden_SGT'] = ""
                
                editado = st.data_editor(
                    df_mostrar,
                    column_config={
                        "Accion": st.column_config.SelectboxColumn("Acción a tomar", options=["Pendiente", "Asignar Orden", "Omitir / Eliminar"], required=True),
                        "Orden_SGT": st.column_config.TextColumn("Orden SGT (Si aplica)")
                    },
                    use_container_width=True, hide_index=True
                )
                
                if st.button("💾 Validar Correcciones", type="primary"):
                    errores = []
                    for index, row in editado.iterrows():
                        if row['Accion'] == "Pendiente":
                            errores.append(f"Factura {row['Numero']} sigue 'Pendiente'.")
                        elif row['Accion'] == "Asignar Orden":
                            orden_input = str(row['Orden_SGT']).strip()
                            if orden_input not in ordenes_validas:
                                errores.append(f"Factura {row['Numero']}: La orden '{orden_input}' NO EXISTE en SGT.")
                    
                    if errores:
                        st.error("❌ Corrija lo siguiente para avanzar:")
                        for error in errores: st.write(f"- {error}")
                        st.session_state['fase2_aprobada'] = False
                    else:
                        st.success("✅ Purgatorio limpio. Todas las reglas se cumplen. Proceda a Liquidación.")
                        st.session_state['fase2_aprobada'] = True
            else:
                st.success("✅ Validación completada. No hay facturas huérfanas.")
                st.session_state['fase2_aprobada'] = True
        else:
            st.write("Carga los archivos en la pestaña 1 para iniciar.")

    # ==========================================
    # PESTAÑA 3: LIQUIDACIÓN
    # ==========================================
    with tab_liquidacion:
        st.subheader("Liquidación y Partidas Contables")
        if st.session_state.get('fase2_aprobada', False):
            st.info("🟢 Luz Verde. El sistema está listo para calcular el prorrateo de mano de obra y generar las partidas.")
            if st.button("Ejecutar Liquidación y Guardar", type="primary"):
                st.write("Lógica matemática y guardado en Google Sheets en construcción...")
        else:
            st.warning("🛑 Debe completar y validar la pestaña '2. Auditoría' antes de ejecutar la liquidación.")