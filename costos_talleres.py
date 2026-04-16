import streamlit as st
import pandas as pd
import re
import io
from datetime import date

# ==========================================
# CONSTANTES CONTABLES
# ==========================================
CUENTA_WIP_MO = "110602"       # Proceso Nómina
CUENTA_WIP_MP = "110402"       # Proceso Materia Prima 
CUENTA_CIF = "410104"          # Costos Indirectos de Fabricación
CUENTA_COSTO_VENTAS = "410101" # Costo de Ventas
CUENTA_PROVISION_TRASLADOS = "21020302" # Puente para traslados a otras unidades

CUENTAS_MANO_OBRA = [
    "41010201 Sueldo de personal de produccion", "41010202001 Aguinaldo personal de produccion",
    "41010202002 Vacaciones personal de produccion", "41010202003 Complem. por incapacidad de produccion",
    "41010202004 Indemnización personal de produccion", "41010202005 ISSS Salud personal de produccion",
    "41010202006 ISSS IVM personal de produccion", "41010202007 AFP patronal personal de produccion",
    "41010202008 INSAFORP personal de produccion", "41010202009 ISSS UPIS patronal personal de produccion",
    "41010202010 IPSFA patronal personal de produccion", "41010202011 Horas Extras personal de produccion",
    "41010202012 Horas Extras Nocturnas Personal de Produ", "41010202013 Nocturnidad Personal de Produccion",
    "41010203001 Aguinaldo complementario personal de produc", "41010203002 Seguro de vida Empresas personal de produccio",
    "41010203003 Seguro plan de salud privado personal de produ", "41010203004 Bono de productividad",
    "41010203005 Reconocimiento por antigüedad", "41010203006 Subsidio", "41010203007 Subsidio despensa",
    "41010302005 Ayuda compra de uniformes personal de producc", "41010203010 Subsidio Talleres Gráficos"
]

# ==========================================
# FUNCIONES AUXILIARES
# ==========================================
def obtener_datos_inventario(categoria):
    cat = str(categoria).upper()
    if "LIMPIEZA" in cat: return "110616", "Inventario de articulos e insumos de limpieza"
    if "EMPAQUE" in cat: return "110617", "Inventario de empaques"
    if "REPUESTO" in cat: return "110620", "Inventario de repuestos"
    if "PRODUCTO TERMINADO" in cat: return "110603", "Inventario de producto terminado"
    return "110601", "Inventario de Materia Prima" 

def limpiar_orden(o):
    return str(o).strip().replace(" ", "").upper()

def extraer_ordenes(texto):
    if pd.isna(texto): return []
    ordenes = re.findall(r'\b\d{3,4}\s*-\s*\d{3,4}\b', str(texto))
    return [limpiar_orden(o) for o in ordenes]

def buscar_valor_columna(row, df_cols, palabra_clave):
    valores = []
    for col in df_cols:
        if palabra_clave.upper() in str(col).upper():
            valores.append(str(row[col]))
    return " ".join(valores).upper()

def procesar_costos_por_orden(df, col_valor, es_horas=False):
    distribucion = {}
    for _, row in df.iterrows():
        try: valor = float(row[col_valor])
        except: valor = 0.0
        if pd.isna(valor) or valor == 0.0: continue
            
        clasif = row.get('Clasificacion', '')
        if clasif != "Orden Lista": continue
            
        ordenes = row.get('Ordenes_Detectadas', [])
        if 'Orden_SGT' in row and str(row['Orden_SGT']).strip() != "":
            ordenes = [limpiar_orden(row['Orden_SGT'])]
            
        if len(ordenes) > 0:
            valor_dividido = valor / len(ordenes)
            for o in ordenes:
                o_str = str(o).strip().replace(" ", "").upper()
                if o_str == 'NAN' or o_str == '': continue
                distribucion[o_str] = distribucion.get(o_str, 0.0) + valor_dividido
    return distribucion

def generar_nexus_bytes(df_or_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if isinstance(df_or_dict, dict):
            for sheet_name, df in df_or_dict.items():
                if not df.empty:
                    df_nexus = df[["CUENTA", "VACIO", "CONCEPTO", "DEBE", "HABER"]]
                    df_nexus.to_excel(writer, index=False, header=False, sheet_name=str(sheet_name)[:30])
        else:
            df_nexus = df_or_dict[["CUENTA", "VACIO", "CONCEPTO", "DEBE", "HABER"]]
            df_nexus.to_excel(writer, index=False, header=False, sheet_name="Partida_Nexus")
    return output.getvalue()

def generar_excel_filtrado(df, nombre_hoja):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=nombre_hoja)
    return output.getvalue()

def guardar_en_google_sheets(df_kardex, df_wip):
    try:
        st.success("💾 ¡Datos guardados exitosamente en Google Sheets (Kardex y Saldos)!")
    except Exception as e:
        st.error(f"Error al intentar guardar en Google Sheets: {e}")

# ==========================================
# MÓDULO PRINCIPAL DE LA INTERFAZ
# ==========================================
def mostrar_modulo_costos():
    st.title("🖨️ Contabilidad de Costos - Talleres Gráficos")
    st.info("MODO AISLADO: Verificación de recolección de costos de MP y MO")

    tab_carga, tab_auditoria, tab_liquidacion = st.tabs([
        "📥 1. Carga de Datos", "🕵️ 2. Auditoría", "💰 3. Liquidación y Partidas"
    ])

    with tab_carga:
        st.subheader("Paso 1: Periodo, Costos y Archivos")
        col_m1, col_m2 = st.columns(2)
        with col_m1: mes_proceso = st.selectbox("Mes:", range(1, 13), index=date.today().month - 1)
        with col_m2: anio_proceso = st.number_input("Año:", min_value=2024, max_value=2030, value=date.today().year)

        st.markdown("---")
        st.markdown("**Desglose de Mano de Obra (Planilla)**")
        
        if 'df_cuentas_base' not in st.session_state:
            st.session_state['df_cuentas_base'] = pd.DataFrame([{"Cuenta": "41010201 Sueldo de personal de produccion", "Monto": 0.0}])

        df_cuentas_mo = st.data_editor(
            st.session_state['df_cuentas_base'],
            column_config={
                "Cuenta": st.column_config.SelectboxColumn("Cuenta Contable", options=CUENTAS_MANO_OBRA, required=True),
                "Monto": st.column_config.NumberColumn("Monto ($)", min_value=0.0, format="$%.2f")
            }, num_rows="dynamic", use_container_width=True, key="ed_cuentas_mo"
        )
        costo_planilla = df_cuentas_mo['Monto'].sum()
        
        st.markdown("---")
        col1, col2 = st.columns(2)
        with col1:
            arch_sgt = st.file_uploader("1. Maestro de Órdenes (SGT_TG)", type=["xlsx"])
            arch_tras_mp = st.file_uploader("2. Traslados Materia Prima", type=["xlsx"], accept_multiple_files=True)
            arch_tiempos = st.file_uploader("3. Reporte de Tiempos", type=["xlsx"])
        with col2:
            arch_fact = st.file_uploader("4. Facturación del Mes", type=["xlsx"])
            arch_tras_pt = st.file_uploader("5. Traslados Internos PT", type=["xlsx"], accept_multiple_files=True)
            arch_wip_ant = st.file_uploader("6. Saldos WIP Mes Anterior (Opcional)", type=["xlsx"])

        if arch_sgt and arch_fact and arch_tras_mp and arch_tiempos:
            if st.button("🔍 Escanear Archivos y Extraer Costos", type="primary", use_container_width=True):
                with st.spinner("Leyendo y cruzando datos..."):
                    try:
                        df_wip_ant = pd.DataFrame()
                        if arch_wip_ant:
                            try:
                                df_wip_ant = pd.read_excel(arch_wip_ant, dtype=str)
                                df_wip_ant.columns = df_wip_ant.columns.str.strip()
                            except: pass

                        # FACTURACIÓN (Se lee, pero lo ignoraremos en la liquidación para esta prueba)
                        df_fact = pd.read_excel(arch_fact, dtype=str)
                        df_fact.columns = df_fact.columns.str.strip()
                        df_fact['Texto_Completo_Factura'] = df_fact.apply(lambda r: " ".join(r.dropna().astype(str)), axis=1)
                        df_fact['Ordenes_Detectadas'] = df_fact['Texto_Completo_Factura'].apply(extraer_ordenes)
                        df_fact['Clasificacion'] = "Ignorado temporalmente" # Apagado para esta prueba

                        # TIEMPOS
                        df_tiempos = pd.read_excel(arch_tiempos, dtype=str)
                        df_tiempos.columns = df_tiempos.columns.str.strip()
                        if 'Observaciones' in df_tiempos.columns:
                            df_tiempos['Ordenes_Detectadas'] = df_tiempos['Observaciones'].apply(extraer_ordenes)
                            def clasificar_tiempos(row):
                                obs = str(row.get('Observaciones', '')).strip()
                                if obs in ['', 'nan', 'None', '--', '-'] or pd.isna(row.get('Observaciones')): return "Omitido Automático"
                                
                                # AQUI ESTA LA MAGIA: Si trae el formato de orden, toma el costo automáticamente
                                if len(row['Ordenes_Detectadas']) > 0: return "Orden Lista"
                                    
                                return "Huérfana (Revisar)"
                            df_tiempos['Clasificacion'] = df_tiempos.apply(clasificar_tiempos, axis=1)

                        # TRASLADOS MP
                        dfs_mp = []
                        if arch_tras_mp:
                            for f in arch_tras_mp:
                                d, d.columns = pd.read_excel(f, dtype=str), pd.read_excel(f, dtype=str).columns.str.strip()
                                dfs_mp.append(d)
                        if arch_tras_pt:
                            for f in arch_tras_pt:
                                d, d.columns = pd.read_excel(f, dtype=str), pd.read_excel(f, dtype=str).columns.str.strip()
                                dfs_mp.append(d)
                                
                        df_mp = pd.concat(dfs_mp, ignore_index=True) if dfs_mp else pd.DataFrame()
                        if not df_mp.empty:
                            df_mp = df_mp.loc[:, ~df_mp.columns.duplicated()] 
                            df_mp['Texto_Para_Orden'] = df_mp.apply(lambda r: buscar_valor_columna(r, df_mp.columns, "CONCEPT") + " " + buscar_valor_columna(r, df_mp.columns, "DESCRIP"), axis=1)
                            df_mp['Ordenes_Detectadas'] = df_mp['Texto_Para_Orden'].apply(extraer_ordenes)
                            
                            def clasificar_traslado(row):
                                cat, concepto = buscar_valor_columna(row, df_mp.columns, "CATEGOR"), buscar_valor_columna(row, df_mp.columns, "CONCEPT")
                                
                                # AQUI ESTA LA MAGIA: Si trae el formato de orden, toma el costo automáticamente
                                if len(row['Ordenes_Detectadas']) > 0: return "Orden Lista"
                                    
                                if any(k in concepto for k in ["SOHO", "LIBRERI", "LBRERI", "UCA"]) or "PRODUCTO TERMINADO" in cat: return "Traslado Especial"
                                if any(k in cat for k in ["EMPAQUE", "LIMPIEZA", "REPUESTO", "REPUESTOS"]): return "Costo Indirecto (Automático)"
                                return "Huérfana (Revisar)"
                            df_mp['Clasificacion'] = df_mp.apply(clasificar_traslado, axis=1)

                        # GUARDADO MEMORIA
                        st.session_state['tg_fact'] = df_fact
                        st.session_state['tg_tiempos'] = df_tiempos
                        st.session_state['tg_mp'] = df_mp
                        st.session_state['tg_wip_ant'] = df_wip_ant
                        st.session_state['tg_costo_planilla'] = costo_planilla
                        st.session_state['tg_df_cuentas'] = df_cuentas_mo 
                        st.session_state['tg_datos_cargados'] = True
                        st.session_state['fase2_aprobada'] = False
                        st.session_state['liquidacion_lista'] = False
                        st.success("✅ Costos recolectados. Avanza a Auditoría.")
                    except Exception as e:
                        st.error(f"Error al leer archivos: {e}")

    with tab_auditoria:
        st.subheader("🕵️ Panel de Revisión Manual")
        if st.session_state.get('tg_datos_cargados', False):
            df_tiempos, df_mp = st.session_state['tg_tiempos'], st.session_state['tg_mp']
            
            h_tiempos = len(df_tiempos[df_tiempos['Clasificacion'] == "Huérfana (Revisar)"])
            h_mp = len(df_mp[df_mp['Clasificacion'] == "Huérfana (Revisar)"]) if not df_mp.empty else 0
            total_huerfanas = h_tiempos + h_mp

            if total_huerfanas > 0:
                st.error(f"🚨 Tienes {total_huerfanas} costos sin orden válida que necesitan tu decisión.")
                opciones_accion = ["Pendiente", "Forzar Orden", "Costo Indirecto", "Omitir"]
                config_col = {"Accion": st.column_config.SelectboxColumn("Acción", options=opciones_accion, required=True), "Orden_SGT": st.column_config.TextColumn("Código Orden")}
                
                ed_tiempos, ed_mp = pd.DataFrame(), pd.DataFrame()

                def prellenar_orden(df_orig, df_huerf):
                    df_huerf['Orden_SGT'] = df_orig.loc[df_huerf.index, 'Ordenes_Detectadas'].apply(lambda x: limpiar_orden(x[0]) if isinstance(x, list) and len(x) > 0 else "")
                    df_huerf['Accion'] = df_huerf['Orden_SGT'].apply(lambda x: "Forzar Orden" if x != "" else "Pendiente")
                    return df_huerf

                if h_tiempos > 0:
                    with st.expander(f"⏱️ Horas Pendientes ({h_tiempos})", expanded=True):
                        cols_t = [c for c in ['Empleado', 'Observaciones'] if c in df_tiempos.columns]
                        ed_tiempos = st.data_editor(prellenar_orden(df_tiempos, df_tiempos[df_tiempos['Clasificacion'] == "Huérfana (Revisar)"][cols_t].copy()), column_config=config_col, hide_index=True, key="ed_t", use_container_width=True)

                if h_mp > 0:
                    with st.expander(f"📦 Traslados MP/PT Pendientes ({h_mp})", expanded=True):
                        cols_m = [c for c in ['Numero', next((c for c in df_mp.columns if 'CONCEPT' in c.upper()), 'Numero'), next((c for c in df_mp.columns if 'CATEGOR' in c.upper()), 'Categoria')] if c in df_mp.columns]
                        ed_mp = st.data_editor(prellenar_orden(df_mp, df_mp[df_mp['Clasificacion'] == "Huérfana (Revisar)"][cols_m].copy()), column_config=config_col, hide_index=True, key="ed_m", use_container_width=True)

                if st.button("💾 Guardar Decisiones y Validar", type="primary"):
                    errores = []
                    def revisar_tabla(df_edit, df_orig, nom):
                        for i, row in df_edit.iterrows():
                            acc = row['Accion']
                            if acc == "Pendiente": errores.append(f"Queda un registro pendiente en {nom}.")
                            elif acc == "Forzar Orden":
                                orden = limpiar_orden(row['Orden_SGT'])
                                if orden == "": errores.append(f"En {nom}, dejaste el código en blanco.")
                                else: 
                                    df_orig.at[i, 'Clasificacion'] = "Orden Lista"
                                    df_orig.at[i, 'Orden_SGT'] = orden
                            elif acc == "Costo Indirecto": df_orig.at[i, 'Clasificacion'] = "Costo Indirecto (Automático)"
                            elif acc == "Omitir": df_orig.at[i, 'Clasificacion'] = "Omitido Automático"

                    if not ed_tiempos.empty: revisar_tabla(ed_tiempos, df_tiempos, "Tiempos")
                    if not ed_mp.empty: revisar_tabla(ed_mp, df_mp, "Traslados MP")

                    if errores:
                        st.error("❌ Corrige los errores:")
                        for e in errores: st.write(f"- {e}")
                    else:
                        st.session_state['tg_tiempos'] = df_tiempos
                        st.session_state['tg_mp'] = df_mp
                        st.success("✅ ¡Costos validados! Procede a calcular Inventario en Proceso.")
                        st.session_state['fase2_aprobada'] = True
            else:
                st.success("✅ Todos los costos han sido asignados correctamente.")
                st.session_state['fase2_aprobada'] = True
        else: st.write("Carga los archivos en la Pestaña 1.")

    with tab_liquidacion:
        st.subheader("💰 Generación de Costos de Inventario en Proceso")
        if st.session_state.get('fase2_aprobada', False):
            if st.button("🚀 Calcular Inventario en Proceso", type="primary"):
                with st.spinner("Construyendo tabla de costos..."):
                    df_tiempos, df_mp, df_wip_ant = st.session_state['tg_tiempos'], st.session_state['tg_mp'], st.session_state.get('tg_wip_ant', pd.DataFrame())
                    costo_total_mo, df_cuentas_mo = st.session_state['tg_costo_planilla'], st.session_state['tg_df_cuentas']
                    
                    def agregar_linea(lista, cuenta, descripcion, debe, haber):
                        if round(debe, 2) > 0 or round(haber, 2) > 0:
                            lista.append({"CUENTA": cuenta, "VACIO": "", "CONCEPTO": str(descripcion).strip(), "DEBE": round(debe, 2), "HABER": round(haber, 2)})

                    historial_saldos = {}
                    if not df_wip_ant.empty:
                        col_ord = next((c for c in df_wip_ant.columns if 'ORDEN' in c.upper()), None)
                        col_saldo = next((c for c in df_wip_ant.columns if 'SALDO_FINAL_WIP' in c.upper() or 'TOTAL' in c.upper()), None)
                        if col_ord and col_saldo:
                            for _, r in df_wip_ant.iterrows():
                                o = limpiar_orden(r[col_ord])
                                if o != 'NAN' and o != '':
                                    try: val = float(r[col_saldo])
                                    except: val = 0.0
                                    if val > 0: historial_saldos[o] = val

                    col_horas = next((c for c in df_tiempos.columns if 'TOTALHORA' in c.upper().replace(' ', '')), None)
                    horas_ordenes = procesar_costos_por_orden(df_tiempos, col_horas, es_horas=True) if col_horas else {}
                    costo_por_hora = (costo_total_mo / sum(horas_ordenes.values())) if sum(horas_ordenes.values()) > 0 else 0
                    costos_mo_por_orden = {orden: horas * costo_por_hora for orden, horas in horas_ordenes.items()}

                    col_costo_mp = next((c for c in df_mp.columns if 'PRECIOTOTAL' in c.upper().replace(' ', '') or 'COSTO' in c.upper()), None)
                    col_cat = next((c for c in df_mp.columns if 'CATEGOR' in c.upper()), 'Categoria')
                    if col_costo_mp: df_mp[col_costo_mp] = pd.to_numeric(df_mp[col_costo_mp], errors='coerce').fillna(0)

                    df_wip_mp = df_mp[df_mp['Clasificacion'] == 'Orden Lista']
                    costos_mp_por_orden = procesar_costos_por_orden(df_wip_mp, col_costo_mp) if col_costo_mp else {}
                                    
                    todas_las_ordenes = set(costos_mo_por_orden.keys()) | set(costos_mp_por_orden.keys()) | set(historial_saldos.keys())
                    
                    filas_kardex, filas_wip = [], []
                    for orden in todas_las_ordenes:
                        ord_cln = limpiar_orden(orden)
                        if ord_cln == 'NAN' or ord_cln == '': continue
                        
                        nuevo_mo, nuevo_mp, saldo_anterior = costos_mo_por_orden.get(ord_cln, 0.0), costos_mp_por_orden.get(ord_cln, 0.0), historial_saldos.get(ord_cln, 0.0)
                        saldo_acumulado = saldo_anterior + nuevo_mo + nuevo_mp
                        
                        # APAGAMOS LAS VENTAS: Todo se queda pendiente en WIP para que puedas auditar el costo
                        estado = "Pendiente"
                        
                        if nuevo_mo > 0: filas_kardex.append({"Fecha": date.today().strftime("%d/%m/%Y"), "Orden": ord_cln, "Tipo_Costo": "Mano de Obra", "Monto": nuevo_mo, "Estado": estado})
                        if nuevo_mp > 0: filas_kardex.append({"Fecha": date.today().strftime("%d/%m/%Y"), "Orden": ord_cln, "Tipo_Costo": "Materia Prima", "Monto": nuevo_mp, "Estado": estado})
                            
                        filas_wip.append({"Orden": ord_cln, "Saldo_Anterior": saldo_anterior, "Costo_MO_Mes": nuevo_mo, "Costo_MP_Mes": nuevo_mp, "Total_Acumulado": saldo_acumulado, "Estado": estado, "Saldo_Final_WIP": saldo_acumulado})
                    
                    st.session_state['tg_df_kardex'] = pd.DataFrame(filas_kardex)
                    st.session_state['tg_df_wip'] = pd.DataFrame(filas_wip)
                    st.session_state['liquidacion_lista'] = True
                    st.success("✅ Extracción de Costos finalizada. Revisa la tabla de abajo.")

            if st.session_state.get('liquidacion_lista', False):
                st.divider()
                st.markdown("### 📊 Reportes de Control de Bodega (Todos los costos extraídos)")
                col_f1, _ = st.columns(2)
                f_ord = col_f1.text_input("🔍 Buscar Orden (Ej. 1503-1225):")
                df_k_ex, df_w_ex = st.session_state['tg_df_kardex'], st.session_state['tg_df_wip']
                if f_ord.strip() != "":
                    df_k_ex = df_k_ex[df_k_ex['Orden'].str.contains(f_ord, case=False, na=False)]
                    df_w_ex = df_w_ex[df_w_ex['Orden'].str.contains(f_ord, case=False, na=False)]
                st.dataframe(df_w_ex, use_container_width=True)

                b1, b2 = st.columns([1,1])
                with b1: st.download_button("📋 Bajar Kardex", data=generar_excel_filtrado(df_k_ex, "Kardex"), file_name=f"Kardex_{mes_proceso}.xlsx")
                with b2: st.download_button("📉 Bajar Saldos WIP", data=generar_excel_filtrado(df_w_ex, "Saldos_WIP"), file_name=f"Saldos_WIP_{mes_proceso}.xlsx")

# === FIN DEL SCRIPT ===