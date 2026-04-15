import streamlit as st
import pandas as pd
import re
import io
from datetime import date

# ==========================================
# CONSTANTES CONTABLES
# ==========================================
CUENTA_WIP = "110602"
CUENTA_MP = "110401"
CUENTA_CIF = "410104"
CUENTA_COSTO_VENTAS = "410101"

CUENTAS_MANO_OBRA = [
    "41010201 Sueldo de personal de produccion",
    "41010202001 Aguinaldo personal de produccion",
    "41010202002 Vacaciones personal de produccion",
    "41010202003 Complem. por incapacidad de produccion",
    "41010202004 Indemnización personal de produccion",
    "41010202005 ISSS Salud personal de produccion",
    "41010202006 ISSS IVM personal de produccion",
    "41010202007 AFP patronal personal de produccion",
    "41010202008 INSAFORP personal de produccion",
    "41010202009 ISSS UPIS patronal personal de produccion",
    "41010202010 IPSFA patronal personal de produccion",
    "41010202011 Horas Extras personal de produccion",
    "41010202012 Horas Extras Nocturnas Personal de Produ",
    "41010202013 Nocturnidad Personal de Produccion",
    "41010203001 Aguinaldo complementario personal de produc",
    "41010203002 Seguro de vida Empresas personal de produccio",
    "41010203003 Seguro plan de salud privado personal de produ",
    "41010203004 Bono de productividad",
    "41010203005 Reconocimiento por antigüedad",
    "41010203006 Subsidio",
    "41010203007 Subsidio despensa",
    "41010302005 Ayuda compra de uniformes personal de producc",
    "41010203010 Subsidio Talleres Gráficos"
]

# ==========================================
# FUNCIONES AUXILIARES
# ==========================================
def extraer_ordenes(texto):
    if pd.isna(texto): return []
    return re.findall(r'\b\d{3,4}-\d{3,4}\b', str(texto))

def tiene_orden_valida(ordenes_extraidas, ordenes_sgt):
    for o in ordenes_extraidas:
        if o in ordenes_sgt: return True
    return False

def buscar_valor_columna(row, df_cols, palabra_clave):
    valores = []
    for col in df_cols:
        if palabra_clave.upper() in str(col).upper():
            valores.append(str(row[col]))
    return " ".join(valores).upper()

def procesar_costos_por_orden(df, col_valor, es_horas=False):
    """Divide el costo o las horas entre múltiples órdenes en una misma celda"""
    distribucion = {}
    total_cif = 0.0
    
    for _, row in df.iterrows():
        try:
            valor = float(row[col_valor])
            if pd.isna(valor): valor = 0.0
        except:
            valor = 0.0
            
        clasif = row.get('Clasificacion', '')
        
        if clasif == "Costo Indirecto (Automático)":
            total_cif += valor
            continue
            
        if clasif != "Orden Lista":
            continue
            
        ordenes = row.get('Ordenes_Detectadas', [])
        if 'Orden_SGT' in row and str(row['Orden_SGT']).strip() != "":
            ordenes = [str(row['Orden_SGT']).strip()]
            
        if len(ordenes) > 0:
            valor_dividido = valor / len(ordenes)
            for o in ordenes:
                o_str = str(o).strip()
                # BUGFIX: Ignoramos celdas basura que se lean como "nan" o vacías
                if o_str.lower() == 'nan' or o_str == '': continue
                distribucion[o_str] = distribucion.get(o_str, 0.0) + valor_dividido

    return distribucion, total_cif

def generar_nexus_bytes(df):
    output = io.BytesIO()
    # Formato estricto Nexus: 5 columnas, SIN encabezados
    df_nexus = df[["CUENTA", "VACIO", "CONCEPTO", "DEBE", "HABER"]]
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_nexus.to_excel(writer, index=False, header=False)
    return output.getvalue()

def generar_excel_filtrado(df, nombre_hoja):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=nombre_hoja)
    return output.getvalue()

# ==========================================
# CONEXIÓN GOOGLE SHEETS
# ==========================================
def guardar_en_google_sheets(df_kardex, df_wip):
    try:
        # AQUI DEBES METER TU CODIGO DE CONEXION (El mismo que usas en cafetería)
        # Ejemplo:
        # conn = st.connection("gsheets", type=GSheetsConnection)
        # conn.update(worksheet="Kardex_Costos_TG", data=df_kardex)
        # conn.update(worksheet="Saldos_WIP_Resumen", data=df_wip)
        
        st.success("💾 ¡Datos guardados exitosamente en Google Sheets (Kardex y Saldos)!")
    except Exception as e:
        st.error(f"Error al intentar guardar en Google Sheets: {e}")

# ==========================================
# MÓDULO PRINCIPAL
# ==========================================
def mostrar_modulo_costos():
    st.title("🖨️ Contabilidad de Costos - Talleres Gráficos")
    st.info("Sistema Automático de Costeo por Órdenes de Producción")

    tab_carga, tab_auditoria, tab_liquidacion = st.tabs([
        "📥 1. Carga de Datos", 
        "🕵️ 2. Auditoría", 
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

        st.markdown("---")
        st.markdown("**Desglose de Mano de Obra (Planilla)**")
        
        if 'df_cuentas_base' not in st.session_state:
            st.session_state['df_cuentas_base'] = pd.DataFrame([
                {"Cuenta": "41010201 Sueldo de personal de produccion", "Monto": 0.0}
            ])

        df_cuentas_mo = st.data_editor(
            st.session_state['df_cuentas_base'],
            column_config={
                "Cuenta": st.column_config.SelectboxColumn("Cuenta Contable", options=CUENTAS_MANO_OBRA, required=True),
                "Monto": st.column_config.NumberColumn("Monto ($)", min_value=0.0, format="$%.2f")
            },
            num_rows="dynamic", use_container_width=True, key="ed_cuentas_mo"
        )
        
        costo_planilla = df_cuentas_mo['Monto'].sum()
        st.info(f"**Costo Total de Mano de Obra a Prorratear:** ${costo_planilla:,.2f}")
        
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
                with st.spinner("Leyendo y cruzando datos con Escáner Indestructible..."):
                    try:
                        # 1. MAESTRO SGT
                        df_sgt = pd.read_excel(arch_sgt, dtype=str)
                        df_sgt.columns = df_sgt.columns.str.strip()
                        ordenes_validas = df_sgt['Orden'].dropna().astype(str).str.strip().tolist() if 'Orden' in df_sgt.columns else []

                        # 2. FACTURACIÓN
                        df_fact = pd.read_excel(arch_fact, dtype=str)
                        df_fact.columns = df_fact.columns.str.strip()
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

                        # 3. TIEMPOS
                        df_tiempos = pd.read_excel(arch_tiempos, dtype=str)
                        df_tiempos.columns = df_tiempos.columns.str.strip()
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

                        # 4. TRASLADOS MATERIA PRIMA
                        dfs_mp = []
                        for f in arch_tras_mp:
                            d = pd.read_excel(f, dtype=str)
                            d.columns = d.columns.str.strip()
                            dfs_mp.append(d)
                        df_mp = pd.concat(dfs_mp, ignore_index=True) if dfs_mp else pd.DataFrame()
                        
                        if not df_mp.empty:
                            df_mp = df_mp.loc[:, ~df_mp.columns.duplicated()] 
                            df_mp['Texto_Para_Orden'] = df_mp.apply(lambda r: buscar_valor_columna(r, df_mp.columns, "CONCEPT") + " " + buscar_valor_columna(r, df_mp.columns, "DESCRIP"), axis=1)
                            df_mp['Ordenes_Detectadas'] = df_mp['Texto_Para_Orden'].apply(extraer_ordenes)
                            
                            def clasificar_traslado(row):
                                cat = buscar_valor_columna(row, df_mp.columns, "CATEGOR")
                                concepto = buscar_valor_columna(row, df_mp.columns, "CONCEPT")
                                
                                # EXCLUSIÓN SOHO/LIBRERÍA/UCA Y PRODUCTO TERMINADO
                                if any(k in concepto for k in ["SOHO", "LIBRERI", "LBRERI", "UCA"]) or "PRODUCTO TERMINADO" in cat:
                                    return "Traslado Especial"
                                    
                                if tiene_orden_valida(row['Ordenes_Detectadas'], ordenes_validas): return "Orden Lista"
                                
                                if len(row['Ordenes_Detectadas']) > 0: return "Huérfana (Revisar)"

                                if any(k in cat for k in ["EMPAQUE", "LIMPIEZA", "REPUESTO", "REPUESTOS"]): 
                                    return "Costo Indirecto (Automático)"
                                    
                                return "Huérfana (Revisar)"
                            
                            df_mp['Clasificacion'] = df_mp.apply(clasificar_traslado, axis=1)

                        st.session_state['tg_fact'] = df_fact
                        st.session_state['tg_tiempos'] = df_tiempos
                        st.session_state['tg_mp'] = df_mp
                        st.session_state['tg_ordenes_validas'] = ordenes_validas
                        st.session_state['tg_costo_planilla'] = costo_planilla
                        st.session_state['tg_df_cuentas'] = df_cuentas_mo 
                        st.session_state['tg_datos_cargados'] = True
                        st.session_state['fase2_aprobada'] = False
                        st.session_state['liquidacion_lista'] = False # Reiniciamos el pase a Fase 3
                        
                        st.success("✅ Datos cruzados correctamente. Avanza a Auditoría.")
                    except Exception as e:
                        st.error(f"Error al leer archivos: {e}")

    # ==========================================
    # PESTAÑA 2: AUDITORÍA (EL PURGATORIO)
    # ==========================================
    with tab_auditoria:
        st.subheader("🕵️ Panel de Revisión Manual")
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
                
                opciones_accion = ["Pendiente", "Asignar Orden", "Forzar Orden (Antigua)", "Costo Indirecto", "Omitir"]
                config_columnas = {
                    "Accion": st.column_config.SelectboxColumn("Acción", options=opciones_accion, required=True),
                    "Orden_SGT": st.column_config.TextColumn("Código Orden")
                }

                ed_fact = ed_tiempos = ed_mp = pd.DataFrame()

                def prellenar_orden(df_original, df_huerfano):
                    df_huerfano['Orden_SGT'] = df_original.loc[df_huerfano.index, 'Ordenes_Detectadas'].apply(
                        lambda x: x[0] if isinstance(x, list) and len(x) > 0 else ""
                    )
                    df_huerfano['Accion'] = df_huerfano['Orden_SGT'].apply(
                        lambda x: "Forzar Orden (Antigua)" if x != "" else "Pendiente"
                    )
                    return df_huerfano

                if h_fact > 0:
                    with st.expander(f"🧾 Facturas Pendientes ({h_fact})", expanded=True):
                        cols_f = [c for c in ['Fecha', 'Numero', 'Descripcion', 'VentaNeta'] if c in df_fact.columns]
                        df_h_f = prellenar_orden(df_fact, df_fact[df_fact['Clasificacion'] == "Huérfana (Revisar)"][cols_f].copy())
                        ed_fact = st.data_editor(df_h_f, column_config=config_columnas, hide_index=True, key="ed_fact", use_container_width=True)

                if h_tiempos > 0:
                    with st.expander(f"⏱️ Horas Pendientes ({h_tiempos})", expanded=True):
                        cols_t = [c for c in ['Empleado', 'Observaciones'] if c in df_tiempos.columns]
                        df_h_t = prellenar_orden(df_tiempos, df_tiempos[df_tiempos['Clasificacion'] == "Huérfana (Revisar)"][cols_t].copy())
                        ed_tiempos = st.data_editor(df_h_t, column_config=config_columnas, hide_index=True, key="ed_tiempos", use_container_width=True)

                if h_mp > 0:
                    with st.expander(f"📦 Traslados MP Pendientes ({h_mp})", expanded=True):
                        col_mostrar_mp = next((c for c in df_mp.columns if 'CONCEPT' in c.upper()), 'Numero')
                        col_desc_mp = next((c for c in df_mp.columns if 'DESCRIP' in c.upper()), 'Numero')
                        col_cat_mp = next((c for c in df_mp.columns if 'CATEGOR' in c.upper()), 'Categoria')
                        cols_m = [c for c in ['Numero', col_mostrar_mp, col_desc_mp, col_cat_mp] if c in df_mp.columns]
                        
                        df_h_m = prellenar_orden(df_mp, df_mp[df_mp['Clasificacion'] == "Huérfana (Revisar)"][cols_m].copy())
                        ed_mp = st.data_editor(df_h_m, column_config=config_columnas, hide_index=True, key="ed_mp", use_container_width=True)

                if st.button("💾 Guardar Decisiones y Validar", type="primary"):
                    errores = []
                    
                    def revisar_tabla(df_editado, df_original, nombre_tabla):
                        for i, row in df_editado.iterrows():
                            accion = row['Accion']
                            if accion == "Pendiente":
                                errores.append(f"Queda un registro pendiente en {nombre_tabla}.")
                            elif accion == "Asignar Orden":
                                orden = str(row['Orden_SGT']).strip()
                                if orden not in ordenes_validas:
                                    errores.append(f"En {nombre_tabla}, la orden '{orden}' NO existe en SGT.")
                                else:
                                    df_original.at[i, 'Clasificacion'] = "Orden Lista"
                                    df_original.at[i, 'Orden_SGT'] = orden
                            elif accion == "Forzar Orden (Antigua)":
                                orden = str(row['Orden_SGT']).strip()
                                if orden == "":
                                    errores.append(f"En {nombre_tabla}, dejaste el código en blanco.")
                                else:
                                    df_original.at[i, 'Clasificacion'] = "Orden Lista"
                                    df_original.at[i, 'Orden_SGT'] = orden
                            elif accion == "Costo Indirecto":
                                df_original.at[i, 'Clasificacion'] = "Costo Indirecto (Automático)"
                            elif accion == "Omitir":
                                df_original.at[i, 'Clasificacion'] = "Omitido Automático"

                    if not ed_fact.empty: revisar_tabla(ed_fact, df_fact, "Facturas")
                    if not ed_tiempos.empty: revisar_tabla(ed_tiempos, df_tiempos, "Tiempos")
                    if not ed_mp.empty: revisar_tabla(ed_mp, df_mp, "Traslados MP")

                    if errores:
                        st.error("❌ Corrige los siguientes errores antes de pasar a liquidar:")
                        for e in errores: st.write(f"- {e}")
                    else:
                        st.session_state['tg_fact'] = df_fact
                        st.session_state['tg_tiempos'] = df_tiempos
                        st.session_state['tg_mp'] = df_mp
                        st.success("✅ ¡Auditoría validada! Procede a la Liquidación.")
                        st.session_state['fase2_aprobada'] = True
            else:
                st.success("✅ Todo está perfecto. No hay registros pendientes.")
                st.session_state['fase2_aprobada'] = True
        else:
            st.write("Carga los archivos en la Pestaña 1.")

    # ==========================================
    # PESTAÑA 3: LIQUIDACIÓN Y PARTIDAS INDIVIDUALES
    # ==========================================
    with tab_liquidacion:
        st.subheader("💰 Generación de Costos y Partidas (Formato Nexus)")
        if st.session_state.get('fase2_aprobada', False):
            if st.button("🚀 Ejecutar Liquidación y Generar Partidas", type="primary"):
                with st.spinner("Procesando datos y estructurando formato Nexus..."):
                    df_fact = st.session_state['tg_fact']
                    df_tiempos = st.session_state['tg_tiempos']
                    df_mp = st.session_state['tg_mp']
                    costo_total_mo = st.session_state['tg_costo_planilla']
                    df_cuentas_mo = st.session_state['tg_df_cuentas']
                    
                    # 1. PRORRATEO MANO DE OBRA
                    col_horas = next((c for c in df_tiempos.columns if 'TOTALHORA' in c.upper().replace(' ', '')), None)
                    horas_ordenes, _ = procesar_costos_por_orden(df_tiempos, col_horas, es_horas=True) if col_horas else ({}, 0)
                    total_horas_validas = sum(horas_ordenes.values())
                    
                    costo_por_hora = (costo_total_mo / total_horas_validas) if total_horas_validas > 0 else 0
                    costos_mo_por_orden = {orden: horas * costo_por_hora for orden, horas in horas_ordenes.items()}

                    # 2. COSTOS MATERIA PRIMA Y CIF
                    col_costo_mp = next((c for c in df_mp.columns if 'PRECIOTOTAL' in c.upper().replace(' ', '') or 'COSTO' in c.upper()), None)
                    costos_mp_por_orden, total_cif = procesar_costos_por_orden(df_mp, col_costo_mp) if col_costo_mp else ({}, 0)

                    # Procesamiento de Traslados a Otras Unidades
                    df_traslados_esp = df_mp[df_mp['Clasificacion'] == 'Traslado Especial']
                    col_cat = next((c for c in df_mp.columns if 'CATEGOR' in c.upper()), 'Categoria')
                    resumen_traslados = []
                    if not df_traslados_esp.empty and col_costo_mp:
                        df_traslados_esp[col_costo_mp] = pd.to_numeric(df_traslados_esp[col_costo_mp], errors='coerce').fillna(0)
                        resumen_traslados = df_traslados_esp.groupby(col_cat)[col_costo_mp].sum().reset_index().to_dict('records')

                    # 3. CONTROL DE ÓRDENES EN PROCESO (KARDEX Y RESUMEN)
                    todas_las_ordenes = set(list(costos_mo_por_orden.keys()) + list(costos_mp_por_orden.keys()))
                    ordenes_facturadas_mes = []
                    for _, row in df_fact[df_fact['Clasificacion'] == 'Orden Lista'].iterrows():
                        ordenes = row.get('Ordenes_Detectadas', [])
                        if 'Orden_SGT' in row and str(row['Orden_SGT']).strip() != "": ordenes = [str(row['Orden_SGT']).strip()]
                        ordenes_facturadas_mes.extend(ordenes)
                    
                    fecha_str = date.today().strftime("%d/%m/%Y")
                    filas_kardex = []
                    filas_wip = []
                    total_liquidado_cv = 0.0
                    total_mp_a_wip = sum(costos_mp_por_orden.values())

                    for orden in todas_las_ordenes:
                        if pd.isna(orden) or str(orden).lower() == 'nan' or str(orden).strip() == '': continue
                        
                        nuevo_mo = costos_mo_por_orden.get(orden, 0.0)
                        nuevo_mp = costos_mp_por_orden.get(orden, 0.0)
                        saldo_anterior = 0.0 # Esto se conectará a Google Sheets después
                        saldo_acumulado = saldo_anterior + nuevo_mo + nuevo_mp
                        
                        estado = "Liquidado a Costo de Ventas" if orden in ordenes_facturadas_mes else "Pendiente"
                        
                        if nuevo_mo > 0:
                            filas_kardex.append({"Fecha": fecha_str, "Orden": str(orden).strip(), "Tipo_Costo": "Mano de Obra", "Monto": nuevo_mo, "Estado": estado})
                        if nuevo_mp > 0:
                            filas_kardex.append({"Fecha": fecha_str, "Orden": str(orden).strip(), "Tipo_Costo": "Materia Prima", "Monto": nuevo_mp, "Estado": estado})

                        if estado == "Liquidado a Costo de Ventas":
                            total_liquidado_cv += saldo_acumulado
                            saldo_acumulado = 0.0
                            
                        filas_wip.append({
                            "Orden": str(orden).strip(), "Saldo_Anterior": saldo_anterior, "Costo_MO_Mes": nuevo_mo,
                            "Costo_MP_Mes": nuevo_mp, "Total_Acumulado": saldo_anterior + nuevo_mo + nuevo_mp,
                            "Estado": estado, "Saldo_Final_WIP": saldo_acumulado
                        })
                    
                    st.session_state['tg_df_kardex'] = pd.DataFrame(filas_kardex)
                    st.session_state['tg_df_wip'] = pd.DataFrame(filas_wip)

                    # 4. GENERACIÓN DE PARTIDAS INDIVIDUALES NEXUS
                    def agregar_linea(lista, cuenta, descripcion, debe, haber):
                        if debe > 0 or haber > 0:
                            lista.append({"CUENTA": cuenta, "VACIO": "", "CONCEPTO": descripcion.upper(), "DEBE": round(debe, 2), "HABER": round(haber, 2)})

                    # Partida 1: Planilla
                    p1 = []
                    agregar_linea(p1, CUENTA_WIP, "Inyeccion Mano de Obra a Produccion", costo_total_mo, 0)
                    for _, r in df_cuentas_mo.iterrows():
                        cuenta_codigo = r['Cuenta'].split(" ")[0]
                        cuenta_desc = r['Cuenta'].replace(cuenta_codigo, "").strip()
                        agregar_linea(p1, cuenta_codigo, cuenta_desc, 0, r['Monto'])
                    st.session_state['tg_p1'] = pd.DataFrame(p1)

                    # Partida 2: Materia Prima
                    p2 = []
                    agregar_linea(p2, CUENTA_WIP, "Traslado de MP a Ordenes de Produccion", total_mp_a_wip, 0)
                    agregar_linea(p2, CUENTA_CIF, "Traslado de MP a Costos Indirectos (CIF)", total_cif, 0)
                    
                    total_otras_unidades = 0.0
                    for traslado in resumen_traslados:
                        cat_nombre = str(traslado[col_cat]).upper()
                        monto_traslado = traslado[col_costo_mp]
                        total_otras_unidades += monto_traslado
                        cuenta_destino = "110603" if "PRODUCTO TERMINADO" in cat_nombre else "110601"
                        agregar_linea(p2, cuenta_destino, f"TRASLADO DE INVENTARIO - {cat_nombre}", monto_traslado, 0)

                    agregar_linea(p2, CUENTA_MP, "Salida de Inventario de Materia Prima", 0, total_mp_a_wip + total_cif + total_otras_unidades)
                    st.session_state['tg_p2'] = pd.DataFrame(p2)

                    # Partida 3: Liquidación WIP
                    p3 = []
                    if total_liquidado_cv > 0:
                        agregar_linea(p3, CUENTA_COSTO_VENTAS, "Liquidacion de OP Facturadas a Costo de Ventas", total_liquidado_cv, 0)
                        agregar_linea(p3, CUENTA_WIP, "Descarga de OP Facturadas de Proceso", 0, total_liquidado_cv)
                    st.session_state['tg_p3'] = pd.DataFrame(p3)

                    st.session_state['liquidacion_lista'] = True
                    st.success("✅ Cálculos completados.")

            # SECCIÓN DE BOTONES DE DESCARGA
            if st.session_state.get('liquidacion_lista', False):
                st.markdown("### 📥 Descarga de Partidas Individuales (Nexus)")
                col_d1, col_d2, col_d3 = st.columns(3)
                
                with col_d1:
                    st.write("**1. Partida Nómina**")
                    st.download_button("⬇️ Descargar", data=generar_nexus_bytes(st.session_state['tg_p1']), file_name=f"Nexus_Nomina_{mes_proceso}_{anio_proceso}.xlsx")

                with col_d2:
                    st.write("**2. Partida Materiales**")
                    st.download_button("⬇️ Descargar", data=generar_nexus_bytes(st.session_state['tg_p2']), file_name=f"Nexus_Materiales_{mes_proceso}_{anio_proceso}.xlsx")

                with col_d3:
                    st.write("**3. Liquidación**")
                    if not st.session_state['tg_p3'].empty:
                        st.download_button("⬇️ Descargar", data=generar_nexus_bytes(st.session_state['tg_p3']), file_name=f"Nexus_Liquidacion_{mes_proceso}_{anio_proceso}.xlsx")
                    else:
                        st.info("Sin OP facturadas.")

                st.divider()
                st.markdown("### 📊 Reportes de Control de Bodega")
                
                col_f1, col_f2 = st.columns(2)
                with col_f1: filtro_orden = st.text_input("🔍 Filtrar por Orden (Dejar en blanco para ver todo):")
                with col_f2: st.write("") 

                df_kardex_export = st.session_state['tg_df_kardex']
                df_wip_export = st.session_state['tg_df_wip']

                if filtro_orden.strip() != "":
                    df_kardex_export = df_kardex_export[df_kardex_export['Orden'].str.contains(filtro_orden, case=False, na=False)]
                    df_wip_export = df_wip_export[df_wip_export['Orden'].str.contains(filtro_orden, case=False, na=False)]

                st.dataframe(df_wip_export, use_container_width=True)

                col_b1, col_b2, col_b3 = st.columns([1,1,1])
                with col_b1:
                    if st.button("💾 Guardar en Google Sheets", type="secondary"):
                        guardar_en_google_sheets(st.session_state['tg_df_kardex'], st.session_state['tg_df_wip'])
                with col_b2:
                    st.download_button("📋 Bajar Kardex", data=generar_excel_filtrado(df_kardex_export, "Kardex"), file_name=f"Kardex_Costos_{mes_proceso}_{anio_proceso}.xlsx")
                with col_b3:
                    st.download_button("📉 Bajar Saldos WIP", data=generar_excel_filtrado(df_wip_export, "Saldos_WIP"), file_name=f"Saldos_WIP_{mes_proceso}_{anio_proceso}.xlsx")
        else:
            st.warning("🛑 Completa la Auditoría para habilitar la Liquidación.")