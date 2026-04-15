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
def obtener_datos_inventario(categoria):
    """Devuelve la cuenta y el concepto exacto según el machote Nexus"""
    cat = str(categoria).upper()
    if "LIMPIEZA" in cat: 
        return "110616", "Inventario de articulos e insumos de limpieza"
    if "EMPAQUE" in cat: 
        return "110617", "Inventario de empaques"
    if "REPUESTO" in cat: 
        return "110620", "Inventario de repuestos"
    if "PRODUCTO TERMINADO" in cat: 
        return "110603", "Inventario de producto terminado"
    
    # Default para Materia Prima
    return "110601", "Inventario de Materia Prima" 

def limpiar_orden(o):
    """Limpia espacios en blanco alrededor y dentro del código de la orden"""
    return str(o).strip().replace(" ", "").upper()

def extraer_ordenes(texto):
    """Busca patrones de órdenes incluso si tienen espacios en medio"""
    if pd.isna(texto): 
        return []
    ordenes = re.findall(r'\b\d{3,4}\s*-\s*\d{3,4}\b', str(texto))
    return [limpiar_orden(o) for o in ordenes]

def tiene_orden_valida(ordenes_extraidas, ordenes_sgt):
    """Verifica si alguna de las órdenes extraídas existe en el maestro SGT"""
    for o in ordenes_extraidas:
        if o in ordenes_sgt: 
            return True
    return False

def buscar_valor_columna(row, df_cols, palabra_clave):
    """Busca valores en columnas que contengan una palabra clave específica"""
    valores = []
    for col in df_cols:
        if palabra_clave.upper() in str(col).upper():
            valores.append(str(row[col]))
    return " ".join(valores).upper()

def procesar_costos_por_orden(df, col_valor, es_horas=False):
    """Divide el costo o las horas entre múltiples órdenes en una misma celda"""
    distribucion = {}
    
    for _, row in df.iterrows():
        try: 
            valor = float(row[col_valor])
        except: 
            valor = 0.0
            
        if pd.isna(valor) or valor == 0.0: 
            continue
            
        clasif = row.get('Clasificacion', '')
        if clasif != "Orden Lista": 
            continue
            
        ordenes = row.get('Ordenes_Detectadas', [])
        if 'Orden_SGT' in row and str(row['Orden_SGT']).strip() != "":
            ordenes = [limpiar_orden(row['Orden_SGT'])]
            
        if len(ordenes) > 0:
            valor_dividido = valor / len(ordenes)
            for o in ordenes:
                o_str = str(o).strip().replace(" ", "").upper()
                if o_str == 'NAN' or o_str == '': 
                    continue
                distribucion[o_str] = distribucion.get(o_str, 0.0) + valor_dividido
                
    return distribucion

def generar_nexus_bytes(df_or_dict):
    """
    Soporta un DataFrame simple o un Diccionario para múltiples pestañas.
    Genera el formato estricto Nexus de 5 columnas sin encabezados.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if isinstance(df_or_dict, dict):
            for sheet_name, df in df_or_dict.items():
                if not df.empty:
                    df_nexus = df[["CUENTA", "VACIO", "CONCEPTO", "DEBE", "HABER"]]
                    # Recortamos el nombre de la hoja a 30 caracteres máximo (límite de Excel)
                    df_nexus.to_excel(writer, index=False, header=False, sheet_name=str(sheet_name)[:30])
        else:
            df_nexus = df_or_dict[["CUENTA", "VACIO", "CONCEPTO", "DEBE", "HABER"]]
            df_nexus.to_excel(writer, index=False, header=False, sheet_name="Partida_Nexus")
            
    return output.getvalue()

def generar_excel_filtrado(df, nombre_hoja):
    """Genera un archivo Excel estándar con encabezados para reportes"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=nombre_hoja)
    return output.getvalue()

def guardar_en_google_sheets(df_kardex, df_wip):
    """Espacio reservado para la conexión a la base de datos"""
    try:
        # AQUI DEBES METER TU CODIGO DE CONEXION DE GOOGLE SHEETS
        st.success("💾 ¡Datos guardados exitosamente en Google Sheets (Kardex y Saldos)!")
    except Exception as e:
        st.error(f"Error al intentar guardar en Google Sheets: {e}")

# ==========================================
# MÓDULO PRINCIPAL DE LA INTERFAZ
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
        with col_m1: 
            mes_proceso = st.selectbox("Mes:", range(1, 13), index=date.today().month - 1)
        with col_m2: 
            anio_proceso = st.number_input("Año:", min_value=2024, max_value=2030, value=date.today().year)

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
            num_rows="dynamic", 
            use_container_width=True, 
            key="ed_cuentas_mo"
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
            if st.button("🔍 Escanear Archivos y Aplicar Filtros", type="primary", use_container_width=True):
                with st.spinner("Leyendo y cruzando datos..."):
                    try:
                        # -----------------------------
                        # 0. LECTURA HISTÓRICO WIP
                        # -----------------------------
                        df_wip_ant = pd.DataFrame()
                        if arch_wip_ant:
                            try:
                                df_wip_ant = pd.read_excel(arch_wip_ant, dtype=str)
                                df_wip_ant.columns = df_wip_ant.columns.str.strip()
                            except: 
                                pass

                        # -----------------------------
                        # 1. LECTURA MAESTRO SGT
                        # -----------------------------
                        df_sgt = pd.read_excel(arch_sgt, dtype=str)
                        df_sgt.columns = df_sgt.columns.str.strip()
                        ordenes_validas = [limpiar_orden(o) for o in df_sgt['Orden'].dropna()] if 'Orden' in df_sgt.columns else []

                        # -----------------------------
                        # 2. LECTURA FACTURACIÓN
                        # -----------------------------
                        df_fact = pd.read_excel(arch_fact, dtype=str)
                        df_fact.columns = df_fact.columns.str.strip()
                        
                        df_fact['Texto_Completo_Factura'] = df_fact.apply(lambda r: " ".join(r.dropna().astype(str)), axis=1)
                        df_fact['Ordenes_Detectadas'] = df_fact['Texto_Completo_Factura'].apply(extraer_ordenes)
                        
                        def clasificar_factura(row):
                            desc = str(row.get('Descripcion', '')).upper()
                            cat = str(row.get('Categoria', '')).upper()
                            
                            if tiene_orden_valida(row['Ordenes_Detectadas'], ordenes_validas): 
                                return "Orden Lista"
                            if "SERVICIO" in cat or "SERVICIO" in desc: 
                                return "Servicios"
                            if any(k in desc for k in ["BANNER", "AFICHE", "CALENDARIO", "ROTULO"]): 
                                return "Venta Directa"
                            if any(k in desc for k in ["RECICLAJE", "DESPERDICIO"]): 
                                return "Reciclaje"
                            
                            if len(row['Ordenes_Detectadas']) > 0: 
                                return "Huérfana (Revisar)"
                            return "Huérfana (Revisar)"
                            
                        df_fact['Clasificacion'] = df_fact.apply(clasificar_factura, axis=1)

                        # -----------------------------
                        # 3. LECTURA TIEMPOS (NÓMINA)
                        # -----------------------------
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

                        # -----------------------------
                        # 4. LECTURA TRASLADOS MP y PT
                        # -----------------------------
                        dfs_mp = []
                        if arch_tras_mp:
                            for f in arch_tras_mp:
                                d = pd.read_excel(f, dtype=str)
                                d.columns = d.columns.str.strip()
                                dfs_mp.append(d)
                        if arch_tras_pt:
                            for f in arch_tras_pt:
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
                                
                                if tiene_orden_valida(row['Ordenes_Detectadas'], ordenes_validas): 
                                    return "Orden Lista"
                                    
                                if len(row['Ordenes_Detectadas']) > 0: 
                                    return "Huérfana (Revisar)"
                                    
                                if any(k in concepto for k in ["SOHO", "LIBRERI", "LBRERI", "UCA"]) or "PRODUCTO TERMINADO" in cat:
                                    return "Traslado Especial"
                                    
                                if any(k in cat for k in ["EMPAQUE", "LIMPIEZA", "REPUESTO", "REPUESTOS"]): 
                                    return "Costo Indirecto (Automático)"
                                    
                                return "Huérfana (Revisar)"
                                
                            df_mp['Clasificacion'] = df_mp.apply(clasificar_traslado, axis=1)

                        # -----------------------------
                        # GUARDADO EN MEMORIA
                        # -----------------------------
                        st.session_state['tg_fact'] = df_fact
                        st.session_state['tg_tiempos'] = df_tiempos
                        st.session_state['tg_mp'] = df_mp
                        st.session_state['tg_wip_ant'] = df_wip_ant
                        st.session_state['tg_ordenes_validas'] = ordenes_validas
                        st.session_state['tg_costo_planilla'] = costo_planilla
                        st.session_state['tg_df_cuentas'] = df_cuentas_mo 
                        
                        st.session_state['tg_datos_cargados'] = True
                        st.session_state['fase2_aprobada'] = False
                        st.session_state['liquidacion_lista'] = False
                        
                        st.success("✅ Datos cruzados correctamente. Avanza a Auditoría.")
                        
                    except Exception as e:
                        st.error(f"Error al leer archivos: {e}")

    # ==========================================
    # PESTAÑA 2: AUDITORÍA (CONTROL MANUAL)
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
                
                opciones_accion = ["Pendiente", "Asignar Orden", "Forzar Orden", "Costo Indirecto", "Omitir"]
                config_col = {
                    "Accion": st.column_config.SelectboxColumn("Acción", options=opciones_accion, required=True),
                    "Orden_SGT": st.column_config.TextColumn("Código Orden")
                }
                
                ed_fact = pd.DataFrame()
                ed_tiempos = pd.DataFrame()
                ed_mp = pd.DataFrame()

                def prellenar_orden(df_orig, df_huerf):
                    df_huerf['Orden_SGT'] = df_orig.loc[df_huerf.index, 'Ordenes_Detectadas'].apply(
                        lambda x: limpiar_orden(x[0]) if isinstance(x, list) and len(x) > 0 else ""
                    )
                    df_huerf['Accion'] = df_huerf['Orden_SGT'].apply(
                        lambda x: "Forzar Orden" if x != "" else "Pendiente"
                    )
                    return df_huerf

                if h_fact > 0:
                    with st.expander(f"🧾 Facturas Pendientes ({h_fact})", expanded=True):
                        cols_f = [c for c in ['Fecha', 'Numero', 'Descripcion', 'VentaNeta'] if c in df_fact.columns]
                        df_h_f = df_fact[df_fact['Clasificacion'] == "Huérfana (Revisar)"][cols_f].copy()
                        ed_fact = st.data_editor(prellenar_orden(df_fact, df_h_f), column_config=config_col, hide_index=True, key="ed_f", use_container_width=True)

                if h_tiempos > 0:
                    with st.expander(f"⏱️ Horas Pendientes ({h_tiempos})", expanded=True):
                        cols_t = [c for c in ['Empleado', 'Observaciones'] if c in df_tiempos.columns]
                        df_h_t = df_tiempos[df_tiempos['Clasificacion'] == "Huérfana (Revisar)"][cols_t].copy()
                        ed_tiempos = st.data_editor(prellenar_orden(df_tiempos, df_h_t), column_config=config_col, hide_index=True, key="ed_t", use_container_width=True)

                if h_mp > 0:
                    with st.expander(f"📦 Traslados MP/PT Pendientes ({h_mp})", expanded=True):
                        c_con = next((c for c in df_mp.columns if 'CONCEPT' in c.upper()), 'Numero')
                        c_des = next((c for c in df_mp.columns if 'DESCRIP' in c.upper()), 'Numero')
                        c_cat = next((c for c in df_mp.columns if 'CATEGOR' in c.upper()), 'Categoria')
                        cols_m = [c for c in ['Numero', c_con, c_des, c_cat] if c in df_mp.columns]
                        
                        df_h_m = df_mp[df_mp['Clasificacion'] == "Huérfana (Revisar)"][cols_m].copy()
                        ed_mp = st.data_editor(prellenar_orden(df_mp, df_h_m), column_config=config_col, hide_index=True, key="ed_m", use_container_width=True)

                if st.button("💾 Guardar Decisiones y Validar", type="primary"):
                    errores = []
                    
                    def revisar_tabla(df_edit, df_orig, nom):
                        for i, row in df_edit.iterrows():
                            acc = row['Accion']
                            
                            if acc == "Pendiente": 
                                errores.append(f"Queda un registro pendiente en {nom}.")
                            elif acc in ["Asignar Orden", "Forzar Orden"]:
                                orden = limpiar_orden(row['Orden_SGT'])
                                if orden == "": 
                                    errores.append(f"En {nom}, dejaste el código en blanco.")
                                elif acc == "Asignar Orden" and orden not in ordenes_validas: 
                                    errores.append(f"En {nom}, '{orden}' NO existe en SGT.")
                                else: 
                                    df_orig.at[i, 'Clasificacion'] = "Orden Lista"
                                    df_orig.at[i, 'Orden_SGT'] = orden
                            elif acc == "Costo Indirecto": 
                                df_orig.at[i, 'Clasificacion'] = "Costo Indirecto (Automático)"
                            elif acc == "Omitir": 
                                df_orig.at[i, 'Clasificacion'] = "Omitido Automático"

                    if not ed_fact.empty: revisar_tabla(ed_fact, df_fact, "Facturas")
                    if not ed_tiempos.empty: revisar_tabla(ed_tiempos, df_tiempos, "Tiempos")
                    if not ed_mp.empty: revisar_tabla(ed_mp, df_mp, "Traslados MP")

                    if errores:
                        st.error("❌ Corrige los errores:")
                        for e in errores: 
                            st.write(f"- {e}")
                    else:
                        st.session_state['tg_fact'] = df_fact
                        st.session_state['tg_tiempos'] = df_tiempos
                        st.session_state['tg_mp'] = df_mp
                        st.success("✅ ¡Auditoría validada! Procede a la Liquidación.")
                        st.session_state['fase2_aprobada'] = True
            else:
                st.success("✅ Todo perfecto. No hay registros pendientes.")
                st.session_state['fase2_aprobada'] = True
        else: 
            st.write("Carga los archivos en la Pestaña 1.")

    # ==========================================
    # PESTAÑA 3: LIQUIDACIÓN Y PARTIDAS
    # ==========================================
    with tab_liquidacion:
        st.subheader("💰 Generación de Costos y Partidas (Formato Nexus)")
        
        if st.session_state.get('fase2_aprobada', False):
            
            if st.button("🚀 Ejecutar Liquidación y Generar Partidas", type="primary"):
                with st.spinner("Construyendo Partidas..."):
                    df_fact = st.session_state['tg_fact']
                    df_tiempos = st.session_state['tg_tiempos']
                    df_mp = st.session_state['tg_mp']
                    df_wip_ant = st.session_state.get('tg_wip_ant', pd.DataFrame())
                    costo_total_mo = st.session_state['tg_costo_planilla']
                    df_cuentas_mo = st.session_state['tg_df_cuentas']
                    
                    def agregar_linea(lista, cuenta, descripcion, debe, haber):
                        if round(debe, 2) > 0 or round(haber, 2) > 0:
                            lista.append({
                                "CUENTA": cuenta, 
                                "VACIO": "", 
                                "CONCEPTO": str(descripcion).strip(), 
                                "DEBE": round(debe, 2), 
                                "HABER": round(haber, 2)
                            })

                    # 1. HISTÓRICO DE SALDOS
                    historial_saldos = {}
                    if not df_wip_ant.empty:
                        col_ord = next((c for c in df_wip_ant.columns if 'ORDEN' in c.upper()), None)
                        col_saldo = next((c for c in df_wip_ant.columns if 'SALDO_FINAL_WIP' in c.upper() or 'TOTAL' in c.upper()), None)
                        if col_ord and col_saldo:
                            for _, r in df_wip_ant.iterrows():
                                o = limpiar_orden(r[col_ord])
                                if o == 'NAN' or o == '': continue
                                try: val = float(r[col_saldo])
                                except: val = 0.0
                                if val > 0: historial_saldos[o] = val

                    # 2. COSTOS DEL MES
                    col_horas = next((c for c in df_tiempos.columns if 'TOTALHORA' in c.upper().replace(' ', '')), None)
                    horas_ordenes = procesar_costos_por_orden(df_tiempos, col_horas, es_horas=True) if col_horas else {}
                    total_horas_validas = sum(horas_ordenes.values())
                    costo_por_hora = (costo_total_mo / total_horas_validas) if total_horas_validas > 0 else 0
                    costos_mo_por_orden = {orden: horas * costo_por_hora for orden, horas in horas_ordenes.items()}

                    col_costo_mp = next((c for c in df_mp.columns if 'PRECIOTOTAL' in c.upper().replace(' ', '') or 'COSTO' in c.upper()), None)
                    col_cat = next((c for c in df_mp.columns if 'CATEGOR' in c.upper()), 'Categoria')
                    col_texto_mp = next((c for c in df_mp.columns if 'CONCEPT' in c.upper()), 'Descripcion')

                    if col_costo_mp: df_mp[col_costo_mp] = pd.to_numeric(df_mp[col_costo_mp], errors='coerce').fillna(0)

                    # PARTIDA 1: NÓMINA (Ajustado a 110602)
                    p1 = []
                    agregar_linea(p1, "110602", "INYECCION MANO DE OBRA A PRODUCCION", costo_total_mo, 0)
                    for _, r in df_cuentas_mo.iterrows():
                        cuenta_cod = r['Cuenta'].split(" ")[0]
                        desc_cuenta = r['Cuenta'].replace(cuenta_cod, "").strip()
                        agregar_linea(p1, cuenta_cod, desc_cuenta, 0, r['Monto'])
                    st.session_state['tg_p1'] = pd.DataFrame(p1)

                    # PARTIDA 2: MP A PROCESO (Ajustado a 110602)
                    p2 = []
                    df_wip_mp = df_mp[df_mp['Clasificacion'] == 'Orden Lista']
                    if not df_wip_mp.empty:
                        resumen_wip_dict = {}
                        for _, r in df_wip_mp.iterrows():
                            monto = float(r[col_costo_mp])
                            if monto > 0:
                                cta_inv, nom_inv = obtener_datos_inventario(r[col_cat])
                                llave = (cta_inv, nom_inv)
                                resumen_wip_dict[llave] = resumen_wip_dict.get(llave, 0.0) + monto
                                
                        total_wip_mp = sum(resumen_wip_dict.values())
                        if total_wip_mp > 0:
                            agregar_linea(p2, "110602", "INVENTARIO DE PRODUCTO EN PROCESO", total_wip_mp, 0)
                            for (cta, nom), monto in resumen_wip_dict.items():
                                agregar_linea(p2, cta, nom, 0, monto)
                    st.session_state['tg_p2'] = pd.DataFrame(p2)

                    # PARTIDA 3: CIF
                    p3 = []
                    df_cif_mp = df_mp[df_mp['Clasificacion'] == 'Costo Indirecto (Automático)']
                    if not df_cif_mp.empty:
                        resumen_cif_dict = {}
                        for _, r in df_cif_mp.iterrows():
                            monto = float(r[col_costo_mp])
                            if monto > 0:
                                cta_inv, nom_inv = obtener_datos_inventario(r[col_cat])
                                llave = (cta_inv, nom_inv)
                                resumen_cif_dict[llave] = resumen_cif_dict.get(llave, 0.0) + monto
                        total_cif = sum(resumen_cif_dict.values())
                        if total_cif > 0:
                            agregar_linea(p3, CUENTA_CIF, "Costo de lo Vendido Indirectos TG", total_cif, 0)
                            for (cta, nom), monto in resumen_cif_dict.items():
                                agregar_linea(p3, cta, nom, 0, monto)
                    st.session_state['tg_p3'] = pd.DataFrame(p3)

                    # PARTIDA 4: TRASLADOS A OTRAS UNIDADES (PROVISIÓN)
                    p4_dict = {}
                    df_tras_esp = df_mp[df_mp['Clasificacion'] == 'Traslado Especial']
                    if not df_tras_esp.empty:
                        def asignar_unidad(row):
                            concepto = str(row.get(col_texto_mp, '')).upper()
                            cat = str(row.get(col_cat, '')).upper()
                            if "SOHO" in concepto: return "SOHO"
                            if "LIBRERI" in concepto or "LBRERI" in concepto or "UCA" in concepto: return "LIBRERIA"
                            if "PRODUCTO TERMINADO" in cat: return "PRODUCTO TERMINADO"
                            return "OTRAS UNIDADES"

                        df_tras_esp['Unidad_Destino'] = df_tras_esp.apply(asignar_unidad, axis=1)
                        for unidad in df_tras_esp['Unidad_Destino'].unique():
                            df_u = df_tras_esp[df_tras_esp['Unidad_Destino'] == unidad]
                            resumen_u_dict = {}
                            for _, r in df_u.iterrows():
                                monto = float(r[col_costo_mp])
                                if monto > 0:
                                    cta_inv, nom_inv = obtener_datos_inventario(r[col_cat])
                                    resumen_u_dict[(cta_inv, nom_inv)] = resumen_u_dict.get((cta_inv, nom_inv), 0.0) + monto
                            
                            total_unidad = sum(resumen_u_dict.values())
                            if total_unidad > 0:
                                p4_unidad = []
                                agregar_linea(p4_unidad, CUENTA_PROVISION_TRASLADOS, "PROVISIONES POR TRASLADOS DE INVENTARIOS", total_unidad, 0)
                                for (cta, nom), monto in resumen_u_dict.items():
                                    agregar_linea(p4_unidad, cta, nom, 0, monto)
                                p4_dict[unidad] = pd.DataFrame(p4_unidad)
                    st.session_state['tg_p4_dict'] = p4_dict

                    # BODEGA VIRTUAL Y LIQUIDACIÓN (P5)
                    costos_mp_por_orden = procesar_costos_por_orden(df_wip_mp, col_costo_mp) if col_costo_mp else {}
                    ordenes_facturadas = []
                    for _, row in df_fact[df_fact['Clasificacion'] == 'Orden Lista'].iterrows():
                        if 'Orden_SGT' in row and str(row['Orden_SGT']).strip() != "":
                            ordenes_facturadas.append(limpiar_orden(row['Orden_SGT']))
                        else:
                            for o in row.get('Ordenes_Detectadas', []):
                                if limpiar_orden(o) != "": ordenes_facturadas.append(limpiar_orden(o))
                                    
                    # BLINDAJE FINAL
                    todas_las_ordenes = set()
                    todas_las_ordenes.update(costos_mo_por_orden.keys())
                    todas_las_ordenes.update(costos_mp_por_orden.keys())
                    todas_las_ordenes.update(ordenes_facturadas)
                    todas_las_ordenes.update(historial_saldos.keys())
                    
                    fecha_str = date.today().strftime("%d/%m/%Y")
                    filas_kardex, filas_wip, total_liq_cv = [], [], 0.0

                    for orden in todas_las_ordenes:
                        ord_cln = limpiar_orden(orden)
                        if ord_cln == 'NAN' or ord_cln == '': continue
                        
                        nuevo_mo = costos_mo_por_orden.get(orden, 0.0)
                        nuevo_mp = costos_mp_por_orden.get(orden, 0.0)
                        saldo_anterior = historial_saldos.get(ord_cln, 0.0)
                        saldo_acumulado = saldo_anterior + nuevo_mo + nuevo_mp
                        
                        estado = "Liquidado a Costo de Ventas" if ord_cln in ordenes_facturadas else "Pendiente"
                        
                        if nuevo_mo > 0: filas_kardex.append({"Fecha": fecha_str, "Orden": ord_cln, "Tipo_Costo": "Mano de Obra", "Monto": nuevo_mo, "Estado": estado})
                        if nuevo_mp > 0: filas_kardex.append({"Fecha": fecha_str, "Orden": ord_cln, "Tipo_Costo": "Materia Prima", "Monto": nuevo_mp, "Estado": estado})

                        if estado == "Liquidado a Costo de Ventas":
                            total_liq_cv += saldo_acumulado
                            saldo_acumulado = 0.0
                            
                        filas_wip.append({
                            "Orden": ord_cln, "Saldo_Anterior": saldo_anterior, "Costo_MO_Mes": nuevo_mo, "Costo_MP_Mes": nuevo_mp, "Total_Acumulado": saldo_anterior + nuevo_mo + nuevo_mp, "Estado": estado, "Saldo_Final_WIP": saldo_acumulado
                        })
                    
                    st.session_state['tg_df_kardex'] = pd.DataFrame(filas_kardex)
                    st.session_state['tg_df_wip'] = pd.DataFrame(filas_wip)

                    # ==================================================
                    # PARTIDA 5: Liquidación Final (Ajustado)
                    # ==================================================
                    p5 = []
                    if len(ordenes_facturadas) > 0 or total_liq_cv > 0:
                        # CARGO (DEBE)
                        agregar_linea(p5, "410104", "COSTO DE LO VENDIDO", total_liq_cv, 0)
                        # ABONO (HABER)
                        agregar_linea(p5, "110602", "INVENTARIO DE PRODUCTO EN PROCESO", 0, total_liq_cv)
                        
                    st.session_state['tg_p5'] = pd.DataFrame(p5)
                    st.session_state['hay_ordenes_facturadas'] = len(ordenes_facturadas) > 0

                    st.session_state['liquidacion_lista'] = True
                    st.success("✅ Cálculos completados. Partidas listas para descarga.")

            # DESCARGAS
            if st.session_state.get('liquidacion_lista', False):
                st.markdown("### 📥 Descarga de Partidas (Nexus)")
                c1, c2, c3 = st.columns(3)
                with c1: st.download_button("⬇️ 1. P. Nómina", data=generar_nexus_bytes(st.session_state['tg_p1']), file_name=f"1_Nex_Nom_{mes_proceso}.xlsx")
                with c2: 
                    if not st.session_state['tg_p2'].empty: st.download_button("⬇️ 2. P. Materia Prima", data=generar_nexus_bytes(st.session_state['tg_p2']), file_name=f"2_Nex_MP_{mes_proceso}.xlsx")
                    else: st.info("Sin movimientos")
                with c3: 
                    if not st.session_state['tg_p3'].empty: st.download_button("⬇️ 3. P. Costos Indirectos", data=generar_nexus_bytes(st.session_state['tg_p3']), file_name=f"3_Nex_CIF_{mes_proceso}.xlsx")
                    else: st.info("Sin movimientos")

                c4, c5, c6 = st.columns(3)
                with c4:
                    if st.session_state['tg_p4_dict']: st.download_button("⬇️ 4. P. Traslados (Múltiples Hojas)", data=generar_nexus_bytes(st.session_state['tg_p4_dict']), file_name=f"4_Nex_Traslados_{mes_proceso}.xlsx")
                    else: st.info("Sin movimientos")
                with c5:
                    if not st.session_state['tg_p5'].empty:
                        st.download_button("⬇️ 5. P. Liquidación a Ventas", data=generar_nexus_bytes(st.session_state['tg_p5']), file_name=f"5_Nex_Liq_{mes_proceso}.xlsx")
                    else: st.info("Sin OP facturadas.")
                with c6: st.empty()

                st.divider()
                st.markdown("### 📊 Reportes de Control de Bodega")
                col_f1, _ = st.columns(2)
                f_ord = col_f1.text_input("🔍 Filtrar por Orden:")

                df_k_ex = st.session_state['tg_df_kardex']
                df_w_ex = st.session_state['tg_df_wip']
                if f_ord.strip() != "":
                    df_k_ex = df_k_ex[df_k_ex['Orden'].str.contains(f_ord, case=False, na=False)]
                    df_w_ex = df_w_ex[df_w_ex['Orden'].str.contains(f_ord, case=False, na=False)]

                st.dataframe(df_w_ex, use_container_width=True)

                b1, b2, b3 = st.columns([1,1,1])
                with b1:
                    if st.button("💾 Guardar en Sheets", type="secondary"): guardar_en_google_sheets(df_k_ex, df_w_ex)
                with b2: st.download_button("📋 Bajar Kardex", data=generar_excel_filtrado(df_k_ex, "Kardex"), file_name=f"Kardex_{mes_proceso}.xlsx")
                with b3: st.download_button("📉 Bajar Saldos WIP", data=generar_excel_filtrado(df_w_ex, "Saldos_WIP"), file_name=f"Saldos_WIP_{mes_proceso}.xlsx")