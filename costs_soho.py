import streamlit as st
import pandas as pd
import io
from datetime import date

# ==========================================
# CONSTANTES CONTABLES
# ==========================================
CTA_PROVISION_PUENTE = "21020302" 
CTA_CONSIGNACION_DR_ENTRADA = "7101" 
CTA_CONSIGNACION_CR_ENTRADA = "8101" 
CTA_CONSIGNACION_DR_SALIDA = "8101" 
CTA_CONSIGNACION_CR_SALIDA = "7101" 

CTA_BASE_GASTO = "410104" 
CTA_BASE_PT = "110603" 
CTA_BASE_MP = "110601" 

# Bodegas oficiales de la unidad actual
UNIDADES_SOHO_SET = {
    "CENTRO SOHO", "BODEGA SOHO", "SOHO"
}

# Lista maestra con y sin tildes para evitar duplicar entradas internas
OTRAS_UNIDADES_INTERNAS = {
    "LIBRERÍA CENTRAL", "LIBRERIA CENTRAL", 
    "LIBRERÍA BODEGA", "LIBRERIA BODEGA",
    "LIBRERÍA BODEGA TG", "LIBRERIA BODEGA TG", 
    "LIBRERÍA CAMPUS", "LIBRERIA CAMPUS", 
    "LIBRERÍA CENTRAL (B001)", "LIBRERIA CENTRAL (B001)",
    "TERRAZA", "CID CAMPUS", "BODEGA CID CAMPUS", 
    "DESPENSA", "CAFETERIA", "CAFETERIA CENTRAL",
    "TALLERES", "GERENCIA COMERCIAL"
}

INV_ACCOUNT_MAP = {
    'MATERIA PRIMA': CTA_BASE_MP,
    'PRODUCTO TERMINADO': CTA_BASE_PT,
    'DEFAULT': CTA_BASE_PT
}

# ==========================================
# FUNCIONES DE ESTRUCTURA
# ==========================================
def gen_nexus_spec(cuenta, concepto, debe, haber, auxiliar=""):
    return {
        "CUENTA": str(cuenta),
        "VACIO": "",
        "CONCEPTO": str(concepto).strip(),
        "AUXILIAR": str(auxiliar).strip(),
        "DEBE": float(round(debe, 2)),
        "HABER": float(round(haber, 2))
    }

def clean_value(val):
    return str(val).strip().upper() if not pd.isna(val) else ""

def gen_excels_memory(partidas_dict, mes_proceso):
    buffers = {}
    nombres_amigables = {
        "TRASLADOS": "Traslados_y_Recepciones",
        "AJUSTES": "Ajustes_de_Inventario",
        "CONSIGNACION": "Movimientos_Consignacion",
        "VENTAS_CONSIGNACION": "Ventas_Consignacion"
    }
    
    for categoria, tabs_dict in partidas_dict.items():
        if not any(len(lst) > 0 for lst in tabs_dict.values()): continue
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, partidas_list in tabs_dict.items():
                if not partidas_list: continue
                df_nexus = pd.DataFrame(partidas_list)
                
                df_grouped = df_nexus.groupby(["CUENTA", "VACIO", "CONCEPTO", "AUXILIAR"], as_index=False)[["DEBE", "HABER"]].sum()
                df_grouped = df_grouped[(df_grouped["DEBE"] > 0) | (df_grouped["HABER"] > 0)]
                df_grouped = df_grouped.sort_values(by=["CONCEPTO", "HABER", "DEBE"], ascending=[True, True, False])
                
                safe_sheet = str(sheet_name).replace("/", "-")[:31]
                df_grouped[["CUENTA", "VACIO", "CONCEPTO", "DEBE", "HABER"]].to_excel(writer, index=False, header=False, sheet_name=safe_sheet)
                
        buffers[f"Partidas_{nombres_amigables.get(categoria, categoria)}_Mes_{mes_proceso}.xlsx"] = output.getvalue()
    return buffers

# ==========================================
# INTERFAZ Y PROCESAMIENTO
# ==========================================
def mostrar_modulo_soho():
    st.title("🏢 Módulo de Inventarios - Centro Soho")
    st.info("Automatización de Partidas Contables de Inventario y Ventas en Consignación.")

    tab_carga, tab_liquidacion = st.tabs(["📥 1. Carga de Datos", "💰 2. Partidas y Nexus"])

    with tab_carga:
        st.subheader("Configuración del Periodo")
        col_m1, col_m2 = st.columns(2)
        with col_m1: mes_proceso = st.selectbox("Mes:", range(1, 13), index=date.today().month - 1, key="m_soho")
        with col_m2: anio_proceso = st.number_input("Año:", min_value=2024, value=date.today().year, key="y_soho")

        st.markdown("---")
        st.subheader("Archivos de Operación (Generación Nexus)")
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            arch_mov_inv = st.file_uploader("1. Reporte de Movimientos (Inventario)", type=["xls", "xlsx"], key="file_inv_soho")
        with col_f2:
            arch_ventas = st.file_uploader("2. Reporte de Ventas (Para Consignación)", type=["xls", "xlsx"], key="file_ven_soho")
            
        st.subheader("Archivos de Auditoría (Validación 1%)")
        col_a1, col_a2 = st.columns(2)
        with col_a1:
            arch_kardex = st.file_uploader("3. Kardex Consolidado", type=["xls", "xlsx"], key="file_kar_soho")
        with col_a2:
            arch_cat = st.file_uploader("4. Maestro Categorías (Opcional)", type=["xls", "xlsx"], key="file_cat_soho")

        btn_scan = st.button("🔍 Procesar Auditoría", type="primary", use_container_width=True, key="btn_soho")

        if btn_scan:
            if not arch_mov_inv and not arch_ventas and not arch_kardex:
                st.warning("⚠️ Debes subir al menos un archivo para procesar.")
                return
                
            with st.spinner("Analizando información..."):
                
                # ====================================================
                # NUEVO BLOQUE: AUDITORÍA DE KARDEX (1%) CON LÓGICA EN CASCADA
                # ====================================================
                if arch_kardex:
                    try:
                        st.markdown("---")
                        st.subheader("🕵️ Resultado de Auditoría Kardex")
                        
                        df_k = pd.read_excel(arch_kardex, dtype=str)
                        df_k.columns = df_k.columns.astype(str).str.strip().str.upper()
                        
                        # Búsqueda dinámica ultra-resistente
                        c_cod = next((c for c in df_k.columns if 'IDPRODUCTO' in c), None)
                        c_pref = next((c for c in df_k.columns if 'PREFI' in c), None)
                        c_doc = next((c for c in df_k.columns if 'DOCUMENTO' in c), None)
                        c_ent_u = next((c for c in df_k.columns if 'ENTRADASUNID' in c), None)
                        c_ent_v = next((c for c in df_k.columns if 'ENTRADASVAL' in c), None)
                        c_costo = next((c for c in df_k.columns if 'COSTOPROMEDIO' in c), None)

                        if not all([c_cod, c_pref, c_ent_u, c_ent_v, c_costo]):
                            st.error("Error: No se encontraron todas las columnas clave en el Kardex (IdProducto, Prefijo, EntradasUnidades, EntradasValor, CostoPromedio).")
                        else:
                            df_k[c_cod] = df_k[c_cod].astype(str).str.strip()

                            # Filtro por Categorías usando Diccionario
                            if arch_cat:
                                df_c = pd.read_excel(arch_cat, dtype=str)
                                df_c.columns = df_c.columns.astype(str).str.strip()
                                c_cod_cat = df_c.columns[0]
                                c_desc_cat = df_c.columns[1] if len(df_c.columns) > 1 else df_c.columns[0]
                                mapa_categorias = dict(zip(
                                    df_c[c_cod_cat].astype(str).str.strip(), 
                                    df_c[c_desc_cat].astype(str).str.upper().str.strip()
                                ))
                                df_k['Cat_Temp'] = df_k[c_cod].map(mapa_categorias).fillna('DESCONOCIDA')
                                df_k = df_k[df_k['Cat_Temp'] != 'SERVICIO']

                            # Limpieza matemática
                            df_k[c_ent_u] = pd.to_numeric(df_k[c_ent_u], errors='coerce').fillna(0)
                            df_k[c_ent_v] = pd.to_numeric(df_k[c_ent_v], errors='coerce').fillna(0)
                            df_k[c_costo] = pd.to_numeric(df_k[c_costo], errors='coerce').fillna(0)
                            
                            df_k['Prefijo_Upper'] = df_k[c_pref].astype(str).str.upper().str.strip()
                            df_k['Doc_Upper'] = df_k[c_doc].astype(str).str.upper().str.strip() if c_doc else ""

                            # LÓGICA EN CASCADA PARA COSTO BASE
                            
                            # 1. CFE (Compras)
                            df_cfe = df_k[df_k['Prefijo_Upper'] == 'CFE']
                            ref_compra = df_cfe.groupby(c_cod).apply(
                                lambda x: x[c_ent_v].sum() / x[c_ent_u].sum() if x[c_ent_u].sum() != 0 else 0
                            ).reset_index(name='Costo_Ref_CFE')

                            # 2. INI o Saldo Anterior
                            df_ini = df_k[(df_k['Prefijo_Upper'].isin(['INI', 'NAN', ''])) | (df_k['Doc_Upper'].str.contains('SALDO ANTERIOR'))]
                            ref_ini = df_ini.groupby(c_cod)[c_costo].first().reset_index(name='Costo_Ref_INI')

                            # 3. TRD (Traslados de Entrada)
                            df_trd = df_k[(df_k['Prefijo_Upper'] == 'TRD') & (df_k[c_ent_u] > 0)]
                            ref_trd = df_trd.groupby(c_cod).apply(
                                lambda x: x[c_ent_v].sum() / x[c_ent_u].sum() if x[c_ent_u].sum() != 0 else 0
                            ).reset_index(name='Costo_Ref_TRD')

                            # Unificar bases
                            ref_base = pd.DataFrame({c_cod: df_k[c_cod].unique()})
                            ref_base = pd.merge(ref_base, ref_compra, on=c_cod, how='left')
                            ref_base = pd.merge(ref_base, ref_ini, on=c_cod, how='left')
                            ref_base = pd.merge(ref_base, ref_trd, on=c_cod, how='left')
                            
                            # Cascada: Toma CFE, si no hay toma INI, si no hay toma TRD.
                            ref_base['Costo_Base'] = ref_base['Costo_Ref_CFE'].fillna(ref_base['Costo_Ref_INI']).fillna(ref_base['Costo_Ref_TRD']).fillna(0)

                            # Costo de Venta (Promedio de FCF / CCF)
                            df_ventas_k = df_k[df_k['Prefijo_Upper'].isin(['FCF', 'CCF'])]
                            costo_ventas = df_ventas_k.groupby(c_cod)[c_costo].mean().reset_index(name='Costo_Venta_Promedio')

                            # Cruzar y Validar Variación
                            df_audit = pd.merge(costo_ventas, ref_base[[c_cod, 'Costo_Base']], on=c_cod, how='inner')
                            
                            # Filtro Antiruido: Quitar si ambos son 0 para evitar el -100% absurdo
                            df_audit = df_audit[(df_audit['Costo_Base'] > 0) | (df_audit['Costo_Venta_Promedio'] > 0)]

                            df_audit['Costo_Base_Safe'] = df_audit['Costo_Base'].replace(0, 1) # Evitar división por cero
                            df_audit['Variacion_%'] = (df_audit['Costo_Venta_Promedio'] / df_audit['Costo_Base_Safe']) - 1

                            anomalias = df_audit[df_audit['Variacion_%'].abs() > 0.01].copy()

                            if not anomalias.empty:
                                st.error(f"🚨 ALERTA: Se detectaron {len(anomalias)} productos con variación mayor al 1%.")
                                anomalias_show = anomalias[[c_cod, 'Costo_Base', 'Costo_Venta_Promedio', 'Variacion_%']].copy()
                                anomalias_show.columns = ['Código', 'Costo Base (CFE/INI/TRD)', 'Costo Prom. Ventas', 'Variación']
                                st.dataframe(anomalias_show.style.format({
                                    'Costo Base (CFE/INI/TRD)': '${:.4f}',
                                    'Costo Prom. Ventas': '${:.4f}',
                                    'Variación': '{:.2%}'
                                }), use_container_width=True)
                            else:
                                st.success("✅ Validación del 1% exitosa. Los costos de venta cuadran con las compras/saldos.")

                    except Exception as e_k:
                        st.error(f"Error procesando el Kardex: {e_k}")

                # ====================================================
                # PROCESAMIENTO 1: ARCHIVO DE INVENTARIOS
                # ====================================================
                try:
                    partidas_dict = {
                        "TRASLADOS": {},
                        "AJUSTES": {"Entradas_por_Ajuste": [], "Salidas_por_Ajuste": []},
                        "CONSIGNACION": {},
                        "VENTAS_CONSIGNACION": {}
                    }

                    if arch_mov_inv:
                        df = pd.read_excel(arch_mov_inv, dtype=str)
                        df.columns = df.columns.str.strip()
                        c_tipo = next((c for c in df.columns if 'TIPO' in c.upper() or 'CONCEPT' in c.upper()), None)
                        c_sender = next((c for c in df.columns if 'SALIDA' in c.upper().replace(' ', '')), None)
                        c_receiver = next((c for c in df.columns if 'INGRESO' in c.upper().replace(' ', '')), None)
                        c_category = next((c for c in df.columns if 'CATEGOR' in c.upper()), 'Categoria')
                        c_total = next((c for c in df.columns if 'PRECIOTOTAL' in c.upper().replace(' ', '') or 'COSTO' in c.upper()), 'PrecioTotal')
                        
                        df[c_total] = pd.to_numeric(df[c_total], errors='coerce').fillna(0)

                        for _, row in df.iterrows():
                            raw_sender, raw_receiver = clean_value(row[c_sender]), clean_value(row[c_receiver])
                            tipo, category, total = clean_value(row[c_tipo]), clean_value(row[c_category]), float(row[c_total])
                            if total == 0: continue

                            is_sender_unidad = raw_sender in UNIDADES_SOHO_SET
                            is_receiver_unidad = raw_receiver in UNIDADES_SOHO_SET
                            
                            if not (is_sender_unidad or is_receiver_unidad) or (is_sender_unidad and is_receiver_unidad) or 'SERVICIO' in category: continue

                            inv_acc = INV_ACCOUNT_MAP.get(category, INV_ACCOUNT_MAP['DEFAULT'])
                            sender_desc = raw_sender if raw_sender else "ORIGEN"
                            receiver_desc = raw_receiver if raw_receiver else "DESTINO"
                            
                            es_traslado = 'TRASLAD' in tipo or (raw_sender and raw_receiver)
                            es_ajuste = 'AJUS' in tipo or not es_traslado

                            # BLOQUE 1: CONSIGNACIÓN (Inventario)
                            if 'CONSIGNACION' in category:
                                if is_receiver_unidad and not is_sender_unidad:
                                    if raw_sender in OTRAS_UNIDADES_INTERNAS or es_traslado:
                                        pass 
                                    else:
                                        desc = f"RECONOCIMIENTO POR ENTRADA DE PRODUCTOS EN CONSIGNACION DE CENTRO SOHO, MES {mes_proceso} DE {anio_proceso}."
                                        if "Entradas_Consignacion" not in partidas_dict["CONSIGNACION"]: partidas_dict["CONSIGNACION"]["Entradas_Consignacion"] = []
                                        partidas_dict["CONSIGNACION"]["Entradas_Consignacion"].extend([
                                            gen_nexus_spec(CTA_CONSIGNACION_DR_ENTRADA, desc, total, 0, raw_receiver),
                                            gen_nexus_spec(CTA_CONSIGNACION_CR_ENTRADA, desc, 0, total, raw_receiver)
                                        ])
                                elif is_sender_unidad:
                                    if es_traslado and receiver_desc != "DESTINO":
                                        tab_name = receiver_desc
                                        desc = f"RECONOCIMIENTO POR {tipo} DE PRODUCTOS EN CONSIGNACION HACIA {receiver_desc}, MES {mes_proceso} DE {anio_proceso}."
                                    else:
                                        tab_name = "Salidas_Consignacion"
                                        desc = f"RECONOCIMIENTO POR {tipo} DE PRODUCTOS EN CONSIGNACION DE CENTRO SOHO, MES {mes_proceso} DE {anio_proceso}."
                                    
                                    if tab_name not in partidas_dict["CONSIGNACION"]: partidas_dict["CONSIGNACION"][tab_name] = []
                                    
                                    p1 = gen_nexus_spec(CTA_CONSIGNACION_DR_ENTRADA, desc, total, 0, raw_sender)
                                    p2 = gen_nexus_spec(CTA_CONSIGNACION_CR_ENTRADA, desc, 0, total, raw_sender)
                                    p3 = gen_nexus_spec(CTA_CONSIGNACION_DR_SALIDA, desc, total, 0, raw_sender)
                                    p3["CONCEPTO"] += " "
                                    p4 = gen_nexus_spec(CTA_CONSIGNACION_CR_SALIDA, desc, 0, total, raw_sender)
                                    p4["CONCEPTO"] += " "
                                    
                                    partidas_dict["CONSIGNACION"][tab_name].extend([p1, p2, p3, p4])

                                    if 'AJUS' in tipo or not es_traslado:
                                        desc_ajuste_costo = f"RECONOCIMIENTO DE SALIDA POR AJUSTE DE PRODUCTO DE CENTRO SOHO, MES {mes_proceso} DE {anio_proceso}."
                                        partidas_dict["AJUSTES"]["Salidas_por_Ajuste"].extend([
                                            gen_nexus_spec(CTA_BASE_GASTO, desc_ajuste_costo, total, 0), 
                                            gen_nexus_spec(inv_acc, desc_ajuste_costo, 0, total)
                                        ])

                            # BLOQUE 2: AJUSTES REGULARES
                            elif es_ajuste:
                                if is_receiver_unidad:
                                    desc = f"RECONOCIMIENTO DE ENTRADA POR AJUSTE DE PRODUCTO DE CENTRO SOHO, MES {mes_proceso} DE {anio_proceso}."
                                    partidas_dict["AJUSTES"]["Entradas_por_Ajuste"].extend([gen_nexus_spec(inv_acc, total, 0), gen_nexus_spec(CTA_BASE_GASTO, desc, 0, total)])
                                else:
                                    desc = f"RECONOCIMIENTO DE SALIDA POR AJUSTE DE PRODUCTO DE CENTRO SOHO, MES {mes_proceso} DE {anio_proceso}."
                                    partidas_dict["AJUSTES"]["Salidas_por_Ajuste"].extend([gen_nexus_spec(CTA_BASE_GASTO, desc, total, 0), gen_nexus_spec(inv_acc, desc, 0, total)])

                            # BLOQUE 3: TRASLADOS ESTÁNDAR
                            elif is_sender_unidad and es_traslado:
                                if receiver_desc not in partidas_dict["TRASLADOS"]: partidas_dict["TRASLADOS"][receiver_desc] = []
                                d_s = f"RECONOCIMIENTO DE SALIDA POR TRASLADO DE PRODUCTO DE {sender_desc} HACIA {receiver_desc}, MES {mes_proceso} DE {anio_proceso}."
                                d_r = d_s.replace("SALIDA", "ENTRADA")
                                partidas_dict["TRASLADOS"][receiver_desc].extend([
                                    gen_nexus_spec(CTA_PROVISION_PUENTE, d_s, total, 0), gen_nexus_spec(inv_acc, d_s, 0, total),
                                    gen_nexus_spec(INV_ACCOUNT_MAP.get('DEFAULT'), d_r, total, 0), gen_nexus_spec(CTA_PROVISION_PUENTE, d_r, 0, total)
                                ])

                    # ====================================================
                    # PROCESAMIENTO 2: ARCHIVO DE VENTAS (NUEVO REQUERIMIENTO)
                    # ====================================================
                    if arch_ventas:
                        df_v = pd.read_excel(arch_ventas, dtype=str)
                        df_v.columns = df_v.columns.str.strip()
                        
                        c_cat_v = next((c for c in df_v.columns if 'CATEGOR' in c.upper()), None)
                        c_costo_v = next((c for c in df_v.columns if 'TOTALCOSTO' in c.upper().replace(' ', '') or 'COSTO' in c.upper()), None)
                        
                        if c_cat_v and c_costo_v:
                            df_v[c_costo_v] = pd.to_numeric(df_v[c_costo_v], errors='coerce').fillna(0)
                            
                            # Filtrar solo ventas de consignacion
                            df_consignacion_ventas = df_v[df_v[c_cat_v].astype(str).str.upper().str.strip() == 'CONSIGNACION']
                            total_ventas_consignacion = df_consignacion_ventas[c_costo_v].sum()
                            
                            if total_ventas_consignacion > 0:
                                desc_venta_consig = f"RECONOCIMIENTO POR VENTA DE PRODUCTOS EN CONSIGNACION DE CENTRO SOHO, MES {mes_proceso} DE {anio_proceso}."
                                desc_costo_consig = f"RECONOCIMIENTO DE COSTO POR VENTA DE PRODUCTOS EN CONSIGNACION DE CENTRO SOHO, MES {mes_proceso} DE {anio_proceso}."
                                
                                partidas_dict["VENTAS_CONSIGNACION"]["Ventas_Consignacion"] = [
                                    # Partida 1: Reversión de la consignación
                                    gen_nexus_spec(CTA_CONSIGNACION_DR_SALIDA, desc_venta_consig, total_ventas_consignacion, 0, "CENTRO SOHO"),
                                    gen_nexus_spec(CTA_CONSIGNACION_CR_SALIDA, desc_venta_consig, 0, total_ventas_consignacion, "CENTRO SOHO"),
                                    
                                    # Partida 2: Reconocimiento de Costo (EXCLUSIVO SOHO)
                                    gen_nexus_spec(CTA_BASE_GASTO, desc_costo_consig, total_ventas_consignacion, 0, "CENTRO SOHO"),
                                    gen_nexus_spec(CTA_BASE_PT, desc_costo_consig, 0, total_ventas_consignacion, "CENTRO SOHO")
                                ]
                        else:
                            st.warning("⚠️ No se encontraron las columnas 'Categoria' o 'TotalCosto' en el reporte de ventas.")

                    if arch_mov_inv or arch_ventas:
                        st.success("✅ Procesamiento contable finalizado con éxito.")
                        st.session_state['nexus_soho_buffers'] = gen_excels_memory(partidas_dict, mes_proceso)
                        st.session_state['tab_results_soho'] = True
                    
                except Exception as e: 
                    st.error(f"Error técnico en procesamiento contable: {e}")

    with tab_liquidacion:
        if st.session_state.get('tab_results_soho', False):
            for f, b in st.session_state['nexus_soho_buffers'].items():
                st.download_button(label=f"⬇️ Descargar {f}", data=b, file_name=f, key=f"dl_soho_{f}")