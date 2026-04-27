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
            
        st.subheader("Archivos de Auditoría (Validación 1% y $0.02)")
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
                
            with st.spinner("Analizando información y filtrando por fecha..."):
                
                # ====================================================
                # BLOQUE: AUDITORÍA DE KARDEX CON FILTRO DOBLE
                # ====================================================
                if arch_kardex:
                    try:
                        st.markdown("---")
                        st.subheader("🕵️ Resultado de Auditoría Kardex (Detallada)")
                        
                        df_k = pd.read_excel(arch_kardex, dtype=str)
                        df_k.columns = df_k.columns.astype(str).str.strip().str.upper()
                        
                        c_cod = next((c for c in df_k.columns if 'IDPRODUCTO' in c), None)
                        c_pref = next((c for c in df_k.columns if 'PREFI' in c), None)
                        c_doc = next((c for c in df_k.columns if 'DOCUMENTO' in c), None)
                        c_ent_u = next((c for c in df_k.columns if 'ENTRADASUNID' in c), None)
                        c_ent_v = next((c for c in df_k.columns if 'ENTRADASVAL' in c), None)
                        c_costo = next((c for c in df_k.columns if 'COSTOPROMEDIO' in c), None)

                        if not all([c_cod, c_pref, c_ent_u, c_ent_v, c_costo]):
                            st.error("Error: No se encontraron todas las columnas clave en el Kardex.")
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
                            
                            df_k['Prefijo_Upper'] = df_k[c_pref].fillna('').astype(str).str.upper().str.strip()
                            df_k['Doc_Upper'] = df_k[c_doc].fillna('').astype(str).str.upper().str.strip() if c_doc else ""

                            # --- CASCADA DE COSTO BASE (A PRUEBA DE CEROS) ---
                            def get_weighted(df_sub):
                                df_valid = df_sub[df_sub[c_ent_u] > 0]
                                if df_valid.empty: return pd.DataFrame(columns=[c_cod, 'Costo_Ref'])
                                return df_valid.groupby(c_cod).apply(
                                    lambda x: x[c_ent_v].sum() / x[c_ent_u].sum() if x[c_ent_u].sum() > 0 else None
                                ).dropna().reset_index(name='Costo_Ref')

                            ref_cfe = get_weighted(df_k[df_k['Prefijo_Upper'] == 'CFE']).rename(columns={'Costo_Ref': 'C_CFE'})
                            ref_trd = get_weighted(df_k[df_k['Prefijo_Upper'] == 'TRD']).rename(columns={'Costo_Ref': 'C_TRD'})
                            ref_pro = get_weighted(df_k[df_k['Prefijo_Upper'] == 'PRO']).rename(columns={'Costo_Ref': 'C_PRO'})
                            ref_eaj = get_weighted(df_k[df_k['Prefijo_Upper'] == 'EAJ']).rename(columns={'Costo_Ref': 'C_EAJ'})

                            df_ini = df_k[(df_k['Prefijo_Upper'].isin(['INI', 'NAN', ''])) | (df_k['Doc_Upper'].str.contains('SALDO ANTERIOR'))]
                            ref_ini = df_ini[df_ini[c_costo] > 0].groupby(c_cod)[c_costo].first().reset_index(name='C_INI')

                            ref_base = pd.DataFrame({c_cod: df_k[c_cod].unique()})
                            for df_ref, col in zip([ref_cfe, ref_ini, ref_trd, ref_pro, ref_eaj], ['C_CFE', 'C_INI', 'C_TRD', 'C_PRO', 'C_EAJ']):
                                if not df_ref.empty:
                                    ref_base = pd.merge(ref_base, df_ref, on=c_cod, how='left')
                                else:
                                    ref_base[col] = pd.NA
                            
                            ref_base['COSTO_BASE'] = ref_base['C_CFE'].fillna(ref_base['C_INI']).fillna(ref_base['C_TRD']).fillna(ref_base['C_PRO']).fillna(ref_base['C_EAJ']).fillna(0)

                            # --- EVALUACIÓN DETALLADA ---
                            df_ventas = df_k[df_k['Prefijo_Upper'].isin(['FCF', 'CCF'])].copy()
                            df_ventas = df_ventas.merge(ref_base[[c_cod, 'COSTO_BASE']], on=c_cod, how='left')
                            
                            df_ventas = df_ventas[(df_ventas['COSTO_BASE'] > 0) | (df_ventas[c_costo] > 0)]
                            df_ventas['COSTO_BASE_SAFE'] = df_ventas['COSTO_BASE'].replace(0, 1) 
                            df_ventas['VARIACION_%'] = (df_ventas[c_costo] / df_ventas['COSTO_BASE_SAFE']) - 1
                            df_ventas['DIFERENCIA_$'] = df_ventas[c_costo] - df_ventas['COSTO_BASE']

                            condicion_porcentaje = df_ventas['VARIACION_%'].abs() > 0.01
                            condicion_moneda = df_ventas['DIFERENCIA_$'].abs() >= 0.019

                            anomalias = df_ventas[condicion_porcentaje & condicion_moneda].copy()

                            if not anomalias.empty:
                                st.error(f"🚨 ALERTA: Se detectaron {len(anomalias)} movimientos de venta con desviación > 1% y diferencia >= $0.02.")
                                anomalias_show = anomalias[[c_cod, 'Prefijo_Upper', c_doc, 'COSTO_BASE', c_costo, 'DIFERENCIA_$', 'VARIACION_%']].copy()
                                anomalias_show.columns = ['Código', 'Tipo Doc', 'N° Documento', 'Costo Base', 'Costo Venta', 'Diferencia ($)', 'Variación']
                                st.dataframe(anomalias_show.style.format({
                                    'Costo Base': '${:.4f}', 'Costo Venta': '${:.4f}',
                                    'Diferencia ($)': '${:.4f}', 'Variación': '{:.2%}'
                                }), use_container_width=True)
                            else:
                                st.success("✅ Validación exitosa. No hay variaciones significativas (mayores al 1% y a $0.02).")

                    except Exception as e_k:
                        st.error(f"Error procesando el Kardex: {e_k}")

                # ====================================================
                # PROCESAMIENTO 1: ARCHIVO DE INVENTARIOS (FILTRO BLINDADO)
                # ====================================================
                try:
                    partidas_dict = {
                        "TRASLADOS": {"Traslados_Consolidados": []},
                        "AJUSTES": {"Ajustes_Globales": []},
                        "CONSIGNACION": {"Movimientos_Consig": []},
                        "VENTAS_CONSIGNACION": {"Ventas_Consig": []}
                    }
                    
                    nombres_meses = {1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril", 5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto", 9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"}

                    if arch_mov_inv:
                        df = pd.read_excel(arch_mov_inv, dtype=str)
                        df.columns = df.columns.str.strip()
                        
                        # Buscar columna explícita o forzar Columna F (índice 5)
                        c_fecha = next((c for c in df.columns if 'FECHA' in str(c).upper()), None)
                        if not c_fecha and len(df.columns) >= 6:
                            c_fecha = df.columns[5]
                            
                        if c_fecha:
                            # Parseo dual (fechas normales y números de serie de Excel)
                            df['Fecha_DT'] = pd.to_datetime(df[c_fecha], errors='coerce', dayfirst=True)
                            mask_nat = df['Fecha_DT'].isna()
                            if mask_nat.any():
                                safe_nums = pd.to_numeric(df.loc[mask_nat, c_fecha], errors='coerce')
                                df.loc[mask_nat, 'Fecha_DT'] = pd.to_datetime(safe_nums, origin='1899-12-30', unit='D', errors='coerce')

                            mask_fecha = (df['Fecha_DT'].dt.month == mes_proceso) & (df['Fecha_DT'].dt.year == anio_proceso)
                            
                            total_antes = len(df)
                            df = df[mask_fecha]
                            total_despues = len(df)
                            
                            st.info(f"📅 **Auditoría de Fechas (Movimientos):** Se conservaron {total_despues} registros del mes de {nombres_meses.get(mes_proceso, mes_proceso)} (Total original en Excel: {total_antes}).")
                        else:
                            st.warning("No se logró identificar la columna F para filtrar las fechas.")

                        if not df.empty:
                            c_tipo = next((c for c in df.columns if 'TIPO' in c.upper() or 'CONCEPT' in c.upper()), None)
                            
                            # EXTRACCIÓN BLINDADA DE COLUMNAS (Para no confundir bodegas con cantidades de salida)
                            c_sender = next((c for c in df.columns if 'BODEGASALIDA' in c.upper().replace(' ', '')), None)
                            if not c_sender: c_sender = next((c for c in df.columns if 'SALIDA' in c.upper().replace(' ', '') and 'BODEGA' in c.upper()), None)
                            
                            c_receiver = next((c for c in df.columns if 'BODEGAINGRESO' in c.upper().replace(' ', '')), None)
                            if not c_receiver: c_receiver = next((c for c in df.columns if 'INGRESO' in c.upper().replace(' ', '') and 'BODEGA' in c.upper()), None)

                            c_category = next((c for c in df.columns if 'CATEGOR' in c.upper()), 'Categoria')
                            
                            # Priorizar PrecioTotal frente a Costos intermedios
                            c_total = next((c for c in df.columns if 'PRECIOTOTAL' in c.upper().replace(' ', '')), None)
                            if not c_total: c_total = next((c for c in df.columns if 'COSTO' in c.upper()), 'PrecioTotal')
                            
                            df[c_total] = pd.to_numeric(df[c_total], errors='coerce').fillna(0)

                            for _, row in df.iterrows():
                                raw_sender = clean_value(row[c_sender]) if c_sender else ""
                                raw_receiver = clean_value(row[c_receiver]) if c_receiver else ""
                                tipo = clean_value(row[c_tipo]) if c_tipo else ""
                                category = clean_value(row[c_category]) if c_category else ""
                                total = float(row[c_total])
                                
                                if total == 0: continue

                                is_sender_unidad = raw_sender in UNIDADES_SOHO_SET
                                is_receiver_unidad = raw_receiver in UNIDADES_SOHO_SET
                                
                                if not (is_sender_unidad or is_receiver_unidad) or (is_sender_unidad and is_receiver_unidad) or 'SERVICIO' in category: continue

                                inv_acc = INV_ACCOUNT_MAP.get(category, INV_ACCOUNT_MAP['DEFAULT'])
                                sender_desc = raw_sender if raw_sender else "ORIGEN"
                                receiver_desc = raw_receiver if raw_receiver else "DESTINO"
                                
                                es_traslado = 'TRASLAD' in tipo or (raw_sender and raw_receiver)
                                es_ajuste = 'AJUS' in tipo or not es_traslado

                                # BLOQUE CONSIGNACIÓN
                                if 'CONSIGNACION' in category:
                                    if is_receiver_unidad and not is_sender_unidad:
                                        if raw_sender in OTRAS_UNIDADES_INTERNAS or es_traslado:
                                            pass 
                                        else:
                                            desc = f"RECONOCIMIENTO POR ENTRADA DE PRODUCTOS EN CONSIGNACION DE CENTRO SOHO, MES {mes_proceso} DE {anio_proceso}."
                                            partidas_dict["CONSIGNACION"]["Movimientos_Consig"].extend([
                                                gen_nexus_spec(CTA_CONSIGNACION_DR_ENTRADA, desc, total, 0, raw_receiver),
                                                gen_nexus_spec(CTA_CONSIGNACION_CR_ENTRADA, desc, 0, total, raw_receiver)
                                            ])
                                    elif is_sender_unidad:
                                        if es_traslado and receiver_desc != "DESTINO":
                                            desc = f"RECONOCIMIENTO POR {tipo} DE PRODUCTOS EN CONSIGNACION HACIA {receiver_desc}, MES {mes_proceso} DE {anio_proceso}."
                                        else:
                                            desc = f"RECONOCIMIENTO POR {tipo} DE PRODUCTOS EN CONSIGNACION DE CENTRO SOHO, MES {mes_proceso} DE {anio_proceso}."
                                        
                                        partidas_dict["CONSIGNACION"]["Movimientos_Consig"].extend([
                                            gen_nexus_spec(CTA_CONSIGNACION_DR_ENTRADA, desc, total, 0, raw_sender),
                                            gen_nexus_spec(CTA_CONSIGNACION_CR_ENTRADA, desc, 0, total, raw_sender),
                                            gen_nexus_spec(CTA_CONSIGNACION_DR_SALIDA, desc + " ", total, 0, raw_sender),
                                            gen_nexus_spec(CTA_CONSIGNACION_CR_SALIDA, desc + " ", 0, total, raw_sender)
                                        ])

                                        if 'AJUS' in tipo or not es_traslado:
                                            desc_ajuste_costo = f"RECONOCIMIENTO DE SALIDA POR AJUSTE DE PRODUCTO DE CENTRO SOHO, MES {mes_proceso} DE {anio_proceso}."
                                            partidas_dict["AJUSTES"]["Ajustes_Globales"].extend([
                                                gen_nexus_spec(CTA_BASE_GASTO, desc_ajuste_costo, total, 0), 
                                                gen_nexus_spec(inv_acc, desc_ajuste_costo, 0, total)
                                            ])

                                # BLOQUE AJUSTES REGULARES
                                elif es_ajuste:
                                    if is_receiver_unidad:
                                        desc = f"RECONOCIMIENTO DE ENTRADA POR AJUSTE DE PRODUCTO DE CENTRO SOHO, MES {mes_proceso} DE {anio_proceso}."
                                        partidas_dict["AJUSTES"]["Ajustes_Globales"].extend([gen_nexus_spec(inv_acc, desc, total, 0), gen_nexus_spec(CTA_BASE_GASTO, desc, 0, total)])
                                    else:
                                        desc = f"RECONOCIMIENTO DE SALIDA POR AJUSTE DE PRODUCTO DE CENTRO SOHO, MES {mes_proceso} DE {anio_proceso}."
                                        partidas_dict["AJUSTES"]["Ajustes_Globales"].extend([gen_nexus_spec(CTA_BASE_GASTO, desc, total, 0), gen_nexus_spec(inv_acc, desc, 0, total)])

                                # BLOQUE TRASLADOS ESTÁNDAR
                                elif is_sender_unidad and es_traslado:
                                    d_s = f"RECONOCIMIENTO DE SALIDA POR TRASLADO DE PRODUCTO DE {sender_desc} HACIA {receiver_desc}, MES {mes_proceso} DE {anio_proceso}."
                                    d_r = d_s.replace("SALIDA", "ENTRADA")
                                    partidas_dict["TRASLADOS"]["Traslados_Consolidados"].extend([
                                        gen_nexus_spec(CTA_PROVISION_PUENTE, d_s, total, 0), gen_nexus_spec(inv_acc, d_s, 0, total),
                                        gen_nexus_spec(INV_ACCOUNT_MAP.get('DEFAULT'), d_r, total, 0), gen_nexus_spec(CTA_PROVISION_PUENTE, d_r, 0, total)
                                    ])
                    if arch_ventas:
                        df_v = pd.read_excel(arch_ventas, dtype=str)
                        df_v.columns = df_v.columns.str.strip()
                        
                        c_fecha_v = next((c for c in df_v.columns if 'FECHA' in str(c).upper()), None)
                        if not c_fecha_v and len(df_v.columns) >= 6:
                            c_fecha_v = df_v.columns[5]
                            
                        if c_fecha_v:
                            df_v['Fecha_DT'] = pd.to_datetime(df_v[c_fecha_v], errors='coerce', dayfirst=True)
                            mask_nat_v = df_v['Fecha_DT'].isna()
                            if mask_nat_v.any():
                                safe_nums_v = pd.to_numeric(df_v.loc[mask_nat_v, c_fecha_v], errors='coerce')
                                df_v.loc[mask_nat_v, 'Fecha_DT'] = pd.to_datetime(safe_nums_v, origin='1899-12-30', unit='D', errors='coerce')

                            mask_fecha_v = (df_v['Fecha_DT'].dt.month == mes_proceso) & (df_v['Fecha_DT'].dt.year == anio_proceso)
                            
                            total_antes_v = len(df_v)
                            df_v = df_v[mask_fecha_v]
                            total_despues_v = len(df_v)
                            
                            st.info(f"📅 **Auditoría de Fechas (Ventas):** Se conservaron {total_despues_v} registros del mes de {nombres_meses.get(mes_proceso, mes_proceso)} (Total original en Excel: {total_antes_v}).")

                        if not df_v.empty:
                            c_cat_v = next((c for c in df_v.columns if 'CATEGOR' in c.upper()), None)
                            c_costo_v = next((c for c in df_v.columns if 'TOTALCOSTO' in c.upper().replace(' ', '') or 'COSTO' in c.upper()), None)
                            
                            if c_cat_v and c_costo_v:
                                df_v[c_costo_v] = pd.to_numeric(df_v[c_costo_v], errors='coerce').fillna(0)
                                
                                df_consignacion_ventas = df_v[df_v[c_cat_v].astype(str).str.upper().str.strip() == 'CONSIGNACION']
                                total_ventas_consignacion = df_consignacion_ventas[c_costo_v].sum()
                                
                                if total_ventas_consignacion > 0:
                                    desc_venta_consig = f"RECONOCIMIENTO POR VENTA DE PRODUCTOS EN CONSIGNACION DE CENTRO SOHO, MES {mes_proceso} DE {anio_proceso}."
                                    desc_costo_consig = f"RECONOCIMIENTO DE COSTO POR VENTA DE PRODUCTOS EN CONSIGNACION DE CENTRO SOHO, MES {mes_proceso} DE {anio_proceso}."
                                    
                                    partidas_dict["VENTAS_CONSIGNACION"]["Ventas_Consignacion"] = [
                                        gen_nexus_spec(CTA_CONSIGNACION_DR_SALIDA, desc_venta_consig, total_ventas_consignacion, 0, "CENTRO SOHO"),
                                        gen_nexus_spec(CTA_CONSIGNACION_CR_SALIDA, desc_venta_consig, 0, total_ventas_consignacion, "CENTRO SOHO"),
                                        gen_nexus_spec(CTA_BASE_GASTO, desc_costo_consig, total_ventas_consignacion, 0, "CENTRO SOHO"),
                                        gen_nexus_spec(CTA_BASE_PT, desc_costo_consig, 0, total_ventas_consignacion, "CENTRO SOHO")
                                    ]

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