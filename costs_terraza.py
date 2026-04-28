import streamlit as st
import pandas as pd
import io
from datetime import date

# ==========================================
# CONSTANTES CONTABLES
# ==========================================
CTA_PROVISION_PUENTE = "21020302" 
CTA_BASE_GASTO = "410104" 
CTA_BASE_PT = "110603" 
CTA_BASE_MP = "110601" 

# Bodegas oficiales de la unidad
UNIDADES_TERRAZA_SET = {
    "TERRAZA", "BODEGA TERRAZA"
}

# Otras unidades para cruces
OTRAS_UNIDADES_INTERNAS = {
    "CENTRO SOHO", "BODEGA SOHO", "SOHO",
    "GERENCIA COMERCIAL", "BODEGA GERENCIA COMERCIAL",
    "CID CAMPUS", "BODEGA CID CAMPUS", 
    "DESPENSA", "CAFETERIA", "CAFETERIA CENTRAL",
    "TALLERES", "LIBRERÍA CENTRAL", "LIBRERÍA BODEGA", 
    "LIBRERÍA BODEGA TG", "LIBRERÍA CAMPUS", "LIBRERÍA CENTRAL (B001)"
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

def format_id(val):
    v = str(val).strip()
    return str(int(float(v))) if v.replace('.', '', 1).isdigit() else v

def gen_excels_memory(partidas_dict, mes_proceso):
    buffers = {}
    nombres_amigables = {
        "TRASLADOS": "Traslados_y_Recepciones",
        "AJUSTES": "Ajustes_de_Inventario"
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
def mostrar_modulo_terraza():
    st.title("☕ Módulo de Inventarios - Terraza")
    st.info("Automatización de Partidas Contables de Inventario.")

    tab_carga, tab_liquidacion = st.tabs(["📥 1. Carga de Datos", "💰 2. Partidas y Nexus"])

    with tab_carga:
        st.subheader("Configuración del Periodo")
        col_m1, col_m2 = st.columns(2)
        with col_m1: mes_proceso = st.selectbox("Mes:", range(1, 13), index=date.today().month - 1, key="m_ter")
        with col_m2: anio_proceso = st.number_input("Año:", min_value=2024, value=date.today().year, key="y_ter")

        st.markdown("---")
        st.subheader("Archivos Base (Generación Nexus y Auditoría)")
        
        arch_mov_inv = st.file_uploader("1. Reporte de Movimientos (Inventario)", type=["xls", "xlsx"], key="file_inv_ter")
        
        col_a1, col_a2 = st.columns(2)
        with col_a1:
            arch_kardex = st.file_uploader("2. Kardex Consolidado (Auditoría)", type=["xls", "xlsx"], key="file_kar_ter")
        with col_a2:
            arch_cat = st.file_uploader("3. Maestro Categorías (Opcional)", type=["xls", "xlsx"], key="file_cat_ter")

        btn_scan = st.button("🔍 Procesar y Generar Partidas", type="primary", use_container_width=True, key="btn_ter")

        if btn_scan:
            if not arch_mov_inv and not arch_kardex:
                st.warning("⚠️ Debes subir al menos un archivo para procesar.")
                return

            with st.spinner("Analizando información y procesando datos..."):
                
                # ====================================================
                # BLOQUE: AUDITORÍA DE KARDEX
                # ====================================================
                if arch_kardex:
                    try:
                        st.markdown("---")
                        st.subheader("🕵️ Resultado de Auditoría Kardex (Detallada)")
                        
                        df_k = pd.read_excel(arch_kardex, dtype=str)
                        df_k.columns = df_k.columns.astype(str).str.strip().str.upper()

                        c_cod = next((c for c in df_k.columns if 'IDPRODUCTO' in c), None)
                        c_pref_k = next((c for c in df_k.columns if 'PREFI' in c), None)
                        c_doc = next((c for c in df_k.columns if 'DOCUMENTO' in c), None)
                        c_ent_u = next((c for c in df_k.columns if 'ENTRADASUNID' in c), None)
                        c_ent_v = next((c for c in df_k.columns if 'ENTRADASVAL' in c), None)
                        c_costo = next((c for c in df_k.columns if 'COSTOPROMEDIO' in c), None)

                        if not all([c_cod, c_pref_k, c_ent_u, c_ent_v, c_costo]):
                            st.error("Error: Faltan columnas en el Kardex para auditar.")
                        else:
                            df_k[c_cod] = df_k[c_cod].astype(str).str.strip()

                            if arch_cat:
                                df_c = pd.read_excel(arch_cat, dtype=str)
                                df_c.columns = df_c.columns.astype(str).str.strip()
                                mapa_categorias = dict(zip(df_c.iloc[:,0].astype(str).str.strip(), df_c.iloc[:,1].astype(str).str.upper().str.strip()))
                                df_k['Cat_Temp'] = df_k[c_cod].map(mapa_categorias).fillna('DESCONOCIDA')
                                df_k = df_k[df_k['Cat_Temp'] != 'SERVICIO']

                            df_k[c_ent_u] = pd.to_numeric(df_k[c_ent_u], errors='coerce').fillna(0)
                            df_k[c_ent_v] = pd.to_numeric(df_k[c_ent_v], errors='coerce').fillna(0)
                            df_k[c_costo] = pd.to_numeric(df_k[c_costo], errors='coerce').fillna(0)
                            
                            df_k['Prefijo_Upper'] = df_k[c_pref_k].fillna('').astype(str).str.upper().str.strip()
                            df_k['Doc_Upper'] = df_k[c_doc].fillna('').astype(str).str.upper().str.strip() if c_doc else ""

                            def get_weighted(df_sub):
                                df_valid = df_sub[df_sub[c_ent_u] > 0]
                                if df_valid.empty: return pd.DataFrame(columns=[c_cod, 'Costo_Ref'])
                                return df_valid.groupby(c_cod).apply(
                                    lambda x: x[c_ent_v].sum() / x[c_ent_u].sum() if x[c_ent_u].sum() > 0 else None
                                ).dropna().reset_index(name='Costo_Ref')

                            ref_cfe = get_weighted(df_k[df_k['Prefijo_Upper'].isin(['CFE', 'FCOM', 'FSE', 'FAC', 'CCF'])]).rename(columns={'Costo_Ref': 'C_CFE'})
                            ref_trd = get_weighted(df_k[df_k['Prefijo_Upper'].isin(['TRD', 'TIN'])]).rename(columns={'Costo_Ref': 'C_TRD'})
                            ref_pro = get_weighted(df_k[df_k['Prefijo_Upper'] == 'PRO']).rename(columns={'Costo_Ref': 'C_PRO'})
                            ref_eaj = get_weighted(df_k[df_k['Prefijo_Upper'].isin(['EAJ', 'AJU', 'ENT', 'REC'])]).rename(columns={'Costo_Ref': 'C_EAJ'})

                            df_ini = df_k[(df_k['Prefijo_Upper'].isin(['INI', 'NAN', ''])) | (df_k['Doc_Upper'].str.contains('SALDO ANTERIOR'))]
                            ref_ini = df_ini[df_ini[c_costo] > 0].groupby(c_cod)[c_costo].first().reset_index(name='C_INI')

                            ref_base = pd.DataFrame({c_cod: df_k[c_cod].unique()})
                            for df_ref, col in zip([ref_cfe, ref_ini, ref_trd, ref_pro, ref_eaj], ['C_CFE', 'C_INI', 'C_TRD', 'C_PRO', 'C_EAJ']):
                                if not df_ref.empty: ref_base = pd.merge(ref_base, df_ref, on=c_cod, how='left')
                                else: ref_base[col] = pd.NA
                            
                            ref_base['COSTO_BASE'] = ref_base['C_CFE'].fillna(ref_base['C_INI']).fillna(ref_base['C_TRD']).fillna(ref_base['C_PRO']).fillna(ref_base['C_EAJ']).fillna(0)

                            df_ventas = df_k[df_k['Prefijo_Upper'].isin(['FCF', 'CCF'])].copy()
                            df_ventas = df_ventas.merge(ref_base[[c_cod, 'COSTO_BASE']], on=c_cod, how='left')
                            
                            df_ventas['COSTO_BASE'] = pd.to_numeric(df_ventas['COSTO_BASE'], errors='coerce').fillna(0)
                            mask_zero_base = df_ventas['COSTO_BASE'] == 0
                            df_ventas.loc[mask_zero_base, 'COSTO_BASE'] = df_ventas.loc[mask_zero_base, c_costo]

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
                                    'Costo Base': '${:.4f}', 'Costo Venta': '${:.4f}', 'Diferencia ($)': '${:.4f}', 'Variación': '{:.2%}'
                                }), use_container_width=True)
                            else:
                                st.success("✅ Validación Kardex exitosa. No hay variaciones significativas.")
                    except Exception as e_k: st.error(f"Error procesando el Kardex: {e_k}")

                # ====================================================
                # PROCESAMIENTO 1: ARCHIVO DE INVENTARIOS
                # ====================================================
                try:
                    partidas_dict = {
                        "TRASLADOS": {},
                        "AJUSTES": {"Entradas_por_Ajuste": [], "Salidas_por_Ajuste": []}
                    }
                    
                    nombres_meses = {1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril", 5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto", 9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"}

                    def extraer_fechas_seguras(df_temp, col_name):
                        s_dt = pd.to_datetime(df_temp[col_name], errors='coerce')
                        m1 = s_dt.isna()
                        if m1.any(): s_dt.loc[m1] = pd.to_datetime(df_temp.loc[m1, col_name], format='%d/%m/%Y', errors='coerce')
                        m2 = s_dt.isna()
                        if m2.any():
                            nums = pd.to_numeric(df_temp.loc[m2, col_name], errors='coerce')
                            s_dt.loc[m2] = pd.to_datetime(nums, origin='1899-12-30', unit='D', errors='coerce')
                        return s_dt

                    if arch_mov_inv:
                        df = pd.read_excel(arch_mov_inv, dtype=str)
                        df.columns = df.columns.str.strip()
                        
                        c_fecha = next((c for c in df.columns if str(c).upper().strip() == 'FECHA'), None)
                        if not c_fecha and len(df.columns) >= 6: c_fecha = df.columns[5]
                            
                        if c_fecha:
                            df['Fecha_DT'] = extraer_fechas_seguras(df, c_fecha)
                            mask_fecha = (df['Fecha_DT'].dt.month == mes_proceso) & (df['Fecha_DT'].dt.year == anio_proceso)
                            total_antes = len(df)
                            df = df[mask_fecha]
                            total_despues = len(df)
                            st.info(f"📅 **Movimientos Inventario:** Se conservaron {total_despues} registros de {nombres_meses.get(mes_proceso, mes_proceso)} (Total original: {total_antes}).")

                        if not df.empty:
                            c_tipo = next((c for c in df.columns if 'TIPO' in c.upper() or 'CONCEPT' in c.upper()), None)
                            
                            c_sender = next((c for c in df.columns if c.upper().strip() in ['BODEGASALIDA', 'BODEGA SALIDA', 'SALIDA']), None)
                            if not c_sender: c_sender = next((c for c in df.columns if 'SALIDA' in c.upper() and 'VAL' not in c.upper() and 'UNID' not in c.upper()), None)
                            
                            c_receiver = next((c for c in df.columns if c.upper().strip() in ['BODEGAINGRESO', 'BODEGA INGRESO', 'INGRESO', 'ENTRADA']), None)
                            if not c_receiver: c_receiver = next((c for c in df.columns if ('INGRESO' in c.upper() or 'ENTRADA' in c.upper()) and 'VAL' not in c.upper() and 'UNID' not in c.upper()), None)
                            
                            c_bodega = next((c for c in df.columns if c.upper().strip() in ['BODEGA', 'SUCURSAL']), None)
                            c_category = next((c for c in df.columns if 'CATEGOR' in c.upper()), 'Categoria')
                            
                            c_total = next((c for c in df.columns if 'PRECIOTOT' in c.upper().replace(' ', '') or 'COSTO' in c.upper()), None)
                            if not c_total: c_total = next((c for c in df.columns if 'TOTAL' in c.upper()), None)
                            
                            df[c_total] = pd.to_numeric(df[c_total], errors='coerce').fillna(0)

                            for _, row in df.iterrows():
                                tipo = clean_value(row[c_tipo]) if c_tipo else ""
                                total = float(row[c_total])
                                if total == 0: continue

                                raw_sender = clean_value(row[c_sender]) if c_sender else ""
                                raw_receiver = clean_value(row[c_receiver]) if c_receiver else ""
                                bod_principal = clean_value(row[c_bodega]) if c_bodega else ""
                                
                                if raw_sender == "" and ("SALIDA" in tipo or "AJUSTE" in tipo): raw_sender = bod_principal
                                if raw_receiver == "" and ("ENTRADA" in tipo or "INGRESO" in tipo): raw_receiver = bod_principal

                                category = clean_value(row[c_category]) if c_category else ""

                                is_sender_ter = raw_sender in UNIDADES_TERRAZA_SET
                                is_receiver_ter = raw_receiver in UNIDADES_TERRAZA_SET
                                
                                if not (is_sender_ter or is_receiver_ter) or (is_sender_ter and is_receiver_ter) or 'SERVICIO' in category: continue

                                inv_acc = INV_ACCOUNT_MAP.get(category, INV_ACCOUNT_MAP['DEFAULT'])
                                sender_desc = raw_sender if raw_sender else "ORIGEN"
                                receiver_desc = raw_receiver if raw_receiver else "DESTINO"
                                
                                es_traslado = 'TRASLAD' in tipo or (raw_sender and raw_receiver)
                                es_ajuste_o_consumo = not es_traslado

                                if es_ajuste_o_consumo:
                                    if is_receiver_ter and not is_sender_ter:
                                        desc = f"RECONOCIMIENTO DE ENTRADA POR AJUSTE DE PRODUCTO DE TERRAZA, MES {mes_proceso} DE {anio_proceso}."
                                        partidas_dict["AJUSTES"]["Entradas_por_Ajuste"].extend([gen_nexus_spec(inv_acc, desc, total, 0), gen_nexus_spec(CTA_BASE_GASTO, desc, 0, total)])
                                    
                                    elif is_sender_ter and not is_receiver_ter:
                                        desc = f"RECONOCIMIENTO DE SALIDA POR AJUSTE O CONSUMO DE PRODUCTO DE TERRAZA, MES {mes_proceso} DE {anio_proceso}."
                                        partidas_dict["AJUSTES"]["Salidas_por_Ajuste"].extend([gen_nexus_spec(CTA_BASE_GASTO, desc, total, 0), gen_nexus_spec(inv_acc, desc, 0, total)])

                                elif is_sender_ter and es_traslado:
                                    safe_rec_t = str(receiver_desc).replace('/', '-').replace('\\', '-')[:25]
                                    if safe_rec_t not in partidas_dict["TRASLADOS"]: partidas_dict["TRASLADOS"][safe_rec_t] = []
                                    d_s = f"RECONOCIMIENTO DE SALIDA POR TRASLADO DE PRODUCTO DE {sender_desc} HACIA {receiver_desc}, MES {mes_proceso} DE {anio_proceso}."
                                    d_r = d_s.replace("SALIDA", "ENTRADA")
                                    partidas_dict["TRASLADOS"][safe_rec_t].extend([
                                        gen_nexus_spec(CTA_PROVISION_PUENTE, d_s, total, 0), gen_nexus_spec(inv_acc, d_s, 0, total),
                                        gen_nexus_spec(INV_ACCOUNT_MAP.get('DEFAULT'), d_r, total, 0), gen_nexus_spec(CTA_PROVISION_PUENTE, d_r, 0, total)
                                    ])

                    if arch_mov_inv:
                        st.success("Procesamiento contable finalizado.")
                        st.session_state['nexus_terraza_buffers'] = gen_excels_memory(partidas_dict, mes_proceso)
                        st.session_state['tab_results_ter'] = True
                        
                except Exception as e: st.error(f"Error técnico en procesamiento: {e}")

    with tab_liquidacion:
        if st.session_state.get('tab_results_ter', False):
            for f, b in st.session_state['nexus_terraza_buffers'].items():
                st.download_button(label=f"⬇️ Descargar {f}", data=b, file_name=f, key=f"dl_{f}")