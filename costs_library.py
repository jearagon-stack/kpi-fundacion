import streamlit as st
import pandas as pd
import io
from datetime import date

# ==========================================
# CONSTANTES CONTABLES - CONOCIMIENTO DE NEGOCIO
# ==========================================
CTA_PROVISION_PUENTE = "21020302" 
CTA_CONSIGNACION_DR_ENTRADA = "7101" 
CTA_CONSIGNACION_CR_ENTRADA = "8101" 
CTA_CONSIGNACION_DR_SALIDA = "8101" 
CTA_CONSIGNACION_CR_SALIDA = "7101" 

CTA_BASE_GASTO = "410104" 
CTA_BASE_PT = "110603" 
CTA_BASE_MP = "110601" 

UNIDADES_LIBRERIA_EXPENSE_MAP = {
    "LIBRERÍA CENTRAL": CTA_BASE_GASTO,
    "LIBRERÍA CAMPUS": CTA_BASE_GASTO,
    "LIBRERÍA CENTRAL (B001)": CTA_BASE_GASTO, 
    "LIBRERÍA BODEGA": CTA_BASE_GASTO,
    "LIBRERÍA BODEGA TG": CTA_BASE_GASTO 
}

INV_ACCOUNT_MAP = {
    'MATERIA PRIMA': CTA_BASE_MP,
    'PRODUCTO TERMINADO': CTA_BASE_PT,
    'DEFAULT': CTA_BASE_PT
}

UNIDADES_LIBRERIA_SET = set(UNIDADES_LIBRERIA_EXPENSE_MAP.keys())

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
        "AJUSTES": "Ajustes_Internos",
        "CONSIGNACION": "Movimientos_Consignacion"
    }
    
    for categoria, tabs_dict in partidas_dict.items():
        if not any(len(lst) > 0 for lst in tabs_dict.values()): continue
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, partidas_list in tabs_dict.items():
                if not partidas_list: continue
                
                df_nexus = pd.DataFrame(partidas_list)
                
                # Agrupar sumando DEBE y HABER
                df_grouped = df_nexus.groupby(["CUENTA", "VACIO", "CONCEPTO", "AUXILIAR"], as_index=False)[["DEBE", "HABER"]].sum()
                df_grouped = df_grouped[(df_grouped["DEBE"] > 0) | (df_grouped["HABER"] > 0)]
                df_grouped = df_grouped.sort_values(by=["CONCEPTO", "HABER", "DEBE"], ascending=[True, True, False])
                
                nexus_export = df_grouped[["CUENTA", "VACIO", "CONCEPTO", "DEBE", "HABER"]].copy() 
                
                safe_sheet = str(sheet_name).replace("/", "-").replace("\\", "-").replace("?", "").replace("*", "").replace("[", "").replace("]", "")
                safe_sheet = safe_sheet[:31] if safe_sheet else "Partidas"
                
                nexus_export.to_excel(writer, index=False, header=False, sheet_name=safe_sheet)
                
        nombre_base = nombres_amigables.get(categoria, categoria)
        nombre_archivo = f"Partidas_{nombre_base}_Mes_{mes_proceso}.xlsx"
        buffers[nombre_archivo] = output.getvalue()
        
    return buffers

def mostrar_modulo_libreria():
    st.title("📚 Módulo de Inventarios - Librería")
    st.info("Tratamiento Contable Automático de Traslados y Ajustes de Librería.")

    tab_carga, tab_liquidacion = st.tabs(["📥 1. Carga de Datos", "💰 2. Partidas y Nexus"])

    with tab_carga:
        st.subheader("Paso 1: Periodo y Archivo Base")
        col_m1, col_m2 = st.columns(2)
        with col_m1: mes_proceso = st.selectbox("Mes:", range(1, 13), index=date.today().month - 1, key="m_lib")
        with col_m2: anio_proceso = st.number_input("Año:", min_value=2024, max_value=2030, value=date.today().year, key="y_lib")

        st.markdown("---")
        arch_mov_inv = st.file_uploader("1. Reporte de Movimientos de Inventario (xls/xlsx)", type=["xls", "xlsx"])
        
        btn_scan = st.button("🔍 Escanear y Procesar Librería", type="primary", use_container_width=True)

        if arch_mov_inv and btn_scan:
            current_expense_map = UNIDADES_LIBRERIA_EXPENSE_MAP

            with st.spinner("Procesando Auditoría Contable de Librería..."):
                try:
                    df = pd.read_excel(arch_mov_inv, dtype=str)
                    df.columns = df.columns.str.strip()
                    
                    c_tipo = next((c for c in df.columns if 'TIPO' in c.upper() or 'CONCEPT' in c.upper()), None)
                    c_sender = next((c for c in df.columns if 'SALIDA' in c.upper().replace(' ', '')), None)
                    c_receiver = next((c for c in df.columns if 'INGRESO' in c.upper().replace(' ', '')), None)
                    c_category = next((c for c in df.columns if 'CATEGOR' in c.upper()), 'Categoria')
                    c_total = next((c for c in df.columns if 'PRECIOTOTAL' in c.upper().replace(' ', '') or 'COSTO' in c.upper()), 'PrecioTotal')
                    
                    if not all([c_tipo, c_sender, c_receiver]):
                        st.error("❌ El archivo no tiene el formato esperado (Faltan columnas de Tipo, BodegaSalida o BodegaIngreso)")
                        st.stop()
                    
                    if c_total in df.columns: df[c_total] = pd.to_numeric(df[c_total], errors='coerce').fillna(0)

                    partidas_dict = {
                        "TRASLADOS": {},
                        "AJUSTES": {"Ajustes_Internos": []},
                        "CONSIGNACION": {"Consignacion": []}
                    }
                    
                    stats = {"consignacion": 0, "servicios": 0, "standard": 0, "adjustments": 0, "internal_ignored": 0}

                    for i, row in df.iterrows():
                        raw_sender = clean_value(row[c_sender])
                        raw_receiver = clean_value(row[c_receiver])
                        tipo = clean_value(row[c_tipo])
                        category = clean_value(row[c_category])
                        total = float(row[c_total])

                        if total == 0: continue

                        is_sender_lib = raw_sender in UNIDADES_LIBRERIA_SET
                        is_receiver_lib = raw_receiver in UNIDADES_LIBRERIA_SET

                        # FILTRO 0: Si ninguna de las dos bodegas es librería, ignorar completamente
                        if not is_sender_lib and not is_receiver_lib:
                            continue

                        # FILTRO 1: Movimientos internos (Librería a Librería) -> SE IGNORAN
                        if is_sender_lib and is_receiver_lib:
                            stats["internal_ignored"] += 1
                            continue

                        # FILTRO 2: Categoría Servicio -> SE IGNORA
                        if 'SERVICIO' in category:
                            stats["servicios"] += 1
                            continue 

                        # Diagnóstico CORRECTO de la naturaleza del movimiento
                        es_traslado = 'TRASLAD' in tipo or (raw_sender != "" and raw_receiver != "")
                        es_ajuste = 'AJUS' in tipo
                        
                        # Regla de oro de traslados: Si entra a la librería desde otra unidad, se ignora
                        if es_traslado and is_receiver_lib and not is_sender_lib:
                            continue

                        sender_desc = raw_sender if raw_sender else "ORIGEN_NO_ESPECIFICADO"
                        receiver_desc = raw_receiver if raw_receiver else "DESTINO_NO_ESPECIFICADO"
                        inv_acc = INV_ACCOUNT_MAP.get(category, INV_ACCOUNT_MAP['DEFAULT'])

                        # ==========================
                        # A. AJUSTES (y entradas/salidas huérfanas que no son consignación)
                        # ==========================
                        if es_ajuste or (not es_traslado and 'CONSIGNACION' not in category):
                            stats["adjustments"] += 1
                            if is_receiver_lib and not is_sender_lib:
                                # Entrada por Ajuste (BodegaSalida viene vacía)
                                unit_nexus_expense = current_expense_map.get(raw_receiver, CTA_BASE_GASTO)
                                desc_adjust_entry = f"RECONOCIMIENTO DE ENTRADA POR AJUSTE DE PRODUCTO DE {receiver_desc}, MES {mes_proceso} DE {anio_proceso}." 
                                partidas_dict["AJUSTES"]["Ajustes_Internos"].extend([
                                    gen_nexus_spec(inv_acc, desc_adjust_entry, total, 0),
                                    gen_nexus_spec(unit_nexus_expense, desc_adjust_entry, 0, total)
                                ])
                            elif is_sender_lib and not is_receiver_lib:
                                # Salida por Ajuste (BodegaIngreso viene vacía)
                                unit_nexus_expense = current_expense_map.get(raw_sender, CTA_BASE_GASTO)
                                desc_adjust_exit = f"RECONOCIMIENTO DE SALIDA POR AJUSTE DE PRODUCTO DE {sender_desc}, MES {mes_proceso} DE {anio_proceso}."
                                partidas_dict["AJUSTES"]["Ajustes_Internos"].extend([
                                    gen_nexus_spec(unit_nexus_expense, desc_adjust_exit, total, 0),
                                    gen_nexus_spec(inv_acc, desc_adjust_exit, 0, total)
                                ])

                        # ==========================
                        # B. CONSIGNACION
                        # ==========================
                        elif 'CONSIGNACION' in category:
                            stats["consignacion"] += 1
                            if is_receiver_lib and not is_sender_lib:
                                # Entrada de Consignación
                                desc_cons_ent = f"RECONOCIMIENTO POR ENTRADA DE PRODUCTOS EN CONSIGNACION DE {receiver_desc}, MES {mes_proceso} DE {anio_proceso}."
                                partidas_dict["CONSIGNACION"]["Consignacion"].extend([
                                    gen_nexus_spec(CTA_CONSIGNACION_DR_ENTRADA, desc_cons_ent, total, 0, raw_receiver),
                                    gen_nexus_spec(CTA_CONSIGNACION_CR_ENTRADA, desc_cons_ent, 0, total, raw_receiver)
                                ])
                            elif is_sender_lib:
                                # Salida / Devolución de Consignación
                                desc_cons_dev = f"RECONOCIMIENTO POR DEVOLUCION DE PRODUCTOS EN CONSIGNACION DE {sender_desc}, MES {mes_proceso} DE {anio_proceso}."
                                partidas_dict["CONSIGNACION"]["Consignacion"].extend([
                                    gen_nexus_spec(CTA_CONSIGNACION_DR_SALIDA, desc_cons_dev, total, 0, raw_sender),
                                    gen_nexus_spec(CTA_CONSIGNACION_CR_SALIDA, desc_cons_dev, 0, total, raw_sender)
                                ])

                        # ==========================
                        # C. TRASLADOS (Las 4 líneas en la misma pestaña)
                        # ==========================
                        elif es_traslado and is_sender_lib:
                            stats["standard"] += 1
                            desc_stand_send = f"RECONOCIMIENTO DE SALIDA POR TRASLADO DE PRODUCTO DE {sender_desc} HACIA {receiver_desc}, MES {mes_proceso} DE {anio_proceso}." 
                            desc_stand_receive = f"RECONOCIMIENTO DE ENTRADA POR TRASLADO DE PRODUCTO DE {sender_desc} HACIA {receiver_desc}, MES {mes_proceso} DE {anio_proceso}." 

                            if receiver_desc not in partidas_dict["TRASLADOS"]:
                                partidas_dict["TRASLADOS"][receiver_desc] = []

                            # Partida 1: Salida de Librería
                            partidas_dict["TRASLADOS"][receiver_desc].extend([
                                gen_nexus_spec(CTA_PROVISION_PUENTE, desc_stand_send, total, 0),
                                gen_nexus_spec(inv_acc, desc_stand_send, 0, total)
                            ])

                            # Partida 2: Entrada al Destino
                            partidas_dict["TRASLADOS"][receiver_desc].extend([
                                gen_nexus_spec(INV_ACCOUNT_MAP.get('DEFAULT'), desc_stand_receive, total, 0),
                                gen_nexus_spec(CTA_PROVISION_PUENTE, desc_stand_receive, 0, total)
                            ])

                    st.success(f"✅ Auditoría completada: {stats['standard']} traslados, {stats['consignacion']} consignaciones, {stats['adjustments']} ajustes procesados.")
                    st.session_state['nexus_libreria_buffers'] = gen_excels_memory(partidas_dict, mes_proceso)
                    st.session_state['results_libreria'] = df
                    st.session_state['tab_results_lib'] = True

                except Exception as e:
                    st.error(f"Error al procesar: {e}")

    with tab_liquidacion:
        st.subheader("Paso 2: Descarga de Partidas Acumuladas (Formato Nexus)")
        if st.session_state.get('tab_results_lib', False):
            buffers = st.session_state['nexus_libreria_buffers']
            for filename, buffer in buffers.items():
                st.download_button(
                    label=f"⬇️ Descargar {filename}",
                    data=buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{filename}"
                )