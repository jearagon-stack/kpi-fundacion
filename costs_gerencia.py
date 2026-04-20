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

# Bodegas oficiales de Gerencia Comercial
UNIDADES_GERENCIA_SET = {
    "GERENCIA COMERCIAL", "ADMINISTRACION COMERCIAL", "BODEGA GERENCIA"
}

# Unidades para evitar duplicidad
OTRAS_UNIDADES_INTERNAS = {
    "LIBRERÍA CENTRAL", "LIBRERIA CENTRAL", "LIBRERÍA BODEGA", "LIBRERIA BODEGA",
    "TERRAZA", "CID CAMPUS", "DESPENSA", "CAFETERIA", "CENTRO SOHO", "SOHO", "TALLERES"
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
def mostrar_modulo_gerencia():
    st.title("💼 Módulo de Inventarios - Gerencia Comercial")
    st.info("Automatización de Partidas de Inventario y Ventas en Consignación.")

    tab_carga, tab_liquidacion = st.tabs(["📥 1. Carga de Datos", "💰 2. Partidas y Nexus"])

    with tab_carga:
        st.subheader("Configuración del Periodo")
        col_m1, col_m2 = st.columns(2)
        with col_m1: mes_proceso = st.selectbox("Mes:", range(1, 13), index=date.today().month - 1, key="m_ger")
        with col_m2: anio_proceso = st.number_input("Año:", min_value=2024, value=date.today().year, key="y_ger")

        st.markdown("---")
        st.subheader("Archivos Base")
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            arch_mov_inv = st.file_uploader("1. Reporte de Movimientos", type=["xls", "xlsx"], key="file_inv_ger")
        with col_f2:
            arch_ventas = st.file_uploader("2. Reporte de Ventas", type=["xls", "xlsx"], key="file_ven_ger")
            
        btn_scan = st.button("🔍 Procesar Auditoría", type="primary", use_container_width=True, key="btn_ger")

        if btn_scan:
            if not arch_mov_inv and not arch_ventas:
                st.warning("⚠️ Sube al menos un archivo para procesar.")
                return
                
            with st.spinner("Analizando información..."):
                try:
                    partidas_dict = {
                        "TRASLADOS": {},
                        "AJUSTES": {"Entradas_por_Ajuste": [], "Salidas_por_Ajuste": []},
                        "CONSIGNACION": {},
                        "VENTAS_CONSIGNACION": {}
                    }

                    # 1. INVENTARIOS
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

                            is_sender_unidad = raw_sender in UNIDADES_GERENCIA_SET
                            is_receiver_unidad = raw_receiver in UNIDADES_GERENCIA_SET
                            
                            if not (is_sender_unidad or is_receiver_unidad) or (is_sender_unidad and is_receiver_unidad) or 'SERVICIO' in category: continue

                            inv_acc = INV_ACCOUNT_MAP.get(category, INV_ACCOUNT_MAP['DEFAULT'])
                            es_traslado = 'TRASLAD' in tipo or (raw_sender and raw_receiver)
                            es_ajuste = 'AJUS' in tipo or not es_traslado

                            # PRIORIDAD CONSIGNACIÓN
                            if 'CONSIGNACION' in category:
                                if is_receiver_unidad and not is_sender_unidad:
                                    if raw_sender not in OTRAS_UNIDADES_INTERNAS and not es_traslado:
                                        desc = f"RECONOCIMIENTO POR ENTRADA DE PRODUCTOS EN CONSIGNACION DE GERENCIA COMERCIAL, MES {mes_proceso} DE {anio_proceso}."
                                        if "Entradas_Consignacion" not in partidas_dict["CONSIGNACION"]: partidas_dict["CONSIGNACION"]["Entradas_Consignacion"] = []
                                        partidas_dict["CONSIGNACION"]["Entradas_Consignacion"].extend([gen_nexus_spec(CTA_CONSIGNACION_DR_ENTRADA, desc, total, 0, raw_receiver), gen_nexus_spec(CTA_CONSIGNACION_CR_ENTRADA, desc, 0, total, raw_receiver)])
                                elif is_sender_unidad:
                                    tab_name = raw_receiver if es_traslado and raw_receiver else "Salidas_Consignacion"
                                    desc = f"RECONOCIMIENTO POR {tipo} DE PRODUCTOS EN CONSIGNACION DE GERENCIA COMERCIAL, MES {mes_proceso} DE {anio_proceso}."
                                    if tab_name not in partidas_dict["CONSIGNACION"]: partidas_dict["CONSIGNACION"][tab_name] = []
                                    partidas_dict["CONSIGNACION"][tab_name].extend([
                                        gen_nexus_spec(CTA_CONSIGNACION_DR_ENTRADA, desc, total, 0, raw_sender), gen_nexus_spec(CTA_CONSIGNACION_CR_ENTRADA, desc, 0, total, raw_sender),
                                        gen_nexus_spec(CTA_CONSIGNACION_DR_SALIDA, desc + " ", total, 0, raw_sender), gen_nexus_spec(CTA_CONSIGNACION_CR_SALIDA, desc + " ", 0, total, raw_sender)
                                    ])
                                    if es_ajuste:
                                        desc_ac = f"RECONOCIMIENTO DE SALIDA POR AJUSTE DE PRODUCTO DE GERENCIA COMERCIAL, MES {mes_proceso} DE {anio_proceso}."
                                        partidas_dict["AJUSTES"]["Salidas_por_Ajuste"].extend([gen_nexus_spec(CTA_BASE_GASTO, desc_ac, total, 0), gen_nexus_spec(inv_acc, desc_ac, 0, total)])

                            # AJUSTES REGULARES
                            elif es_ajuste:
                                if is_receiver_unidad:
                                    desc = f"RECONOCIMIENTO DE ENTRADA POR AJUSTE DE PRODUCTO DE GERENCIA COMERCIAL, MES {mes_proceso} DE {anio_proceso}."
                                    partidas_dict["AJUSTES"]["Entradas_por_Ajuste"].extend([gen_nexus_spec(inv_acc, desc, total, 0), gen_nexus_spec(CTA_BASE_GASTO, desc, 0, total)])
                                else:
                                    desc = f"RECONOCIMIENTO DE SALIDA POR AJUSTE DE PRODUCTO DE GERENCIA COMERCIAL, MES {mes_proceso} DE {anio_proceso}."
                                    partidas_dict["AJUSTES"]["Salidas_por_Ajuste"].extend([gen_nexus_spec(CTA_BASE_GASTO, desc, total, 0), gen_nexus_spec(inv_acc, desc, 0, total)])

                            # TRASLADOS
                            elif is_sender_unidad and es_traslado:
                                if raw_receiver not in partidas_dict["TRASLADOS"]: partidas_dict["TRASLADOS"][raw_receiver] = []
                                d_s = f"RECONOCIMIENTO DE SALIDA POR TRASLADO DE PRODUCTO HACIA {raw_receiver}, MES {mes_proceso} DE {anio_proceso}."
                                d_r = d_s.replace("SALIDA", "ENTRADA")
                                partidas_dict["TRASLADOS"][raw_receiver].extend([gen_nexus_spec(CTA_PROVISION_PUENTE, d_s, total, 0), gen_nexus_spec(inv_acc, d_s, 0, total), gen_nexus_spec(INV_ACCOUNT_MAP.get('DEFAULT'), d_r, total, 0), gen_nexus_spec(CTA_PROVISION_PUENTE, d_r, 0, total)])

                    # 2. VENTAS CONSIGNACIÓN
                    if arch_ventas:
                        df_v = pd.read_excel(arch_ventas, dtype=str)
                        df_v.columns = df_v.columns.str.strip()
                        c_cat_v = next((c for c in df_v.columns if 'CATEGOR' in c.upper()), None)
                        c_costo_v = next((c for c in df_v.columns if 'TOTALCOSTO' in c.upper().replace(' ', '') or 'COSTO' in c.upper()), None)
                        if c_cat_v and c_costo_v:
                            df_v[c_costo_v] = pd.to_numeric(df_v[c_costo_v], errors='coerce').fillna(0)
                            total_vc = df_v[df_v[c_cat_v].astype(str).str.upper().str.strip() == 'CONSIGNACION'][c_costo_v].sum()
                            if total_vc > 0:
                                desc_v = f"RECONOCIMIENTO POR VENTA DE PRODUCTOS EN CONSIGNACION DE GERENCIA COMERCIAL, MES {mes_proceso} DE {anio_proceso}."
                                desc_c = f"RECONOCIMIENTO DE COSTO POR VENTA DE PRODUCTOS EN CONSIGNACION DE GERENCIA COMERCIAL, MES {mes_proceso} DE {anio_proceso}."
                                partidas_dict["VENTAS_CONSIGNACION"]["Ventas_Consignacion"] = [gen_nexus_spec(CTA_CONSIGNACION_DR_SALIDA, desc_v, total_vc, 0, "GERENCIA COMERCIAL"), gen_nexus_spec(CTA_CONSIGNACION_CR_SALIDA, desc_v, 0, total_vc, "GERENCIA COMERCIAL"), gen_nexus_spec(CTA_BASE_GASTO, desc_c, total_vc, 0, "GERENCIA COMERCIAL"), gen_nexus_spec(CTA_BASE_PT, desc_c, 0, total_vc, "GERENCIA COMERCIAL")]

                    st.success("✅ Procesamiento finalizado.")
                    st.session_state['nexus_gerencia_buffers'] = gen_excels_memory(partidas_dict, mes_proceso)
                    st.session_state['tab_results_ger'] = True
                except Exception as e: st.error(f"Error: {e}")

    with tab_liquidacion:
        if st.session_state.get('tab_results_ger', False):
            for f, b in st.session_state['nexus_gerencia_buffers'].items():
                st.download_button(label=f"⬇️ Descargar {f}", data=b, file_name=f, key=f"dl_ger_{f}")