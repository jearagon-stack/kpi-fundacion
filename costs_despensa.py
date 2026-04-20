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

# Bodegas oficiales de la unidad actual
UNIDADES_DESPENSA_SET = {
    "DESPENSA"
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
        "AJUSTES": "Ajustes_de_Inventario"
    }
    
    for categoria, tabs_dict in partidas_dict.items():
        if not any(len(lst) > 0 for lst in tabs_dict.values()): continue
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, partidas_list in tabs_dict.items():
                if not partidas_list: continue
                df_nexus = pd.DataFrame(partidas_list)
                
                # Agrupación por cuenta y concepto para Nexus
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
def mostrar_modulo_despensa():
    st.title("🛒 Módulo de Inventarios - Despensa")
    st.info("Automatización de Partidas Contables de Inventario para Despensa.")

    tab_carga, tab_liquidacion = st.tabs(["📥 1. Carga de Datos", "💰 2. Partidas y Nexus"])

    with tab_carga:
        st.subheader("Configuración del Periodo")
        col_m1, col_m2 = st.columns(2)
        with col_m1: mes_proceso = st.selectbox("Mes:", range(1, 13), index=date.today().month - 1, key="m_despensa")
        with col_m2: anio_proceso = st.number_input("Año:", min_value=2024, value=date.today().year, key="y_despensa")

        st.markdown("---")
        arch_mov_inv = st.file_uploader("Subir Reporte de Movimientos (Inventario)", type=["xls", "xlsx"], key="file_inv_despensa")
        btn_scan = st.button("🔍 Procesar Auditoría", type="primary", use_container_width=True, key="btn_despensa")

        if btn_scan:
            if not arch_mov_inv:
                st.warning("⚠️ Debes subir el archivo de movimientos para procesar.")
                return

            with st.spinner("Analizando información..."):
                try:
                    df = pd.read_excel(arch_mov_inv, dtype=str)
                    df.columns = df.columns.str.strip()
                    c_tipo = next((c for c in df.columns if 'TIPO' in c.upper() or 'CONCEPT' in c.upper()), None)
                    c_sender = next((c for c in df.columns if 'SALIDA' in c.upper().replace(' ', '')), None)
                    c_receiver = next((c for c in df.columns if 'INGRESO' in c.upper().replace(' ', '')), None)
                    c_category = next((c for c in df.columns if 'CATEGOR' in c.upper()), 'Categoria')
                    c_total = next((c for c in df.columns if 'PRECIOTOTAL' in c.upper().replace(' ', '') or 'COSTO' in c.upper()), 'PrecioTotal')
                    
                    df[c_total] = pd.to_numeric(df[c_total], errors='coerce').fillna(0)
                    
                    partidas_dict = {
                        "TRASLADOS": {},
                        "AJUSTES": {"Entradas_por_Ajuste": [], "Salidas_por_Ajuste": []}
                    }

                    for _, row in df.iterrows():
                        raw_sender, raw_receiver = clean_value(row[c_sender]), clean_value(row[c_receiver])
                        tipo, category, total = clean_value(row[c_tipo]), clean_value(row[c_category]), float(row[c_total])
                        if total == 0: continue

                        is_sender_unidad = raw_sender in UNIDADES_DESPENSA_SET
                        is_receiver_unidad = raw_receiver in UNIDADES_DESPENSA_SET
                        
                        if not (is_sender_unidad or is_receiver_unidad) or (is_sender_unidad and is_receiver_unidad) or 'SERVICIO' in category: continue

                        inv_acc = INV_ACCOUNT_MAP.get(category, INV_ACCOUNT_MAP['DEFAULT'])
                        sender_desc = raw_sender if raw_sender else "ORIGEN"
                        receiver_desc = raw_receiver if raw_receiver else "DESTINO"
                        
                        es_traslado = 'TRASLAD' in tipo or (raw_sender and raw_receiver)

                        # A. AJUSTES REGULARES
                        if 'AJUS' in tipo or not es_traslado:
                            if is_receiver_unidad: 
                                desc = f"RECONOCIMIENTO DE ENTRADA POR AJUSTE DE PRODUCTO DE DESPENSA, MES {mes_proceso} DE {anio_proceso}."
                                partidas_dict["AJUSTES"]["Entradas_por_Ajuste"].extend([gen_nexus_spec(inv_acc, desc, total, 0), gen_nexus_spec(CTA_BASE_GASTO, desc, 0, total)])
                            else: 
                                desc = f"RECONOCIMIENTO DE SALIDA POR AJUSTE DE PRODUCTO DE DESPENSA, MES {mes_proceso} DE {anio_proceso}."
                                partidas_dict["AJUSTES"]["Salidas_por_Ajuste"].extend([gen_nexus_spec(CTA_BASE_GASTO, desc, total, 0), gen_nexus_spec(inv_acc, desc, 0, total)])

                        # B. TRASLADOS ESTÁNDAR
                        elif is_sender_unidad and es_traslado:
                            if receiver_desc not in partidas_dict["TRASLADOS"]: partidas_dict["TRASLADOS"][receiver_desc] = []
                            d_s = f"RECONOCIMIENTO DE SALIDA POR TRASLADO DE PRODUCTO DE {sender_desc} HACIA {receiver_desc}, MES {mes_proceso} DE {anio_proceso}."
                            d_r = d_s.replace("SALIDA", "ENTRADA")
                            partidas_dict["TRASLADOS"][receiver_desc].extend([
                                gen_nexus_spec(CTA_PROVISION_PUENTE, d_s, total, 0), gen_nexus_spec(inv_acc, d_s, 0, total),
                                gen_nexus_spec(INV_ACCOUNT_MAP.get('DEFAULT'), d_r, total, 0), gen_nexus_spec(CTA_PROVISION_PUENTE, d_r, 0, total)
                            ])

                    st.success("✅ Procesamiento finalizado con éxito.")
                    st.session_state['nexus_despensa_buffers'] = gen_excels_memory(partidas_dict, mes_proceso)
                    st.session_state['tab_results_despensa'] = True
                except Exception as e: st.error(f"Error técnico: {e}")

    with tab_liquidacion:
        if st.session_state.get('tab_results_despensa', False):
            for f, b in st.session_state['nexus_despensa_buffers'].items():
                st.download_button(label=f"⬇️ Descargar {f}", data=b, file_name=f, key=f"dl_despensa_{f}")