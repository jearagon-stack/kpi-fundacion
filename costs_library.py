import streamlit as st
import pandas as pd
import io
from datetime import date

# ==========================================
# CONSTANTES CONTABLES - CONOCIMIENTO DE NEGOCIO
# ==========================================
# Cuentas FIJAS para Librería
CTA_PROVISION_PUENTE = "21020302" # Provision por traslados 
CTA_CONSIGNACION_DR_ENTRADA = "7101" # Mercaderias entregadas proveedores 
CTA_CONSIGNACION_CR_ENTRADA = "8101" # Mercaderias recibidas 
CTA_CONSIGNACION_DR_SALIDA = "8101" # Mercaderias recibidas (Dev) 
CTA_CONSIGNACION_CR_SALIDA = "7101" # Mercaderias entregadas proveedores (Dev) 

# Cuentas BASE para Mapeos Dinámicos
CTA_BASE_GASTO = "410104" # Gasto por defecto
CTA_BASE_PT = "110603" # Producto Terminado
CTA_BASE_MP = "110601" # Materia Prima

# ==========================================
# INICIALIZACIÓN DE CONOCIMIENTO 
# ==========================================
# Mapeo de Unidades de Librería a sus cuentas de gasto/costo.
UNIDADES_LIBRERIA_EXPENSE_MAP = {
    "LIBRERÍA CENTRAL": CTA_BASE_GASTO,
    "LIBRERÍA CAMPUS": CTA_BASE_GASTO,
    "LIBRERÍA CENTRAL (B001)": CTA_BASE_GASTO, 
    "LIBRERÍA BODEGA": CTA_BASE_GASTO 
}

# Mapeo de Categorías a Cuentas de Inventario
INV_ACCOUNT_MAP = {
    'MATERIA PRIMA': CTA_BASE_MP,
    'PRODUCTO TERMINADO': CTA_BASE_PT,
    'DEFAULT': CTA_BASE_PT
}

UNIDADES_LIBRERIA_SET = set(UNIDADES_LIBRERIA_EXPENSE_MAP.keys())

# ==========================================
# FUNCIONES AUXILIARES
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

def gen_excels_memory(df_partidas_dict, mes_proceso):
    """Genera Excels de Nexus acumulando los valores por cuenta y concepto."""
    buffers = {}
    
    # Nombres elegantes para los archivos a descargar
    nombres_amigables = {
        "SEND": "Salidas_y_Traslados",
        "RECEIVE": "Ingresos_y_Recepciones",
        "ADJUST": "Ajustes_Internos"
    }
    
    for key, partidas_list in df_partidas_dict.items():
        if not partidas_list: continue
        
        df_nexus = pd.DataFrame(partidas_list)
        
        # --- CORRECCIÓN 1: ACUMULADOR (Agrupar y Sumar) ---
        # Esto junta todas las filas que tienen misma Cuenta y Concepto, y suma los montos
        df_grouped = df_nexus.groupby(["CUENTA", "VACIO", "CONCEPTO", "AUXILIAR"], as_index=False)[["DEBE", "HABER"]].sum()
        
        # Filtramos para que no salgan líneas con valor cero
        df_grouped = df_grouped[(df_grouped["DEBE"] > 0) | (df_grouped["HABER"] > 0)]
        
        # Ordenamos un poco para que se vea ordenado en el Excel
        df_grouped = df_grouped.sort_values(by=["CONCEPTO", "HABER", "DEBE"], ascending=[True, True, False])
        
        nexus_export = df_grouped[["CUENTA", "VACIO", "CONCEPTO", "DEBE", "HABER"]].copy() 
        
        # --- CORRECCIÓN 2: NOMBRES DE ARCHIVO AMIGABLES ---
        nombre_base = nombres_amigables.get(key, key)
        nombre_archivo = f"Partidas_{nombre_base}_Mes_{mes_proceso}.xlsx"
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            nexus_export.to_excel(writer, index=False, header=False, sheet_name="Partidas")
        buffers[nombre_archivo] = output.getvalue()
        
    return buffers

# ==========================================
# MÓDULO PRINCIPAL DE LA INTERFAZ
# ==========================================
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

                    partidas_dict = {"SEND": [], "RECEIVE": [], "ADJUST": []}
                    stats = {"consignacion": 0, "servicios": 0, "standard": 0, "adjustments": 0, "entries_to_libreria_ignored": 0}

                    for i, row in df.iterrows():
                        raw_sender = clean_value(row[c_sender])
                        raw_receiver = clean_value(row[c_receiver])
                        tipo = clean_value(row[c_tipo])
                        category = clean_value(row[c_category])
                        total = float(row[c_total])

                        if total == 0: continue

                        if raw_receiver in UNIDADES_LIBRERIA_SET and raw_sender not in UNIDADES_LIBRERIA_SET:
                            stats["entries_to_libreria_ignored"] += 1
                            continue 

                        if 'SERVICIO' in category:
                            stats["servicios"] += 1
                            continue 

                        if raw_sender in UNIDADES_LIBRERIA_SET:
                            unit_nexus_expense = current_expense_map.get(raw_sender, CTA_BASE_GASTO)
                            aux_unit = raw_sender.replace("LIBRERÍA ", "")
                            inv_acc = INV_ACCOUNT_MAP.get(category, INV_ACCOUNT_MAP['DEFAULT'])

                            desc_consignacion = f"RECONOCIMIENTO POR {tipo} DE PRODUCTOS EN CONSIGNACION DE LIBRERÍA {aux_unit}, MES {mes_proceso} DE {anio_proceso}."

                            if 'AJUS' in tipo or 'AJUSTE' in tipo:
                                stats["adjustments"] += 1
                                if raw_receiver in UNIDADES_LIBRERIA_SET:
                                    desc_adjust_entry = f"RECONOCIMIENTO DE ENTRADA POR AJUSTE DE PRODUCTO DE LIBRERÍA {aux_unit}, MES {mes_proceso} DE {anio_proceso}." 
                                    partidas_dict["ADJUST"].extend([
                                        gen_nexus_spec(inv_acc, desc_adjust_entry, total, 0),
                                        gen_nexus_spec(unit_nexus_expense, desc_adjust_entry, 0, total)
                                    ])
                                else:
                                    desc_adjust_exit = f"RECONOCIMIENTO DE SALIDA POR AJUSTE DE PRODUCTO DE LIBRERÍA {aux_unit}, MES {mes_proceso} DE {anio_proceso}."
                                    partidas_dict["ADJUST"].extend([
                                        gen_nexus_spec(unit_nexus_expense, desc_adjust_exit, total, 0),
                                        gen_nexus_spec(inv_acc, desc_adjust_exit, 0, total)
                                    ])
                            
                            elif 'CONSIGNACION' in category:
                                stats["consignacion"] += 1
                                if ('ENTRADA' in tipo or 'INGRESO' in tipo) and 'SALIDA' not in tipo:
                                    partidas_dict["RECEIVE"].extend([
                                        gen_nexus_spec(CTA_CONSIGNACION_DR_ENTRADA, desc_consignacion, total, 0, raw_sender),
                                        gen_nexus_spec(CTA_CONSIGNACION_CR_ENTRADA, desc_consignacion, 0, total, raw_sender)
                                    ])
                                else:
                                    desc_cons_dev = desc_consignacion.replace("POR ENTRADA", "POR DEVOLUCION").replace("POR INGRESO", "POR DEVOLUCION")
                                    partidas_dict["SEND"].extend([
                                        gen_nexus_spec(CTA_CONSIGNACION_DR_SALIDA, desc_cons_dev, total, 0, raw_sender),
                                        gen_nexus_spec(CTA_CONSIGNACION_CR_SALIDA, desc_cons_dev, 0, total, raw_sender)
                                    ])

                            else:
                                stats["standard"] += 1
                                desc_stand_send = f"RECONOCIMIENTO DE SALIDA POR TRASLADO DE PRODUCTO DE {aux_unit} A {raw_receiver}, MES {mes_proceso} DE {anio_proceso}." 

                                # Partida 1: SENDING Unit
                                partidas_dict["SEND"].extend([
                                    gen_nexus_spec(CTA_PROVISION_PUENTE, desc_stand_send, total, 0),
                                    gen_nexus_spec(inv_acc, desc_stand_send, 0, total)
                                ])

                                # Partida 2: RECEIVING Unit 
                                desc_stand_receive = desc_stand_send.replace("SALIDA POR TRASLADO", "ENTRADA POR TRASLADO")
                                partidas_dict["RECEIVE"].extend([
                                    gen_nexus_spec(INV_ACCOUNT_MAP.get('DEFAULT'), desc_stand_receive, total, 0),
                                    gen_nexus_spec(CTA_PROVISION_PUENTE, desc_stand_receive, 0, total)
                                ])

                    st.success(f"✅ Auditoría de Librería completada: {stats['standard'] + stats['consignacion'] + stats['adjustments']} operaciones procesadas y acumuladas.")
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