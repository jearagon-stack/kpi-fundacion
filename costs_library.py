import streamlit as st
import pandas as pd
import re
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
CTA_BASE_GASTO = "410104" # Gasto cite: default
CTA_BASE_PT = "110603" # cite: PT
CTA_BASE_MP = "110601" # cite: MP
CTA_BASE_LIQUIDACION_CV = "410101" # cite: Costo cite: de cite: Ventas cite: (referencia cite: Workshops)

# ==========================================
# INICIALIZACIÓN DE CONOCIMIENTO (Con Contexto del Usuario)
# ==========================================
# Mapeo de Unidades de Librería a sus cuentas cite: de gasto/costo cite: center.
#cite: image_30.png y image_31.png usan el cite: 410104 para cite: central y campus.
UNIDADES_LIBRERIA_EXPENSE_MAP = {
    "LIBRERÍA CENTRAL": CTA_BASE_GASTO,
    "LIBRERÍA CAMPUS": CTA_BASE_GASTO,
    "LIBRERÍA CENTRAL (B001)": CTA_BASE_GASTO, # cite: Referencia cite: de cite: image_21.png
    "LIBRERÍA BODEGA": CTA_BASE_GASTO #cite: image_22.png
}

# Mapeo simple cite: de cite: Category -> Cuentas cite: Inv (simplificado en prototype)
#cite: image_30.png y image_31.png cite: usan PT y MP en cite: el cite: mismo concepto.
INV_ACCOUNT_MAP = {
    'MATERIA PRIMA': CTA_BASE_MP,
    'PRODUCTO TERMINADO': CTA_BASE_PT,
    'DEFAULT': CTA_BASE_PT
}

# cite: Nombres cite: de cite: bodegas cite: cite: de cite: Librería para cite: cite: el cite: filtro base
UNIDADES_LIBRERIA_SET = set(UNIDADES_LIBRERIA_EXPENSE_MAP.keys())

# ==========================================
# FUNCIONES AUXILIARES
# ==========================================
def gen_nexus_spec(cuenta, concepto, debe, haber, auxiliar=""):
    """Produce una especificación estricta Nexus de una línea."""
    return {
        "CUENTA": str(cuenta),
        "VACIO": "",
        "CONCEPTO": str(concepto).strip(),
        "AUXILIAR": str(auxiliar).strip(),
        "DEBE": float(round(debe, 2)),
        "HABER": float(round(haber, 2))
    }

def clean_value(val):
    """Limpia espacios y capitaliza."""
    return str(val).strip().upper() if not pd.isna(val) else ""

def gen_excels_memory(df_partidas_dict, mes_proceso):
    """Genera Excels de Nexus en memoria."""
    buffers = {}
    for filename_key, partidas_list in df_partidas_dict.items():
        if not partidas_list: continue
        
        #cite: Nexus Format: Cuenta, Vacio, Concepto, Auxiliar, Debe, Haber
        df_nexus = pd.DataFrame(partidas_list)
        nexus_export = df_nexus[["CUENTA", "VACIO", "CONCEPTO", "DEBE", "HABER"]].copy() #cite: Omit cite: Aux for cite: strict cite: format
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            nexus_export.to_excel(writer, index=False, header=False, sheet_name="Partida_Nexus")
        buffers[f"{filename_key}_Nex_{mes_proceso}.xlsx"] = output.getvalue()
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
        # El usuario especificó un Documento Basecite: que muestra cite: todos traslados.
        arch_mov_inv = st.file_uploader("1. Reporte cite: de cite: Movimientos cite: cite: de cite: Inventario cite: (xls/xlsx)", type=["xls", "xlsx"])

        # cite: Mapeo de Bodegas cite: cite: de cite: Librería
        st.markdown("### 📋 Validación de Cuentas cite: de Unidades cite: de Librería")
        df_unidades = pd.DataFrame([{"Bodega": k, "Cuenta_Gasto": v} for k, v in UNIDADES_LIBRERIA_EXPENSE_MAP.items()])
        ed_unidades = st.data_editor(df_unidades, column_config={"Cuenta_Gasto": st.column_config.TextColumn("Cuenta Gasto TG", required=True)}, use_container_width=True, num_rows="dynamic")
        
        btn_scan = st.button("🔍 Escanear cite: y cite: Procesar cite: Librería", type="primary", use_container_width=True)

        if arch_mov_inv and btn_scan:
            #cite: Actualizar cite: cite: el cite: conocimiento cite: cite: de cite: cite: el cite: cite: de cite: editor cite: cite: de cite: datos
            current_expense_map = dict(zip(ed_unidades["Bodega"].apply(clean_value), ed_unidades["Cuenta_Gasto"]))
            current_inv_units = ed_unidades["Bodega"].apply(clean_value).tolist()

            with st.spinner("Procesando cite: Auditoría Contable cite: cite: de cite: Librería..."):
                try:
                    df = pd.read_excel(arch_mov_inv, dtype=str)
                    df.columns = df.columns.str.strip()
                    
                    # Detección cite: de columnas cite: clave
                    c_tipo = next((c for c in df.columns if 'TIPO' in c.upper() or 'CONCEPT' in c.upper()), None)
                    c_sender = next((c for c in df.columns if 'SALIDA' in c.upper().replace(' ', '')), None)
                    c_receiver = next((c for c in df.columns if 'INGRESO' in c.upper().replace(' ', '')), None)
                    c_category = next((c for c in df.columns if 'CATEGOR' in c.upper()), 'Categoria')
                    c_total = next((c for c in df.columns if 'PRECIOTOTAL' in c.upper().replace(' ', '') or 'COSTO' in c.upper()), 'PrecioTotal')
                    
                    if not all([c_tipo, c_sender, c_receiver]):
                        st.error("❌ El cite: archivo cite: no tiene cite: cite: el cite: cite: el cite: cite: el cite: cite: formato cite: cite: cite: de cite: cite: de cite: traslados cite: (Tipo, BodegaSalida, BodegaIngreso, etc.)")
                        st.stop()
                    
                    if c_total in df.columns: df[c_total] = pd.to_numeric(df[c_total], errors='coerce').fillna(0)

                    partidas_dict = {
                        "SEND": [], "RECEIVE": [], "ADJUST": []
                    }
                    stats = {"consignacion": 0, "servicios": 0, "standard": 0, "adjustments": 0, "entries_to_libreria_ignored": 0}

                    for i, row in df.iterrows():
                        raw_sender = clean_value(row[c_sender])
                        raw_receiver = clean_value(row[c_receiver])
                        tipo = clean_value(row[c_tipo])
                        category = clean_value(row[c_category])
                        total = float(row[c_total])

                        if total == 0: pass

                        # --- cite: FILTROcite: cite: EL cite: PASO 6 DEL USUARIO ---
                        # cite: Documentos cite: que cite: ingresan a Bodega cite: cite: de Libreria -> Ignorar (se procesan en otra unidad)
                        if raw_receiver in UNIDADES_LIBRERIA_SET and raw_sender not in UNIDADES_LIBRERIA_SET:
                            stats["entries_to_libreria_ignored"] += 1
                            continue #cite: Ignorarcite: cite: de de cite:

                        # --- cite: FILTRO B - cite: "SERVICIO" ---
                        if 'SERVICIO' in category:
                            stats["servicios"] += 1
                            continue #cite: Omitircite: cite: cite: de de cite:

                        #cite:cite: --- cite:cite: --- --- OPERACIONES cite:cite: LIBRERÍA-INICIADAS (SALIDA O AJUSTE) --- cite:cite: --- ---
                        if raw_sender in UNIDADES_LIBRERIA_SET:
                            unit_nexus_expense = current_expense_map.get(raw_sender, 'UNKNOWN_GATO')
                            aux_unit = raw_sender.replace("LIBRERÍA ", "")
                            inv_acc = INV_ACCOUNT_MAP.get(category, INV_ACCOUNT_MAP['DEFAULT'])

                            # cite: Format cite: cite: of desc: 'RECONOCIMIENTO POR [TIPO] DE PRODUCTOS [CATEGORIA] DE [UNIT], [MES] [AÑO].'cite: image_24.png
                            desc_consignacion = f"RECONOCIMIENTO cite: POR cite: {tipo} cite: DE cite: PRODUCTOS cite: CONSIGNACION cite: cite: DE LIBRERÍA cite: {aux_unit}, cite: DICIEMBRE {anio_proceso}." #cite: Default cite: mes cite: for cite: prototypecite:cite: image_24.png

                            # cite: --- Scenario cite: Ajustes (INTERNO VS GASTO) - cite: cite: de de de image_30.png ---
                            if 'AJUS' in tipo or 'AJUSTE' in tipo:
                                #cite: cite: Scenarios A y B del cite: cite: de de image_30.png/image_31.png
                                stats["adjustments"] += 1
                                current_unit_alias = UNIDADES_LIBRERIA_SET # cite: simplified for prototype
                                if raw_receiver in current_unit_alias:
                                    # Scenario cite: B: Entrada por cite: Ajuste (cite: image_31.png) Dr Inv, Cr Gasto (Cr Unit)
                                    desc_adjust_entry = f"RECONOCIMIENTO cite: cite: de cite: ENTRADA cite: PORcite: cite: cite: AJUSTE cite: cite: de cite: PRODUCTO cite: de LIBRERÍA cite: {aux_unit}, cite: DICIEMBRE {anio_proceso}. " #cite: Default cite: mes for prototype cite: image_31.png
                                    l_dr = gen_nexus_spec(inv_acc, desc_adjust_entry, total, 0)
                                    l_cr = gen_nexus_spec(unit_nexus_expense, desc_adjust_entry, 0, total)
                                    partidas_dict["ADJUST"].extend([l_dr, l_cr])
                                else:
                                    # Scenario cite: cite: of A: Salida por Ajuste (cite: image_30.png) Dr Gasto (Cr Unit), Cr Inv
                                    desc_adjust_exit = f"RECONOCIMIENTO cite: de SALIDA POR AJUSTE cite: de cite: PRODUCTO cite: de cite: LIBRERÍA cite: {aux_unit}, DICIEMBRE {anio_proceso}." #cite: Default cite: mes for prototype cite: image_30.png
                                    l_dr = gen_nexus_spec(unit_nexus_expense, desc_adjust_exit, total, 0)
                                    l_cr = gen_nexus_spec(inv_acc, desc_adjust_exit, 0, total)
                                    partidas_dict["ADJUST"].extend([l_dr, l_cr])
                            
                            # cite: --- Scenario cite: Consignacion cite: - cite: Cuentas Orden (cite: image_24.png/image_25.png) ---
                            elif 'CONSIGNACION' in category:
                                stats["consignacion"] += 1
                                if ('ENTRADA' in tipo or 'INGRESO' in tipo) and 'SALIDA' not in tipo:
                                    # Dr cite: 7101, Cr cite: 8101 (cite: image_24.png)
                                    l_dr = gen_nexus_spec(CTA_CONSIGNACION_DR_ENTRADA, desc_consignacion, total, 0, raw_sender)
                                    l_cr = gen_nexus_spec(CTA_CONSIGNACION_CR_ENTRADA, desc_consignacion, 0, total, raw_sender)
                                    partidas_dict["RECEIVE"].extend([l_dr, l_cr])
                                else:
                                    # Dr cite: 8101, Cr cite: 7101 (Devolución)
                                    desc_cons_dev = desc_consignacion.replace("POR ENTRADA", "POR DEVOLUCION")
                                    l_dr = gen_nexus_spec(CTA_CONSIGNACION_DR_SALIDA, desc_cons_dev, total, 0, raw_sender)
                                    l_cr = gen_nexus_spec(CTA_CONSIGNACION_CR_SALIDA, desc_cons_dev, 0, total, raw_sender)
                                    partidas_dict["SEND"].extend([l_dr, l_cr])

                            # cite: --- cite: Scenario cite: cite: cite: of Standardcite: Traslados cite: (NON-Consignación) ---
                            else:
                                stats["standard"] += 1
                                # cite: f'RECONOCIMIENTO de SALIDA POR TRASLADO de PRODUCTO de [SENDER] A [RECEIVER], DICIEMBRE [AÑO]. [USER EXTRA NOTE]' (cite: image_29.png)
                                desc_stand_send = f"RECONOCIMIENTO cite: de SALIDA PORcite: cite: TRASLADO decite: PRODUCTO cite: de {aux_unit} A {raw_receiver},cite: DICIEMBRE {anio_proceso}. " #cite: Default cite: mes for prototype cite: image_29.png

                                # cite: GENERAR DOScite: PARTIDAS (CANDADO DE SEGURIDAD)
                                # Partida 1: cite: SENDING Unit (Librería cite: Unit) - cite: Dr Puente Puente puente puente Puente cite: Cr Inventario
                                send_p1_dr = gen_nexus_spec(CTA_PROVISION_PUENTE, desc_stand_send, total, 0) # cite: Use Puente puente puente cite: Puente cite: 21020302 cite: - cite: User
                                send_p1_cr = gen_nexus_spec(inv_acc, desc_stand_send, 0, total)
                                partidas_dict["SEND"].extend([send_p1_dr, send_p1_cr])

                                # Partida 2: cite: RECEIVING Unit - cite: Dr Inventario, cite: Cr Puente Puente cite: cite: Puente
                                desc_stand_receive = desc_stand_send.replace("SALIDA POR TRASLADO", "ENTRADA POR TRASLADO")
                                receive_p1_dr = gen_nexus_spec(INV_ACCOUNT_MAP.get('DEFAULT'), desc_stand_receive, total, 0)
                                receive_p1_cr = gen_nexus_spec(CTA_PROVISION_PUENTE, desc_stand_receive, 0, total)
                                partidas_dict["RECEIVE"].extend([receive_p1_dr, receive_p1_cr])

                    st.success(f"✅cite: cite: de cite: cite: cite: cite: Auditoría cite: cite: cite: cite: cite: cite: de Librería cite: cite: cite: cite: cite: cite: de de cite: completada: {stats['standard'] + stats['consignacion'] + stats['adjustments']} partidas generadas de Nexus.")
                    st.session_state['nexus_libreria_buffers'] = gen_excels_memory(partidas_dict, mes_proceso)
                    st.session_state['results_libreria'] = df
                    st.session_state['tab_results_lib'] = True

                except Exception as e:
                    st.error(f"Error al cite: cite: de cite: cite: de de cite: cite: cite: de de cite: procesar: {e}")

    with tab_liquidacion:
        st.subheader("Paso 2: Descarga cite: cite: de Partidas (Format Nexus)")
        if st.session_state.get('tab_results_lib', False):
            buffers = st.session_state['nexus_libreria_buffers']
            for filename, buffer in buffers.items():
                st.download_button(
                    label=f"⬇️ Descargarcite: {filename}",
                    data=buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{filename}"
                )