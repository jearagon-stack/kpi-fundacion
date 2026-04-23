import streamlit as st
import pandas as pd
from datetime import date
from utils import obtener_dataframe, conectar_hoja
import io

def mostrar_modulo_costos():
    st.title("☕ Contabilidad de Costos - Cafetería")

    tab1, tab2, tab3 = st.tabs(["📝 Generar Cierre", "🚚 Partidas (Próximamente)", "🔍 Histórico (Próximamente)"])

    # ==========================================
    # CONSTANTES Y BODEGAS OFICIALES
    # ==========================================
    BODEGAS_CAFETERIA = [
        "CAFETERIA CENTRAL", 
        "CAFETERIA ABASTECIMIENTO", 
        "CAFETERIA ICAS", 
        "CAFETERIA POLIDEPORTIVO", 
        "CAFETERIA EVENTOS", 
        "CAFETERIA JARDINES",
        "TERRAZA",
        "CENTRO SOHO"
    ]

    def generar_excel_bytes(filas):
        df_p = pd.DataFrame(filas)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_p.to_excel(writer, index=False, header=False, sheet_name='Hoja1')
        return output.getvalue()

    meses_texto = {1:"ENERO", 2:"FEBRERO", 3:"MARZO", 4:"ABRIL", 5:"MAYO", 6:"JUNIO", 7:"JULIO", 8:"AGOSTO", 9:"SEPTIEMBRE", 10:"OCTUBRE", 11:"NOVIEMBRE", 12:"DICIEMBRE"}

    # ==========================================
    # PESTAÑA 1: GENERAR CIERRE
    # ==========================================
    with tab1:
        st.subheader("1. Cierre mediante Kardex Valuado")

        if 'memoria_cierre' not in st.session_state:
            col1, col2, col3 = st.columns(3)
            with col1: mes_cierre = st.selectbox("Mes a costear:", range(1, 13), index=date.today().month - 1)
            with col2: anio_cierre = st.number_input("Año:", min_value=2024, max_value=2030, value=2026)
            with col3: unidad_cierre = st.selectbox("Unidad:", ["CAFETERIA"])

            st.divider()
            
            # --- HISTÓRICO DE VENTAS (PARA SUBSIDIOS) ---
            df_ventas_db = obtener_dataframe("Historico_Ventas")
            ventas_m = subsidio_m = 0.0
            if not df_ventas_db.empty:
                df_ventas_db['Fecha_DT'] = pd.to_datetime(df_ventas_db['Fecha'], format='%d/%m/%Y', errors='coerce')
                f_v = (df_ventas_db['Fecha_DT'].dt.year == anio_cierre) & (df_ventas_db['Fecha_DT'].dt.month == mes_cierre)
                ventas_m = pd.to_numeric(df_ventas_db[f_v]['Venta_Real'], errors='coerce').sum()
                subsidio_m = pd.to_numeric(df_ventas_db[f_v]['Subsidio_UCA'], errors='coerce').sum()

            porcentaje_subsidio = (subsidio_m / ventas_m) if ventas_m > 0 else 0.0
            st.info(f"📊 Ingresos del Mes: Ventas ${ventas_m:,.2f} | Subsidio ${subsidio_m:,.2f}")

            # --- CARGA DE ARCHIVOS ---
            col_u1, col_u2 = st.columns(2)
            with col_u1:
                arch_k_valuado = st.file_uploader("1. Kardex Valuado (Multi-archivos)", type=["xlsx"], accept_multiple_files=True)
            with col_u2:
                arch_k_resumen = st.file_uploader("2. Kardex Resumen (Para Categorías)", type=["xlsx"])

            if arch_k_valuado and arch_k_resumen:
                if st.button("⚙️ Procesar Auditoría y Cierre", type="primary", use_container_width=True):
                    with st.spinner("Analizando información..."):
                        try:
                            # 1. Diccionario de Categorías
                            df_res = pd.read_excel(arch_k_resumen, dtype=str)
                            df_res.columns = df_res.columns.astype(str).str.strip().str.upper()
                            c_cod_res = next((c for c in df_res.columns if 'IDPRODUCTO' in c or 'COD' in c), df_res.columns[0])
                            c_cat_res = next((c for c in df_res.columns if 'CATEGOR' in c), None)
                            mapa_cat = dict(zip(df_res[c_cod_res].str.strip(), df_res[c_cat_res].str.upper().str.strip())) if c_cat_res else {}

                            # 2. Cargar Kardex
                            dfs = [pd.read_excel(f, dtype=str) for f in arch_k_valuado]
                            df_k = pd.concat(dfs, ignore_index=True)
                            df_k.columns = df_k.columns.astype(str).str.strip().str.upper()

                            # Identificación de columnas
                            c_cod = next((c for c in df_k.columns if 'IDPRODUCTO' in c), None)
                            c_bod = next((c for c in df_k.columns if 'NOMBREBODEGA' in c), None)
                            c_doc = next((c for c in df_k.columns if 'DOCUMENTO' in c), None)
                            c_pref = next((c for c in df_k.columns if 'PREFI' in c), None)
                            c_ent_u = next((c for c in df_k.columns if 'ENTRADASUNID' in c), None)
                            c_ent_v = next((c for c in df_k.columns if 'ENTRADASVAL' in c), None)
                            c_sal_u = next((c for c in df_k.columns if 'SALIDASUNID' in c), None)
                            c_sal_v = next((c for c in df_k.columns if 'SALIDASVAL' in c), None)
                            c_costo = next((c for c in df_k.columns if 'COSTOPROMEDIO' in c), None)
                            c_saldo_v = next((c for c in df_k.columns if 'SALDOVAL' in c), None)
                            c_fec = next((c for c in df_k.columns if 'FECHA' in c), None)

                            # Herencia de Bodega y Filtro
                            df_k[c_cod] = df_k[c_cod].str.strip()
                            df_k[c_bod] = df_k.groupby(c_cod)[c_bod].ffill().bfill()
                            df_k = df_k[df_k[c_bod].str.strip().isin(BODEGAS_CAFETERIA)]

                            # Omitir Servicios
                            df_k['CAT_REAL'] = df_k[c_cod].map(mapa_cat).fillna('DESCONOCIDA')
                            df_k = df_k[df_k['CAT_REAL'] != 'SERVICIO']

                            # Conversión Numérica
                            for col in [c_ent_u, c_ent_v, c_sal_u, c_sal_v, c_costo, c_saldo_v]:
                                df_k[col] = pd.to_numeric(df_k[col], errors='coerce').fillna(0.0)

                            # ORDENAR POR FECHA
                            if c_fec:
                                df_k['FECHA_DT'] = pd.to_datetime(df_k[c_fec], errors='coerce')
                                df_k = df_k.sort_values(by=[c_cod, 'FECHA_DT']).reset_index(drop=True)

                            # --- EXTRACCIÓN DE TOTALES (CORREGIDO PARA EVITAR PÉRDIDAS) ---
                            # Saldo Inicial: Primera fila de cada producto en las bodegas filtradas
                            sum_ini = df_k[df_k[c_doc].astype(str).upper().str.contains('SALDO ANTERIOR')][c_saldo_v].sum()
                            # Saldo Final: Última fila de cada producto
                            sum_fin = df_k.groupby(c_cod).tail(1)[c_saldo_v].sum()
                            # Compras y Traslados
                            sum_com = df_k[df_k[c_pref].fillna('').str.upper() == 'CFE'][c_ent_v].sum()
                            sum_tr_in = df_k[df_k[c_pref].fillna('').str.upper() == 'TRD'][c_ent_v].sum()
                            sum_tr_out = df_k[df_k[c_pref].fillna('').str.upper().isin(['TRS', 'TRD']) & (df_k[c_sal_u] > 0)][c_sal_v].sum()

                            # --- AUDITORÍA DE CASCADA CON IDENTIFICADOR DE ORIGEN ---
                            def get_ref(df_sub, pref):
                                v = df_sub[df_sub[c_pref].fillna('').str.upper() == pref]
                                v = v[v[c_ent_u] > 0]
                                if v.empty: return pd.Series(dtype=float)
                                return v.groupby(c_cod).apply(lambda x: x[c_ent_v].sum() / x[c_ent_u].sum()).replace(0, pd.NA)

                            ref_cfe = get_ref(df_k, 'CFE')
                            ref_ini = df_k[df_k[c_doc].astype(str).upper().str.contains('SALDO ANTERIOR') & (df_k[c_costo] > 0)].groupby(c_cod)[c_costo].first()
                            ref_trd = get_ref(df_k, 'TRD')
                            ref_pro = get_ref(df_k, 'PRO')

                            # Construcción de Costo Base con Origen
                            base_audit = pd.DataFrame(index=df_k[c_cod].unique())
                            base_audit['C_BASE'] = ref_cfe
                            base_audit['ORIGEN'] = 'COMPRA'
                            
                            mask_ini = base_audit['C_BASE'].isna()
                            base_audit.loc[mask_ini, 'C_BASE'] = ref_ini
                            base_audit.loc[mask_ini, 'ORIGEN'] = 'SALDO ANTERIOR'
                            
                            mask_trd = base_audit['C_BASE'].isna()
                            base_audit.loc[mask_trd, 'C_BASE'] = ref_trd
                            base_audit.loc[mask_trd, 'ORIGEN'] = 'TRASLADO'
                            
                            mask_pro = base_audit['C_BASE'].isna()
                            base_audit.loc[mask_pro, 'C_BASE'] = ref_pro
                            base_audit.loc[mask_pro, 'ORIGEN'] = 'PRODUCCION'
                            
                            base_audit['C_BASE'] = base_audit['C_BASE'].fillna(0.0)

                            # Evaluación de Ventas
                            df_v = df_k[df_k[c_pref].fillna('').str.upper().isin(['FCF', 'CCF'])].copy()
                            df_v = df_v.merge(base_audit, left_on=c_cod, right_index=True, how='left')
                            df_v['VAR_P'] = (df_v[c_costo] / df_v['C_BASE'].replace(0, 1)) - 1
                            df_v['DIF_M'] = df_v[c_costo] - df_v['C_BASE']
                            
                            # Filtro: Desviaciones significativas
                            anomalias_all = df_v[(df_v['VAR_P'].abs() > 0.01) & (df_v['DIF_M'].abs() >= 0.019)].copy()

                            st.session_state['memoria_cierre'] = {
                                'inicial': sum_ini, 'compras': sum_com, 'traslados': sum_tr_in - sum_tr_out, 
                                'final': sum_fin, 'consumo': sum_ini + sum_com + (sum_tr_in - sum_tr_out) - sum_fin,
                                'subsidio': subsidio_m, 'anio': anio_cierre, 'mes': mes_cierre,
                                'anomalias': anomalias_all[[c_cod, c_pref, c_doc, 'C_BASE', c_costo, 'ORIGEN', 'DIF_M', 'VAR_P']]
                            }
                            st.rerun()

                        except Exception as e:
                            st.error(f"Error técnico: {e}")

        else:
            mem = st.session_state['memoria_cierre']
            st.success(f"📦 Datos en Memoria: {meses_texto[mem['mes']]} {mem['anio']}")
            
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Inicial", f"${mem['inicial']:,.2f}")
            c2.metric("Compras", f"${mem['compras']:,.2f}")
            c3.metric("Traslados", f"${mem['traslados']:,.2f}")
            c4.metric("Final", f"${mem['final']:,.2f}")
            c5.metric("Consumo", f"${mem['consumo']:,.2f}")

            st.divider()
            
            # --- NUEVA LÓGICA DE DOS PARTES ---
            st.subheader("🕵️ Auditoría de Costos")
            
            # Parte 1: El Filtro de "Elevación por otros motivos" (Alertas Críticas)
            st.error("⚠️ ALERTAS CRÍTICAS: Elevación de costo NO relacionada a Compras")
            anom_criticas = mem['anomalias'][mem['anomalias']['ORIGEN'] != 'COMPRA']
            if not anom_criticas.empty:
                st.dataframe(anom_criticas.style.format({'C_BASE':'${:.4f}', 'COSTOPROMEDIO':'${:.4f}', 'DIF_M':'${:.4f}', 'VAR_P':'{:.2%}'}), use_container_width=True)
            else:
                st.success("No hay elevaciones sospechosas por Ajustes o Traslados.")

            # Parte 2: El reporte completo (Para validación total)
            with st.expander("🔍 Ver Auditoría Completa (Incluye variaciones por Compra)"):
                st.write("Esta lista muestra todo lo que varió > 1%, incluyendo compras normales.")
                st.dataframe(mem['anomalias'].style.format({'C_BASE':'${:.4f}', 'COSTOPROMEDIO':'${:.4f}', 'DIF_M':'${:.4f}', 'VAR_P':'{:.2%}'}), use_container_width=True)

            if st.button("🗑️ Limpiar y Volver a Empezar"):
                del st.session_state['memoria_cierre']
                st.rerun()

    with tab2: st.warning("Pestaña deshabilitada temporalmente para pulir la auditoría.")
    with tab3: st.warning("Pestaña deshabilitada temporalmente para pulir la auditoría.")