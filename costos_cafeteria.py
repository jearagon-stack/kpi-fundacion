import streamlit as st
import pandas as pd
from datetime import date
from utils import obtener_dataframe, conectar_hoja
import io

def mostrar_modulo_costos():
    st.title("☕ Contabilidad de Costos - Cafetería")

    tab1, tab2, tab3 = st.tabs(["📝 Generar Cierre", "🚚 Partidas de Traslados", "🔍 Consultar Histórico"])

    # ==========================================
    # FUNCIONES DE APOYO Y MAPEO DE UNIDADES
    # ==========================================
    mapa_subunidades = {
        "CAFETERIA": ["TERRAZA", "CENTRO SOHO", "CAFETERIA CENTRAL", "CAFETERIA", "CAFETERIA ABASTECIMIENTO"],
        "DESPENSA":  ["DESPENSA"]
    }

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

    destinos_ignorados = [
        "CAFETERIA EVENTOS", 
        "CAFETERIA JARDINES", 
        "CAFETERIA POLIDEPORTIVO", 
        "CAFETERIA ICAS", 
        "PRODUCCION CAFETERIA CENTRAL",
        "CAFETERIA CENTRAL"
    ]

    def es_de_unidad(bodega, unidad_param):
        b = str(bodega).strip().upper()
        if unidad_param == "CAFETERIA":
            return ("CAFETERIA" in b) or (b in ["TERRAZA", "CENTRO SOHO"])
        elif unidad_param == "DESPENSA":
            return "DESPENSA" in b
        return b == unidad_param.upper()

    def generar_excel_bytes(filas):
        df_p = pd.DataFrame(filas)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_p.to_excel(writer, index=False, header=False, sheet_name='Hoja1')
        return output.getvalue()

    def extraer_subunidad(nombre_archivo, unidad_base):
        nombre_upper = str(nombre_archivo).upper()
        subunidades = ["POLIDEPORTIVO", "ICAS", "JARDINES", "EVENTOS", "CENTRAL", "ABASTECIMIENTO"]
        for sub in subunidades:
            if sub in nombre_upper: return f"{unidad_base} {sub}"
        return unidad_base

    meses_texto = {1:"ENERO", 2:"FEBRERO", 3:"MARZO", 4:"ABRIL", 5:"MAYO", 6:"JUNIO", 7:"JULIO", 8:"AGOSTO", 9:"SEPTIEMBRE", 10:"OCTUBRE", 11:"NOVIEMBRE", 12:"DICIEMBRE"}
    lista_meses = list(meses_texto.values())

    # ==========================================
    # PESTAÑA 1: GENERAR CIERRE (NUEVO KARDEX)
    # ==========================================
    with tab1:
        st.subheader("1. Cierre Contable mediante Kardex")

        if 'memoria_cierre' not in st.session_state:
            col_config1, col_config2 = st.columns([2, 1])
            with col_config1:
                tipo_cierre = st.radio("Tipo de proceso:", ["Mensual Estándar", "Consolidación Especial (Multi-mes)"], horizontal=True)

            col1, col2, col3 = st.columns(3)
            with col1: mes_cierre = st.selectbox("Mes a costear:", range(1, 13), index=date.today().month - 1)
            with col2: anio_cierre = st.number_input("Año:", min_value=2024, max_value=2030, value=2026)
            with col3: unidad_cierre = st.selectbox("Unidad:", ["CAFETERIA", "DESPENSA"])
            
            es_consolidado = (tipo_cierre == "Consolidación Especial (Multi-mes)")
            pesos, nombres_meses = [], []

            if es_consolidado:
                num_meses = st.number_input("¿Dividir en cuántos meses?", min_value=1, max_value=12, value=mes_cierre)
                st.markdown(f"**Distribución de costos para {num_meses} meses finalizando en {meses_texto[mes_cierre]}:**")
                for i in range(0, num_meses, 3):
                    cols = st.columns(3)
                    for j in range(3):
                        idx_actual = i + j
                        if idx_actual < num_meses:
                            idx_sugerido = mes_cierre - (num_meses - 1) + idx_actual
                            if idx_sugerido < 1: idx_sugerido += 12
                            with cols[j]:
                                n = st.selectbox(f"Mes {idx_actual+1}:", options=lista_meses, index=idx_sugerido-1, key=f"n_gen_{idx_actual}")
                                p = st.number_input(f"% Venta {n}", min_value=0.0, max_value=100.0, value=100.0/num_meses, key=f"p_gen_{idx_actual}")
                                pesos.append(p / 100); nombres_meses.append(n)

            st.divider()
            
            # --- HISTÓRICO DE VENTAS (PARA SUBSIDIOS) ---
            df_ventas = obtener_dataframe("Historico_Ventas")
            ventas_mes = subsidio_mes = 0.0
            if not df_ventas.empty:
                df_ventas['Fecha_DT'] = pd.to_datetime(df_ventas['Fecha'], format='%d/%m/%Y', errors='coerce')
                filtro_v = (df_ventas['Fecha_DT'].dt.year == anio_cierre) & (df_ventas['Unidad'] == unidad_cierre)
                if es_consolidado:
                    meses_indices = [list(meses_texto.keys())[list(meses_texto.values()).index(m)] for m in nombres_meses]
                    filtro_v &= df_ventas['Fecha_DT'].dt.month.isin(meses_indices)
                else:
                    filtro_v &= (df_ventas['Fecha_DT'].dt.month == mes_cierre)
                ventas_mes = pd.to_numeric(df_ventas[filtro_v]['Venta_Real'], errors='coerce').sum()
                subsidio_mes = pd.to_numeric(df_ventas[filtro_v]['Subsidio_UCA'], errors='coerce').sum()

            porcentaje_subsidio = (subsidio_mes / ventas_mes) if ventas_mes > 0 else 0.0
            st.info(f"📊 Ingresos Totales Periodo: Ventas ${ventas_mes:,.2f} | Subsidio ${subsidio_mes:,.2f}")

            costo_diferido_anterior = st.number_input("Costo Diferido de Arrastre (110602):", min_value=0.0, value=0.0)
            
            # --- CARGA DE ARCHIVOS ---
            col_u1, col_u2 = st.columns(2)
            with col_u1:
                arch_k_valuado = st.file_uploader("1. Kardex Valuado (Multi-archivos)", type=["xlsx"], accept_multiple_files=True)
            with col_u2:
                arch_k_resumen = st.file_uploader("2. Kardex Resumen (Categorías)", type=["xlsx"])

            if arch_k_valuado and arch_k_resumen:
                if 'huerfanos_df' not in st.session_state:
                    forzar_calculo = st.checkbox("⚠️ Forzar cálculo ciego (Omitir revisión de cuentas)")
                    
                    if st.button("⚙️ Procesar Archivos y Generar Memoria", type="primary", use_container_width=True):
                        with st.spinner("Ensamblando Kardex, leyendo fechas y aplicando cascada de costos..."):
                            try:
                                # 1. DICCIONARIO DE CATEGORÍAS (KARDEX RESUMEN)
                                df_res = pd.read_excel(arch_k_resumen, dtype=str)
                                df_res.columns = df_res.columns.astype(str).str.strip().str.upper()
                                c_cod_res = next((c for c in df_res.columns if 'IDPRODUCTO' in c or 'COD' in c), df_res.columns[0])
                                c_cat_res = next((c for c in df_res.columns if 'CATEGOR' in c), None)
                                mapa_cat = dict(zip(df_res[c_cod_res].str.strip(), df_res[c_cat_res].str.upper().str.strip())) if c_cat_res else {}

                                # 2. UNIFICAR KARDEX VALUADO (Multi-Archivo)
                                dfs_k = [pd.read_excel(f, dtype=str) for f in arch_k_valuado]
                                df_k = pd.concat(dfs_k, ignore_index=True)
                                df_k.columns = df_k.columns.astype(str).str.strip().str.upper()

                                # Mapeo de Columnas
                                c_cod = next((c for c in df_k.columns if 'IDPRODUCTO' in c), None)
                                c_bod = next((c for c in df_k.columns if 'BODEGA' in c), None)
                                c_doc = next((c for c in df_k.columns if 'DOCUMENTO' in c), None)
                                c_pref = next((c for c in df_k.columns if 'PREFI' in c), None)
                                c_ent_u = next((c for c in df_k.columns if 'ENTRADASUNID' in c), None)
                                c_ent_v = next((c for c in df_k.columns if 'ENTRADASVAL' in c), None)
                                c_sal_u = next((c for c in df_k.columns if 'SALIDASUNID' in c), None)
                                c_sal_v = next((c for c in df_k.columns if 'SALIDASVAL' in c), None)
                                c_costo = next((c for c in df_k.columns if 'COSTOPROMEDIO' in c), None)
                                c_saldo_u = next((c for c in df_k.columns if 'SALDOUNID' in c), None)
                                c_saldo_v = next((c for c in df_k.columns if 'SALDOVAL' in c), None)
                                c_fec = next((c for c in df_k.columns if 'FECHA' in c), None)

                                # Limpieza y Ordenamiento por fecha
                                df_k[c_cod] = df_k[c_cod].str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
                                if c_fec:
                                    df_k['FECHA_DT'] = pd.to_datetime(df_k[c_fec], errors='coerce')
                                    df_k = df_k.sort_values(by=[c_cod, 'FECHA_DT']).reset_index(drop=True)

                                # Herencia de Bodega para 'Saldo anterior'
                                df_k[c_bod] = df_k.groupby(c_cod)[c_bod].ffill().bfill()
                                
                                # Filtros
                                bodegas_validas = BODEGAS_CAFETERIA if unidad_cierre == "CAFETERIA" else ["DESPENSA"]
                                df_k = df_k[df_k[c_bod].str.strip().isin(bodegas_validas)]
                                
                                df_k['CAT_REAL'] = df_k[c_cod].map(mapa_cat).fillna('DESCONOCIDA')
                                df_k = df_k[df_k['CAT_REAL'] != 'SERVICIO']

                                basura = ['G222', 'G231', '21455979']
                                df_k = df_k[~df_k[c_cod].isin(basura)]

                                for col in [c_ent_u, c_ent_v, c_sal_u, c_sal_v, c_costo, c_saldo_u, c_saldo_v]:
                                    df_k[col] = pd.to_numeric(df_k[col], errors='coerce').fillna(0.0)

                                df_k['Prefijo_Upper'] = df_k[c_pref].fillna('').astype(str).str.upper().str.strip()
                                df_k['Doc_Upper'] = df_k[c_doc].fillna('').astype(str).str.upper().str.strip()

                                # 3. DICCIONARIO CONTABLE (DE LA NUBE) Y CUENTAS HUÉRFANAS
                                df_dic_ctas = obtener_dataframe("Categorias_Costos")
                                df_dic_ctas['Codigo'] = df_dic_ctas['Codigo'].astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
                                mapa_cuentas = dict(zip(df_dic_ctas['Codigo'], df_dic_ctas['Cuenta_Contable']))

                                df_k['CUENTA'] = df_k[c_cod].map(mapa_cuentas)
                                mask_huerfanas = df_k['CUENTA'].isna() | df_k['CUENTA'].astype(str).str.strip().str.upper().isin(["", "NAN", "NAT", "NONE"])
                                
                                df_faltantes = df_k[mask_huerfanas][[c_cod, 'CAT_REAL', c_bod]].drop_duplicates(subset=[c_cod])
                                df_faltantes.columns = ['Codigo', 'Categoria', 'ORIGEN_ARCHIVO']

                                if not df_faltantes.empty and not forzar_calculo:
                                    st.session_state['pre_proceso'] = {'df_k': df_k, 'c_cod': c_cod, 'c_bod': c_bod, 'c_ent_v': c_ent_v, 'c_sal_v': c_sal_v, 'c_saldo_v': c_saldo_v, 'c_pref': c_pref, 'c_doc': c_doc, 'c_ent_u': c_ent_u, 'c_costo': c_costo, 'c_sal_u': c_sal_u}
                                    st.session_state['huerfanos_df'] = df_faltantes
                                    st.rerun()

                                # 4. MATEMÁTICA DE CONSUMO POR CUENTA (Agrupado para las Partidas)
                                df_k['CUENTA'] = df_k['CUENTA'].fillna("SIN CUENTA REGISTRADA")
                                df_k = df_k[~df_k['CUENTA'].astype(str).str.upper().isin(["NO APLICA", "0", "0.0", "OMITIDO_MANUAL"])]

                                # Extracción de rubros globales y por cuenta
                                cuentas_validas = df_k['CUENTA'].unique()
                                consumo_por_cuenta = {}
                                total_ini_val = total_com_val = total_tras_val = total_fin_val = costo_operativo = 0.0

                                # Evaluamos cuenta por cuenta iterando sobre el dataframe filtrado
                                for cta in cuentas_validas:
                                    df_cta = df_k[df_k['CUENTA'] == cta]
                                    
                                    # INI: Primer Saldo Anterior cronológico por producto de esta cuenta
                                    ini_cta = df_cta[df_cta['Doc_Upper'].str.contains('SALDO ANTERIOR')].groupby(c_cod).first()[c_saldo_v].sum()
                                    # FIN: Última fila cronológica por producto de esta cuenta
                                    fin_cta = df_cta.groupby(c_cod).tail(1)[c_saldo_v].sum()
                                    # COMPRAS y TRASLADOS de la cuenta
                                    com_cta = df_cta[df_cta['Prefijo_Upper'] == 'CFE'][c_ent_v].sum()
                                    tras_in_cta = df_cta[df_cta['Prefijo_Upper'] == 'TRD'][c_ent_v].sum()
                                    tras_out_cta = df_cta[df_cta['Prefijo_Upper'].isin(['TRS', 'TRD']) & (df_cta[c_sal_u] > 0)][c_sal_v].sum()
                                    
                                    tras_neto_cta = tras_in_cta - tras_out_cta
                                    
                                    total_ini_val += ini_cta
                                    total_com_val += com_cta
                                    total_tras_val += tras_neto_cta
                                    total_fin_val += fin_cta
                                    
                                    val_consumo = ini_cta + com_cta + tras_neto_cta - fin_cta
                                    if abs(val_consumo) > 0.01:
                                        consumo_por_cuenta[cta] = val_consumo
                                        costo_operativo += val_consumo

                                costo_dif_mes = float(costo_operativo) * float(porcentaje_subsidio)
                                costo_real = float(costo_operativo) - float(costo_dif_mes) + float(costo_diferido_anterior)

                                # 5. AUDITORÍA DE CASCADA (Doble Filtro)
                                def get_ref(df_sub, pref):
                                    v = df_sub[df_sub['Prefijo_Upper'] == pref]
                                    v = v[v[c_ent_u] > 0]
                                    return v.groupby(c_cod).apply(lambda x: x[c_ent_v].sum() / x[c_ent_u].sum()).replace(0, None)

                                ref_cfe = get_ref(df_k, 'CFE')
                                ref_ini = df_k[df_k['Doc_Upper'].str.contains('SALDO ANTERIOR') & (df_k[c_costo] > 0)].groupby(c_cod)[c_costo].first()
                                ref_trd = get_ref(df_k, 'TRD')
                                ref_pro = get_ref(df_k, 'PRO')
                                ref_eaj = get_ref(df_k, 'EAJ')

                                base_c = pd.DataFrame(index=df_k[c_cod].unique())
                                base_c['C_BASE'] = ref_cfe.fillna(ref_ini).fillna(ref_trd).fillna(ref_pro).fillna(ref_eaj).fillna(0)

                                df_v = df_k[df_k['Prefijo_Upper'].isin(['FCF', 'CCF'])].copy()
                                df_v = df_v.merge(base_c, left_on=c_cod, right_index=True, how='left')
                                df_v['VAR_P'] = (df_v[c_costo] / df_v['C_BASE'].replace(0, 1)) - 1
                                df_v['DIF_M'] = df_v[c_costo] - df_v['C_BASE']
                                
                                anomalias = df_v[(df_v['VAR_P'].abs() > 0.01) & (df_v['DIF_M'].abs() >= 0.019)]

                                st.session_state['memoria_cierre'] = {
                                    'df_k_raw': df_k, # Guardamos un snapshot por si queremos extraer detalles
                                    'grp_ini_sum': total_ini_val, 'grp_comp_sum': total_com_val, 
                                    'grp_tras_sum': total_tras_val, 'grp_fin_sum': total_fin_val,
                                    'costo_dif_mes': costo_dif_mes, 'costo_real': costo_real, 'costo_operativo': costo_operativo,
                                    'costo_diferido_anterior': costo_diferido_anterior, 'consumo_por_cuenta': consumo_por_cuenta,
                                    'mes_cierre': mes_cierre, 'anio_cierre': anio_cierre, 'unidad_cierre': unidad_cierre,
                                    'es_consolidado': es_consolidado, 'pesos': pesos, 'nombres_meses': nombres_meses,
                                    'anomalias': anomalias[[c_cod, 'Prefijo_Upper', c_doc, 'C_BASE', c_costo, 'DIF_M', 'VAR_P']]
                                }
                                
                                if 'huerfanos_df' in st.session_state: del st.session_state['huerfanos_df']
                                if 'pre_proceso' in st.session_state: del st.session_state['pre_proceso']
                                
                                st.rerun()

                            except Exception as e: 
                                st.error(f"Error procesando Kardex: {e}")

                else:
                    # ==========================================
                    # RESOLUCIÓN DE CUENTAS HUÉRFANAS
                    # ==========================================
                    st.error("🚨 ALERTA: PRODUCTOS SIN CUENTA CONTABLE DETECTADOS")
                    st.write("El sistema ha pausado el cálculo para que decidas qué hacer con estos códigos. Puedes asignarles una cuenta, usar su nombre de categoría o tirarlos a la basura (omitirlos) para que no afecten tu costo.")
                    
                    df_edit = st.session_state['huerfanos_df'].copy()
                    if 'Accion' not in df_edit.columns:
                        df_edit['Accion'] = "Omitir (No sumar al costo)"
                        df_edit['Cuenta_Manual'] = ""

                    edited_df = st.data_editor(
                        df_edit,
                        column_config={
                            "Codigo": "Código Producto",
                            "Categoria": "Categoría Nativa",
                            "ORIGEN_ARCHIVO": "Bodega Base",
                            "Accion": st.column_config.SelectboxColumn(
                                "¿Qué hacer?",
                                options=["Omitir (No sumar al costo)", "Usar Categoría Nativa", "Escribir Cuenta Manual"],
                                required=True
                            ),
                            "Cuenta_Manual": st.column_config.TextColumn(
                                "Cuenta (Si es Manual)",
                                help="Si elegiste 'Escribir Cuenta Manual', teclea aquí la cuenta (Ej: 110603 o Materia Prima)"
                            )
                        },
                        hide_index=True,
                        disabled=["Codigo", "Categoria", "ORIGEN_ARCHIVO"],
                        use_container_width=True
                    )

                    col_b1, col_b2 = st.columns(2)
                    if col_b1.button("✅ Aplicar Decisiones y Generar Cierre", type="primary", use_container_width=True):
                        with st.spinner("Aplicando reglas y calculando..."):
                            df_k = st.session_state['pre_proceso']['df_k']
                            c_cod = st.session_state['pre_proceso']['c_cod']
                            c_saldo_v = st.session_state['pre_proceso']['c_saldo_v']
                            c_ent_v = st.session_state['pre_proceso']['c_ent_v']
                            c_sal_v = st.session_state['pre_proceso']['c_sal_v']
                            c_sal_u = st.session_state['pre_proceso']['c_sal_u']
                            c_ent_u = st.session_state['pre_proceso']['c_ent_u']
                            c_costo = st.session_state['pre_proceso']['c_costo']

                            codigos_omitir = edited_df[edited_df['Accion'] == 'Omitir (No sumar al costo)']['Codigo'].tolist()
                            df_asignar = edited_df[edited_df['Accion'] == 'Escribir Cuenta Manual']
                            df_categoria = edited_df[edited_df['Accion'] == 'Usar Categoría Nativa']

                            mask_omitir = df_k[c_cod].isin(codigos_omitir)
                            if mask_omitir.any(): df_k.loc[mask_omitir, 'CUENTA'] = 'OMITIDO_MANUAL'
                            
                            for _, row in df_asignar.iterrows():
                                c_manual = str(row['Cuenta_Manual']).strip()
                                if c_manual == "": c_manual = "SIN CUENTA REGISTRADA"
                                df_k.loc[df_k[c_cod] == row['Codigo'], 'CUENTA'] = c_manual
                            
                            for _, row in df_categoria.iterrows():
                                cat_val = str(row['Categoria']).strip()
                                df_k.loc[df_k[c_cod] == row['Codigo'], 'CUENTA'] = cat_val
                            
                            df_k['CUENTA'] = df_k['CUENTA'].fillna("SIN CUENTA REGISTRADA")
                            df_k = df_k[~df_k['CUENTA'].astype(str).str.upper().isin(["NO APLICA", "0", "0.0", "OMITIDO_MANUAL"])]

                            cuentas_validas = df_k['CUENTA'].unique()
                            consumo_por_cuenta = {}
                            total_ini_val = total_com_val = total_tras_val = total_fin_val = costo_operativo = 0.0

                            for cta in cuentas_validas:
                                df_cta = df_k[df_k['CUENTA'] == cta]
                                ini_cta = df_cta[df_cta['Doc_Upper'].str.contains('SALDO ANTERIOR')].groupby(c_cod).first()[c_saldo_v].sum()
                                fin_cta = df_cta.groupby(c_cod).tail(1)[c_saldo_v].sum()
                                com_cta = df_cta[df_cta['Prefijo_Upper'] == 'CFE'][c_ent_v].sum()
                                tras_in_cta = df_cta[df_cta['Prefijo_Upper'] == 'TRD'][c_ent_v].sum()
                                tras_out_cta = df_cta[df_cta['Prefijo_Upper'].isin(['TRS', 'TRD']) & (df_cta[c_sal_u] > 0)][c_sal_v].sum()
                                tras_neto_cta = tras_in_cta - tras_out_cta
                                
                                total_ini_val += ini_cta
                                total_com_val += com_cta
                                total_tras_val += tras_neto_cta
                                total_fin_val += fin_cta
                                
                                val_consumo = ini_cta + com_cta + tras_neto_cta - fin_cta
                                if abs(val_consumo) > 0.01:
                                    consumo_por_cuenta[cta] = val_consumo
                                    costo_operativo += val_consumo
                    
                            costo_dif_mes = float(costo_operativo) * float(porcentaje_subsidio)
                            costo_real = float(costo_operativo) - float(costo_dif_mes) + float(costo_diferido_anterior)

                            # Auditoría (Sin recalcular, solo aplicando lógica igual)
                            def get_ref(df_sub, pref):
                                v = df_sub[df_sub['Prefijo_Upper'] == pref]
                                v = v[v[c_ent_u] > 0]
                                return v.groupby(c_cod).apply(lambda x: x[c_ent_v].sum() / x[c_ent_u].sum()).replace(0, None)

                            ref_cfe = get_ref(df_k, 'CFE')
                            ref_ini = df_k[df_k['Doc_Upper'].str.contains('SALDO ANTERIOR') & (df_k[c_costo] > 0)].groupby(c_cod)[c_costo].first()
                            ref_trd = get_ref(df_k, 'TRD')
                            ref_pro = get_ref(df_k, 'PRO')
                            ref_eaj = get_ref(df_k, 'EAJ')

                            base_c = pd.DataFrame(index=df_k[c_cod].unique())
                            base_c['C_BASE'] = ref_cfe.fillna(ref_ini).fillna(ref_trd).fillna(ref_pro).fillna(ref_eaj).fillna(0)

                            df_v = df_k[df_k['Prefijo_Upper'].isin(['FCF', 'CCF'])].copy()
                            df_v = df_v.merge(base_c, left_on=c_cod, right_index=True, how='left')
                            df_v['VAR_P'] = (df_v[c_costo] / df_v['C_BASE'].replace(0, 1)) - 1
                            df_v['DIF_M'] = df_v[c_costo] - df_v['C_BASE']
                            
                            anomalias = df_v[(df_v['VAR_P'].abs() > 0.01) & (df_v['DIF_M'].abs() >= 0.019)]

                            st.session_state['memoria_cierre'] = {
                                'df_k_raw': df_k,
                                'grp_ini_sum': total_ini_val, 'grp_comp_sum': total_com_val, 
                                'grp_tras_sum': total_tras_val, 'grp_fin_sum': total_fin_val,
                                'costo_dif_mes': costo_dif_mes, 'costo_real': costo_real, 'costo_operativo': costo_operativo,
                                'costo_diferido_anterior': costo_diferido_anterior, 'consumo_por_cuenta': consumo_por_cuenta,
                                'mes_cierre': mes_cierre, 'anio_cierre': anio_cierre, 'unidad_cierre': unidad_cierre,
                                'es_consolidado': es_consolidado, 'pesos': pesos, 'nombres_meses': nombres_meses,
                                'anomalias': anomalias[[c_cod, 'Prefijo_Upper', df_k.columns[df_k.columns.str.contains('DOCUMENTO')][0], 'C_BASE', c_costo, 'DIF_M', 'VAR_P']]
                            }
                            
                            if 'huerfanos_df' in st.session_state: del st.session_state['huerfanos_df']
                            if 'pre_proceso' in st.session_state: del st.session_state['pre_proceso']
                            
                            st.rerun()

        else:
            # ==========================================
            # DESCARGA DE PARTIDAS Y GUARDADO
            # ==========================================
            mem = st.session_state['memoria_cierre']
            st.success(f"📦 **DATOS EN MEMORIA:** Cierre de {mem['unidad_cierre']} - Periodo: {meses_texto[mem['mes_cierre']]} {mem['anio_cierre']}")
            
            r1, r2, r3, r4, r5, r6 = st.columns(6)
            r1.metric("Inicial (+)", f"${mem['grp_ini_sum']:,.2f}")
            r2.metric("Compras (+)", f"${mem['grp_comp_sum']:,.2f}")
            r3.metric("Traslados Netos (+/-)", f"${mem['grp_tras_sum']:,.2f}")
            r4.metric("Final (-)", f"${mem['grp_fin_sum']:,.2f}")
            r5.metric("Diferido (-)", f"${mem['costo_dif_mes']:,.2f}")
            r6.metric("Real (=)", f"${mem['costo_real']:,.2f}")

            st.divider()
            if st.checkbox("📂 Ver Detalle de Consumo por Cuentas"):
                df_det_view = pd.DataFrame(list(mem['consumo_por_cuenta'].items()), columns=['Cuenta Contable', 'Consumo (Impacto)'])
                st.dataframe(df_det_view, use_container_width=True)

            st.subheader("🕵️ Auditoría de Costo de Venta (Doble Filtro)")
            if not mem['anomalias'].empty:
                st.error(f"🚨 Se detectaron {len(mem['anomalias'])} facturas con variaciones mayores al 1% y a $0.02.")
                res = mem['anomalias'].copy()
                res.columns = ['Código', 'Tipo', 'Documento', 'Costo Base', 'Costo Venta', 'Dif $', 'Var %']
                st.dataframe(res.style.format({'Costo Base':'${:.4f}', 'Costo Venta':'${:.4f}', 'Dif $':'${:.4f}', 'Var %':'{:.2%}'}), use_container_width=True)
                st.session_state['auditoria_aprobada'] = False
            else:
                st.success("✅ Validación del Kardex impecable: No hay desviaciones significativas.")
                st.session_state['auditoria_aprobada'] = True

            if not st.session_state.get('auditoria_aprobada', False):
                st.warning("⚠️ Debes justificar o corregir estas variaciones operativas antes de dar el cierre por bueno.")
                if st.button("🗑️ Descartar Memoria y Subir Archivos Nuevos"):
                    del st.session_state['memoria_cierre']
                    st.rerun()
                if st.button("⚠️ Forzar Aprobación (Bajo mi responsabilidad)"):
                    st.session_state['auditoria_aprobada'] = True
                    st.rerun()
            else:
                st.success("✅ Protocolo de Integridad Aprobado. Proceda con las descargas y el cierre.")
                
                def mostrar_descargas_logic(c_op, c_dif, c_ant, label_m, dict_consumo, total_op_base, key_suffix):
                    st.markdown(f"#### 📥 Partidas: **{label_m}**")
                    conc_v = f"RECONOCIMIENTO DE COSTO DE VENTA DE {mem['unidad_cierre']}, {label_m} {mem['anio_cierre']}."
                    f_v = [["410104", "", conc_v, round(c_op, 2), 0.00, ""]]
                    
                    for c, m in dict_consumo.items():
                        m_perc = m * (c_op / total_op_base) if total_op_base > 0 else 0
                        if abs(m_perc) > 0.01:
                            f_v.append([c, "", conc_v, 0.00, round(m_perc, 2), ""])
                    
                    conc_p = f"RECONOCIMIENTO DE COSTO DE LO VENDIDO EN PROCESO {mem['unidad_cierre']} {label_m} {mem['anio_cierre']}"
                    f_p = [["410104", "", conc_p, round(c_ant, 2), 0.00, ""], ["110602", "", conc_p, 0.00, round(c_ant, 2), ""]]
                    conc_d = f"DIFERIMIENTO DE COSTO EN PROCESO {mem['unidad_cierre']} {label_m} {mem['anio_cierre']}"
                    f_d = [["110602", "", conc_d, round(c_dif, 2), 0.00, ""], ["410104", "", conc_d, 0.00, round(c_dif, 2), ""]]

                    p1, p2, p3 = st.columns(3)
                    with p1: st.download_button(f"⬇️ P1 {label_m}", generar_excel_bytes(f_v), f"1_Costo_{label_m}.xlsx", key=f"gen_p1_{key_suffix}", use_container_width=True)
                    with p2: st.download_button(f"⬇️ P2 {label_m}", generar_excel_bytes(f_p), f"2_Ant_{label_m}.xlsx", key=f"gen_p2_{key_suffix}", use_container_width=True)
                    with p3: st.download_button(f"⬇️ P3 {label_m}", generar_excel_bytes(f_d), f"3_Dif_{label_m}.xlsx", key=f"gen_p3_{key_suffix}", use_container_width=True)

                if not mem['es_consolidado']:
                    mostrar_descargas_logic(mem['costo_operativo'], mem['costo_dif_mes'], mem['costo_diferido_anterior'], meses_texto[mem['mes_cierre']], mem['consumo_por_cuenta'], mem['costo_operativo'], "std")
                else:
                    for i in range(len(mem['pesos'])):
                        mostrar_descargas_logic(mem['costo_operativo']*mem['pesos'][i], mem['costo_dif_mes']*mem['pesos'][i], mem['costo_diferido_anterior']*mem['pesos'][i], mem['nombres_meses'][i], mem['consumo_por_cuenta'], mem['costo_operativo'], f"cons_{i}")

                if st.button("💾 Cerrar Periodo y Guardar Base", type="primary", use_container_width=True):
                    with st.spinner("Guardando en la nube (El guardado de detalle Kardex está deshabilitado temporalmente por migración)..."):
                        ws_res = conectar_hoja("Cierres_Costos")
                        fecha_hoy = date.today().strftime('%d/%m/%Y')
                        if ws_res:
                            ws_res.append_row([fecha_hoy, mem['mes_cierre'], mem['anio_cierre'], mem['unidad_cierre'], round(mem['grp_ini_sum'],2), round(mem['grp_comp_sum'],2), round(mem['grp_fin_sum'],2), round(mem['costo_diferido_anterior'],2), round(mem['costo_dif_mes'],2), round(mem['costo_real'],2)])
                            
                            del st.session_state['memoria_cierre']
                            st.session_state['auditoria_aprobada'] = False
                            st.cache_data.clear()
                            st.rerun()