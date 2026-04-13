import streamlit as st
import pandas as pd
from datetime import date
from utils import obtener_dataframe, conectar_hoja
import io

def mostrar_modulo_costos():
    st.title("Contabilidad de Costos")

    tab1, tab2, tab3 = st.tabs(["📝 Generar Cierre", "🚚 Partidas de Traslados", "🔍 Consultar Histórico"])

    # ==========================================
    # FUNCIONES DE APOYO Y MAPEO DE UNIDADES
    # ==========================================
    mapa_subunidades = {
        "CAFETERIA": ["TERRAZA", "CENTRO SOHO", "CAFETERIA CENTRAL", "CAFETERIA", "CAFETERIA ABASTECIMIENTO"], # <-- Aquí agregué la bodega
        "DESPENSA":  ["DESPENSA"]
    }

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

    def cargar_y_marcar(archivo):
        df = pd.read_excel(archivo)
        cols_upper = df.columns.astype(str).str.upper().str.replace(' ', '').str.replace('.', '')
        if not (any('COD' in c and 'PROD' in c for c in cols_upper) or any('IDPRODUCTO' in c for c in cols_upper)):
            for i in range(min(15, len(df))):
                row_str = df.iloc[i].astype(str).str.upper().str.replace(' ', '').str.replace('.', '')
                if any('COD' in val and 'PROD' in val for val in row_str) or any('IDPRODUCTO' in val for val in row_str):
                    df.columns = df.iloc[i].astype(str).str.strip()
                    df = df.iloc[i+1:].reset_index(drop=True)
                    break
        col_rename = {}
        for col in df.columns:
            c_norm = str(col).upper().replace(' ', '').replace('.', '')
            if ('COD' in c_norm and 'PROD' in c_norm) or 'IDPRODUCTO' in c_norm:
                col_rename[col] = 'Codigo'
        df.rename(columns=col_rename, inplace=True)
        df['ORIGEN_ARCHIVO'] = archivo.name
        return df

    def limpiar_nativos_nexus(df):
        col_cat = next((c for c in df.columns if 'CATEGORIA' in str(c).upper()), None)
        if col_cat:
            df[col_cat] = df[col_cat].astype(str).str.upper().str.strip()
            df = df[df[col_cat] != 'SERVICIO']
        return df

    def consolidar(lista):
        return pd.concat([limpiar_nativos_nexus(cargar_y_marcar(a)) for a in lista], ignore_index=True)

    def get_num(df, keys):
        for k in keys:
            for col in df.columns:
                c_norm = str(col).upper().replace(' ', '').replace('.', '')
                if all(p in c_norm for p in k): return pd.to_numeric(df[col], errors='coerce').fillna(0)
        return pd.Series(0, index=df.index)

    meses_texto = {1:"ENERO", 2:"FEBRERO", 3:"MARZO", 4:"ABRIL", 5:"MAYO", 6:"JUNIO", 7:"JULIO", 8:"AGOSTO", 9:"SEPTIEMBRE", 10:"OCTUBRE", 11:"NOVIEMBRE", 12:"DICIEMBRE"}
    lista_meses = list(meses_texto.values())

    # ==========================================
    # PESTAÑA 1: GENERAR CIERRE
    # ==========================================
    with tab1:
        st.subheader("1. Cierre Contable")

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

            df_hist_tras = obtener_dataframe("Historico_Traslados")
            traslados_mes = 0.0
            if not df_hist_tras.empty:
                destinos_a_buscar = mapa_subunidades.get(unidad_cierre, [unidad_cierre])
                f_t = (pd.to_numeric(df_hist_tras['Mes'], errors='coerce') == mes_cierre) & \
                      (pd.to_numeric(df_hist_tras['Año'], errors='coerce') == anio_cierre) & \
                      (df_hist_tras['Destino'].isin(destinos_a_buscar))
                traslados_mes = pd.to_numeric(df_hist_tras[f_t]['Monto'], errors='coerce').sum()

            porcentaje_subsidio = (subsidio_mes / ventas_mes) if ventas_mes > 0 else 0.0
            st.info(f"📊 Ingresos Totales Periodo: Ventas ${ventas_mes:,.2f} | Subsidio ${subsidio_mes:,.2f} | Traslados Recibidos: ${traslados_mes:,.2f}")

            costo_diferido_anterior = st.number_input("Costo Diferido de Arrastre (110602):", min_value=0.0, value=0.0)
            
            col_u1, col_u2, col_u3 = st.columns(3)
            with col_u1: arch_ini = st.file_uploader("1. Inv. Inicial", type=["xlsx"], accept_multiple_files=True)
            with col_u2: arch_com = st.file_uploader("2. Compras", type=["xlsx"], accept_multiple_files=True)
            with col_u3: arch_fin = st.file_uploader("3. Inv. Final", type=["xlsx"], accept_multiple_files=True)

            if arch_ini and arch_com and arch_fin:
                if st.button("⚙️ Procesar Archivos y Guardar en Memoria", type="primary", use_container_width=True):
                    with st.spinner("Procesando miles de filas..."):
                        try:
                            df_inicial = consolidar(arch_ini); df_compras = consolidar(arch_com); df_final = consolidar(arch_fin)
                            df_dic = obtener_dataframe("Categorias_Costos")
                            def limpiar_cod(s): return s.astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
                            for df in [df_dic, df_inicial, df_compras, df_final]: df['Codigo'] = limpiar_cod(df['Codigo'])
                            
                            basura = ['G222', 'G231', '21455979']
                            df_inicial = df_inicial[~df_inicial['Codigo'].isin(basura)]
                            df_compras = df_compras[~df_compras['Codigo'].isin(basura)]
                            df_final = df_final[~df_final['Codigo'].isin(basura)]

                            df_ini_m = pd.merge(df_inicial, df_dic, on='Codigo', how='left')
                            df_com_m = pd.merge(df_compras, df_dic, on='Codigo', how='left')
                            df_fin_m = pd.merge(df_final, df_dic, on='Codigo', how='left')

                            df_ini_m['Valor'] = get_num(df_ini_m, [['EXISTENCIAS'], ['SALDO']]) * get_num(df_ini_m, [['COSTO', 'U']])
                            df_fin_m['Valor'] = get_num(df_fin_m, [['EXISTENCIAS'], ['SALDO']]) * get_num(df_fin_m, [['COSTO', 'U']])
                            df_com_m['Valor'] = get_num(df_com_m, [['TOTAL'], ['MONTO']])

                            grp_ini = df_ini_m.groupby('Cuenta_Contable')['Valor'].sum()
                            grp_comp = df_com_m.groupby('Cuenta_Contable')['Valor'].sum()
                            grp_fin = df_fin_m.groupby('Cuenta_Contable')['Valor'].sum()
                            
                            grp_tras = df_hist_tras[f_t].groupby('Cuenta_Contable')['Monto'].sum() if not df_hist_tras.empty else pd.Series(0.0)

                            todas_cuentas = set(grp_ini.index).union(grp_comp.index).union(grp_fin.index).union(grp_tras.index)
                            consumo_por_cuenta = {}; costo_operativo = 0.0
                            
                            cuentas_invalidas = ["", "NAN", "NAT", "NO APLICA", "0", "0.0", "NONE"]
                            for cta in todas_cuentas:
                                if pd.isna(cta): continue
                                cta_str = str(cta).strip().upper().replace(".0", "")
                                val = grp_ini.get(cta, 0) + grp_comp.get(cta, 0) + grp_tras.get(cta, 0) - grp_fin.get(cta, 0)
                                if val != 0 and cta_str not in cuentas_invalidas: 
                                    consumo_por_cuenta[cta_str] = val
                                    costo_operativo += val
                            
                            costo_dif_mes = costo_operativo * porcentaje_subsidio
                            costo_real = costo_operativo - costo_dif_mes + costo_diferido_anterior

                            st.session_state['memoria_cierre'] = {
                                'df_ini_m': df_ini_m, 'df_com_m': df_com_m, 'df_fin_m': df_fin_m,
                                'grp_ini_sum': grp_ini.sum(), 'grp_comp_sum': grp_comp.sum(), 'grp_fin_sum': grp_fin.sum(),
                                'costo_dif_mes': costo_dif_mes, 'costo_real': costo_real, 'costo_operativo': costo_operativo,
                                'costo_diferido_anterior': costo_diferido_anterior, 'consumo_por_cuenta': consumo_por_cuenta,
                                'mes_cierre': mes_cierre, 'anio_cierre': anio_cierre, 'unidad_cierre': unidad_cierre,
                                'es_consolidado': es_consolidado, 'pesos': pesos, 'nombres_meses': nombres_meses
                            }
                            st.session_state['datos_auditoria'] = {'consumo': consumo_por_cuenta, 'ventas': ventas_mes, 'costo_real': costo_real}
                            st.rerun() 
                        except Exception as e: st.error(f"Error procesando: {e}")

        else:
            mem = st.session_state['memoria_cierre']
            st.success(f"📦 **DATOS EN MEMORIA:** Cierre de {mem['unidad_cierre']} - Periodo: {meses_texto[mem['mes_cierre']]} {mem['anio_cierre']}")
            
            r1, r2, r3, r4, r5 = st.columns(5)
            r1.metric("Inicial", f"${mem['grp_ini_sum']:,.2f}")
            r2.metric("Compras", f"${mem['grp_comp_sum']:,.2f}")
            r3.metric("Final", f"${mem['grp_fin_sum']:,.2f}")
            r4.metric("Diferido", f"${mem['costo_dif_mes']:,.2f}")
            r5.metric("Real", f"${mem['costo_real']:,.2f}")

            st.divider()
            if st.checkbox("📂 Ver Detalle de Movimientos / Cuentas"):
                df_det_view = pd.DataFrame(list(mem['consumo_por_cuenta'].items()), columns=['Cuenta Contable', 'Consumo (Impacto)'])
                st.dataframe(df_det_view, use_container_width=True)

            if not st.session_state.get('auditoria_aprobada', False):
                st.warning("⚠️ Ve al módulo **'VALIDACIÓN DE COSTOS'** en el menú izquierdo para auditar y aprobar este periodo.")
                if st.button("🗑️ Descartar Memoria y Subir Archivos Nuevos"):
                    del st.session_state['memoria_cierre']
                    if 'datos_auditoria' in st.session_state: del st.session_state['datos_auditoria']
                    st.rerun()
            else:
                st.success("✅ Protocolo de Integridad Aprobado. Proceda con las descargas y el cierre.")
                
                def mostrar_descargas_logic(c_op, c_dif, c_ant, label_m, dict_consumo, total_op_base, key_suffix):
                    st.markdown(f"#### 📥 Partidas: **{label_m}**")
                    conc_v = f"RECONOCIMIENTO DE COSTO DE VENTA DE {mem['unidad_cierre']}, {label_m} {mem['anio_cierre']}."
                    f_v = [["410104", "", conc_v, round(c_op, 2), 0.00, ""]]
                    
                    for c, m in dict_consumo.items():
                        m_perc = m * (c_op / total_op_base) if total_op_base > 0 else 0
                        if m_perc > 0.01:
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
                    with st.spinner("Guardando en la nube..."):
                        ws_res = conectar_hoja("Cierres_Costos"); ws_det = conectar_hoja("Detalle_Cuentas")
                        fecha_hoy = date.today().strftime('%d/%m/%Y')
                        if ws_res and ws_det:
                            ws_res.append_row([fecha_hoy, mem['mes_cierre'], mem['anio_cierre'], mem['unidad_cierre'], round(mem['grp_ini_sum'],2), round(mem['grp_comp_sum'],2), round(mem['grp_fin_sum'],2), round(mem['costo_diferido_anterior'],2), round(mem['costo_dif_mes'],2), round(mem['costo_real'],2)])
                            df_det_c = pd.concat([mem['df_ini_m'][['Codigo','Cuenta_Contable','Valor','ORIGEN_ARCHIVO']].rename(columns={'Valor':'Inicial'}),
                                               mem['df_com_m'][['Codigo','Cuenta_Contable','Valor','ORIGEN_ARCHIVO']].rename(columns={'Valor':'Compra'}),
                                               mem['df_fin_m'][['Codigo','Cuenta_Contable','Valor','ORIGEN_ARCHIVO']].rename(columns={'Valor':'Final'})]).fillna(0)
                            df_det_c = df_det_c.groupby(['Codigo','Cuenta_Contable','ORIGEN_ARCHIVO']).sum().reset_index()
                            df_det_c['Consumo'] = df_det_c['Inicial'] + df_det_c['Compra'] - df_det_c['Final']
                            filas_g = []
                            for _, r in df_det_c.iterrows():
                                if r['Inicial']!=0 or r['Compra']!=0 or r['Final']!=0:
                                    u_r = extraer_subunidad(r['ORIGEN_ARCHIVO'], mem['unidad_cierre'])
                                    filas_g.append([fecha_hoy, mem['mes_cierre'], mem['anio_cierre'], u_r, str(r['Cuenta_Contable']), round(r['Inicial'],2), round(r['Compra'],2), round(r['Final'],2), round(r['Consumo'],2), r['Codigo'], r['ORIGEN_ARCHIVO']])
                            if filas_g: ws_det.append_rows(filas_g)
                            
                            del st.session_state['memoria_cierre']
                            del st.session_state['datos_auditoria']
                            st.session_state['auditoria_aprobada'] = False
                            st.cache_data.clear()
                            st.rerun()

    # ==========================================
    # PESTAÑA 2: TRASLADOS
    # ==========================================
    with tab2:
        st.subheader("🚚 Registro Automático de Traslados Nexus")
        ct1, ct2, ct3 = st.columns(3)
        with ct1:
            m_t = st.selectbox("Mes:", range(1, 13), index=date.today().month - 1, key="mt_reg")
        with ct2:
            a_t = st.number_input("Año:", min_value=2024, value=2026, key="at_reg")
        with ct3:
            unidad_t = st.selectbox("Unidad que recibe (Destino):", ["CAFETERIA", "DESPENSA"], key="uni_reg")
            
        st.info(f"💡 El sistema extraerá solo los traslados que apliquen para el módulo **{unidad_t}**. Filtraremos montos cero y servicios.")
        
        arch_t = st.file_uploader("Reporte Nexus (I:Cod, K:Cant, N:Monto, AA:Cat, AB:Origen, AC:Destino)", type=["xlsx"], key="atf_reg")
        if arch_t:
            try:
                df_t_raw = pd.read_excel(arch_t, usecols="I,K,N,AA,AB,AC", names=['Codigo', 'Cantidad', 'Monto', 'Categoria', 'Origen', 'Destino'], skiprows=1)
                df_t_raw['Codigo'] = df_t_raw['Codigo'].astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
                df_t_raw['Origen'] = df_t_raw['Origen'].astype(str).str.strip().str.upper()
                df_t_raw['Destino'] = df_t_raw['Destino'].astype(str).str.strip().str.upper()
                df_t_raw['Categoria'] = df_t_raw['Categoria'].astype(str).str.strip().str.upper()
                df_t_raw['Monto'] = pd.to_numeric(df_t_raw['Monto'], errors='coerce').fillna(0)
                df_t_raw['Cantidad'] = pd.to_numeric(df_t_raw['Cantidad'], errors='coerce').fillna(0)
                
                destinos_validos = mapa_subunidades.get(unidad_t, [unidad_t])
                
                df_validos = df_t_raw[(df_t_raw['Destino'].isin(destinos_validos)) & (df_t_raw['Monto'] > 0) & (df_t_raw['Categoria'] != 'SERVICIO')]
                
                if df_validos.empty:
                    st.warning(f"⚠️ No hay traslados válidos (monto > 0, no servicios) hacia ninguna bodega de {unidad_t}.")
                else:
                    df_dic_t = obtener_dataframe("Categorias_Costos")
                    df_dic_t['Codigo'] = df_dic_t['Codigo'].astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
                    df_f_t = pd.merge(df_validos, df_dic_t[['Codigo', 'Cuenta_Contable']], on='Codigo', how='left')
                    
                    df_f_t = df_f_t.dropna(subset=['Cuenta_Contable'])
                    cuentas_invalidas = ["", "NAN", "NAT", "NO APLICA", "0", "0.0", "NONE"]
                    df_f_t = df_f_t[~df_f_t['Cuenta_Contable'].astype(str).str.strip().str.upper().isin(cuentas_invalidas)]
                    
                    if df_f_t.empty:
                        st.warning("⚠️ Los movimientos detectados no tienen cuentas contables asignadas válidas. Asigne las cuentas en el Maestro para poder generar partidas.")
                    else:
                        st.success(f"✅ Se identificaron {len(df_f_t)} movimientos listos para registro en {unidad_t}.")
                        st.dataframe(df_f_t, use_container_width=True)
                        
                        st.markdown("#### 📥 Partidas Generadas Automáticamente:")
                        grupos = df_f_t.groupby(['Origen', 'Destino'])
                        num_grupos = len(grupos)
                        
                        if num_grupos > 0:
                            cols_descarga = st.columns(min(num_grupos, 4))
                            
                            for idx, ((origen_val, destino_val), df_grupo) in enumerate(grupos):
                                grp_t = df_grupo.groupby('Cuenta_Contable')['Monto'].sum()
                                
                                conc_t = f"TRASLADO DE {origen_val} A {destino_val}, {meses_texto[m_t]} {a_t}"
                                f_p_t = []
                                for c, m in grp_t.items():
                                    if m > 0.01: f_p_t.append([str(c).replace(".0",""), "", conc_t, round(m, 2), 0, ""])
                                for c, m in grp_t.items():
                                    if m > 0.01: f_p_t.append([str(c).replace(".0",""), "", conc_t, 0, round(m, 2), ""])
                                
                                with cols_descarga[idx % min(num_grupos, 4)]:
                                    st.download_button(f"⬇️ {destino_val} (desde {origen_val})", generar_excel_bytes(f_p_t), f"Traslado_{origen_val}_A_{destino_val}_{meses_texto[m_t]}.xlsx", key=f"dl_t_{idx}")

                        st.divider()
                        if st.button("💾 Guardar Traslados en Historial"):
                            with st.spinner("Validando integridad de datos..."):
                                df_historico = obtener_dataframe("Historico_Traslados")
                                
                                def crear_llave(row):
                                    return f"{int(float(row['Mes']))}|{int(float(row['Año']))}|{str(row['Origen']).strip()}|{str(row['Destino']).strip()}|{str(row['Código']).strip()}"
                                
                                if not df_historico.empty:
                                    df_historico['llave'] = df_historico.apply(crear_llave, axis=1)
                                    llaves_nuevas = [f"{m_t}|{a_t}|{str(r['Origen']).strip()}|{str(r['Destino']).strip()}|{str(r['Codigo']).strip()}" for _, r in df_f_t.iterrows()]
                                    duplicados = [l for l in llaves_nuevas if l in df_historico['llave'].values]
                                    
                                    if duplicados:
                                        st.error(f"❌ **MOVIMIENTOS DUPLICADOS DETECTADOS:** Ya existen registros para este periodo ({m_t}/{a_t}) con los mismos destinos y productos.")
                                        st.warning("Verifique el archivo para no duplicar datos.")
                                        with st.expander("Ver detalle de llaves duplicadas"):
                                            st.write(duplicados)
                                    else:
                                        ws_t = conectar_hoja("Historico_Traslados")
                                        fecha_h = date.today().strftime('%d/%m/%Y')
                                        filas_t = [[fecha_h, m_t, a_t, r['Origen'], r['Destino'], str(r['Cuenta_Contable']), round(r['Monto'],2), r['Codigo'], "Traslado Nexus", r['Cantidad']] for _, r in df_f_t.iterrows()]
                                        ws_t.append_rows(filas_t)
                                        st.success("✅ Traslados guardados con éxito.")
                                else:
                                    ws_t = conectar_hoja("Historico_Traslados")
                                    fecha_h = date.today().strftime('%d/%m/%Y')
                                    filas_t = [[fecha_h, m_t, a_t, r['Origen'], r['Destino'], str(r['Cuenta_Contable']), round(r['Monto'],2), r['Codigo'], "Traslado Nexus", r['Cantidad']] for _, r in df_f_t.iterrows()]
                                    ws_t.append_rows(filas_t)
                                    st.success("✅ Base inicializada. Traslados guardados con éxito.")

            except Exception as e: st.error(f"Error: {e}")

    # ==========================================
    # PESTAÑA 3: CONSULTA HISTORIAL
    # ==========================================
    with tab3:
        st.subheader("🔍 Consulta de Historial Operativo")
        if st.button("🔄 Actualizar Base", key="update_hist"): st.cache_data.clear()
        df_resumen = obtener_dataframe("Cierres_Costos"); df_detalle = obtener_dataframe("Detalle_Cuentas")
        if not df_resumen.empty:
            df_resumen['Periodo'] = df_resumen['Mes'].astype(str).str.replace('.0','') + "/" + df_resumen['Año'].astype(str).str.replace('.0','') + " - " + df_resumen['Unidad']
            per_sel = st.selectbox("Seleccione Cierre:", df_resumen['Periodo'].unique().tolist(), index=len(df_resumen['Periodo'].unique())-1)
            if per_sel:
                fila = df_resumen[df_resumen['Periodo'] == per_sel].iloc[-1]
                m_f, a_f = str(fila['Mes']).strip(), str(fila['Año']).strip()
                v_ini, v_com, v_fin = float(fila.iloc[4]), float(fila.iloc[5]), float(fila.iloc[6])
                v_ant, v_dif, v_real = float(fila.iloc[7]), float(fila.iloc[8]), float(fila.iloc[9])
                c1, c2, c3, c4, c5 = st.columns(5)
                c1.metric("Inicial", f"${v_ini:,.2f}"); c2.metric("Compras", f"${v_com:,.2f}")
                c3.metric("Final", f"${v_fin:,.2f}"); c4.metric("Diferido", f"${v_dif:,.2f}"); c5.metric("Real", f"${v_real:,.2f}")
                df_det_h = df_detalle[(df_detalle['Mes'].astype(str).str.replace('.0','') == m_f) & (df_detalle['Año'].astype(str).str.replace('.0','') == a_f)]
                if st.checkbox("📂 Ver Detalle de Movimientos", key="chk_det"):
                    st.dataframe(df_det_h, use_container_width=True)
                st.divider(); st.markdown("#### 📥 Reconstruir Partidas Nexus")
                tipo_h = st.radio("Formato:", ["Cierre Estándar", "Cierre Consolidado"], key="th")
                consumo_h = df_det_h.groupby('Cuenta')['Consumo'].sum().to_dict()
                op_h = sum(consumo_h.values())
                if tipo_h == "Cierre Estándar":
                    cv_h = f"RECONOCIMIENTO DE COSTO DE VENTA DE CAFETERIA, {m_f}/{a_f}."
                    fv_h = [["410104", "", cv_h, round(op_h, 2), 0, ""]]
                    for c, m in consumo_h.items():
                        cta_str = str(c).replace(".0","").strip()
                        if cta_str != "" and cta_str.lower() not in ["nan", "nat", "no aplica"]:
                            if m > 0.01: fv_h.append([cta_str, "", cv_h, 0, round(m, 2), ""])
                    st.download_button(f"⬇️ Bajar P1 {m_f}", generar_excel_bytes(fv_h), f"H_P1_{m_f}.xlsx", key="hp1_std")