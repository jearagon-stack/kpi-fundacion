import streamlit as st
import pandas as pd
from datetime import date
from utils import obtener_dataframe, conectar_hoja
import io
from validacion import ejecutar_auditoria_costos

def mostrar_modulo_costos():
    st.title("Contabilidad de Costos")

    # --- Pestañas Originales ---
    tab1, tab2, tab3 = st.tabs(["📝 Generar Cierre", "🚚 Partidas de Traslados", "🔍 Consultar Histórico"])

    # ==========================================
    # FUNCIONES DE APOYO (INTEGRAS)
    # ==========================================
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
            if sub in nombre_upper:
                return f"{unidad_base} {sub}"
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

    # ==========================================
    # PESTAÑA 1: GENERAR CIERRE
    # ==========================================
    with tab1:
        st.subheader("1. Cierre Contable")
        col_config1, col_config2 = st.columns([2, 1])
        with col_config1:
            tipo_cierre = st.radio("Tipo de proceso:", ["Mensual Estándar", "Consolidación Especial (Multi-mes)"], horizontal=True)

        col1, col2, col3 = st.columns(3)
        with col1: mes_cierre = st.selectbox("Mes a costear:", range(1, 13), index=date.today().month - 1)
        with col2: anio_cierre = st.number_input("Año:", min_value=2024, max_value=2030, value=2026)
        with col3: unidad_cierre = st.selectbox("Unidad:", ["CAFETERIA"])

        meses_texto = {1:"ENERO", 2:"FEBRERO", 3:"MARZO", 4:"ABRIL", 5:"MAYO", 6:"JUNIO", 7:"JULIO", 8:"AGOSTO", 9:"SEPTIEMBRE", 10:"OCTUBRE", 11:"NOVIEMBRE", 12:"DICIEMBRE"}
        lista_meses = list(meses_texto.values())
        
        es_consolidado = (tipo_cierre == "Consolidación Especial (Multi-mes)")
        pesos, nombres_meses = [], []

        if es_consolidado:
            # CORRECCIÓN: Ahora podés elegir desde 1 hasta el número del mes seleccionado
            num_meses = st.number_input("¿Dividir en cuántos meses?", min_value=1, max_value=mes_cierre, value=min(mes_cierre, 3))
            
            st.markdown(f"**Distribución de costos para el periodo de {num_meses} meses finalizando en {meses_texto[mes_cierre]}:**")
            
            # Usamos columnas dinámicas. Si son muchos meses, se ajustan automáticamente.
            cols_dist = st.columns(num_meses) if num_meses <= 6 else [st.container() for _ in range(num_meses)]
            
            for i in range(num_meses):
                idx_sugerido = mes_cierre - (num_meses - 1) + i
                if idx_sugerido < 1: idx_sugerido += 12
                
                # Si hay muchas columnas, las ponemos en filas para que no se vea amontonado
                target = cols_dist[i] if num_meses <= 6 else st
                
                with target:
                    n = st.selectbox(f"Mes {i+1}:", options=lista_meses, index=idx_sugerido-1, key=f"n_gen_{i}")
                    p = st.number_input(f"% Venta {n}", min_value=0.0, max_value=100.0, value=100.0/num_meses, key=f"p_gen_{i}")
                    pesos.append(p / 100); nombres_meses.append(n)

        st.divider()
        # --- Lógica de Ingresos ---
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
            f_t = (pd.to_numeric(df_hist_tras['Mes'], errors='coerce') == mes_cierre) & (pd.to_numeric(df_hist_tras['Año'], errors='coerce') == anio_cierre) & (df_hist_tras['Destino'] == unidad_cierre)
            traslados_mes = pd.to_numeric(df_hist_tras[f_t]['Monto'], errors='coerce').sum()

        porcentaje_subsidio = (subsidio_mes / ventas_mes) if ventas_mes > 0 else 0.0
        st.info(f"📊 Ingresos: Ventas ${ventas_mes:,.2f} | Subsidio ${subsidio_mes:,.2f} | Traslados Recibidos: ${traslados_mes:,.2f}")

        costo_diferido_anterior = st.number_input("Costo Diferido de Arrastre (110602):", min_value=0.0, value=0.0)
        
        col_u1, col_u2, col_u3 = st.columns(3)
        with col_u1: arch_ini = st.file_uploader("1. Inv. Inicial", type=["xlsx"], accept_multiple_files=True)
        with col_u2: arch_com = st.file_uploader("2. Compras", type=["xlsx"], accept_multiple_files=True)
        with col_u3: arch_fin = st.file_uploader("3. Inv. Final", type=["xlsx"], accept_multiple_files=True)

        if arch_ini and arch_com and arch_fin:
            try:
                df_inicial = consolidar(arch_ini); df_compras = consolidar(arch_com); df_final = consolidar(arch_fin)
                df_dic = obtener_dataframe("Categorias_Costos")
                def limpiar_cod(s): return s.astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
                for df in [df_dic, df_inicial, df_compras, df_final]: df['Codigo'] = limpiar_cod(df['Codigo'])
                
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
                for cta in todas_cuentas:
                    val = grp_ini.get(cta, 0) + grp_comp.get(cta, 0) + grp_tras.get(cta, 0) - grp_fin.get(cta, 0)
                    if val != 0: consumo_por_cuenta[cta] = val; costo_operativo += val
                
                costo_dif_mes = costo_operativo * porcentaje_subsidio
                costo_real = costo_operativo - costo_dif_mes + costo_diferido_anterior

                r1, r2, r3, r4, r5 = st.columns(5)
                r1.metric("Inicial", f"${grp_ini.sum():,.2f}"); r2.metric("Compras", f"${grp_comp.sum():,.2f}")
                r3.metric("Final", f"${grp_fin.sum():,.2f}"); r4.metric("Diferido", f"${costo_dif_mes:,.2f}"); r5.metric("Real", f"${costo_real:,.2f}")

                es_apto = ejecutar_auditoria_costos(df_ini_m, df_com_m, df_fin_m, consumo_por_cuenta, ventas_mes, costo_real, mes_cierre, anio_cierre, unidad_cierre)

                if es_apto:
                    def mostrar_descargas_logic(c_op, c_dif, c_ant, label_m, dict_consumo, total_op_base, key_suffix):
                        st.markdown(f"#### 📥 Partidas: **{label_m}**")
                        conc_v = f"RECONOCIMIENTO DE COSTO DE VENTA DE CAFETERIA, {label_m} {anio_cierre}."
                        f_v = [["410104", "", conc_v, round(c_op, 2), 0.00, ""]]
                        for c, m in dict_consumo.items():
                            m_perc = m * (c_op / total_op_base) if total_op_base > 0 else 0
                            if m_perc > 0: f_v.append([str(c).replace(".0",""), "", conc_v, 0.00, round(m_perc, 2), ""])
                        
                        conc_p = f"RECONOCIMIENTO DE COSTO DE LO VENDIDO EN PROCESO CAFETERIA {label_m} {anio_cierre}"
                        f_p = [["410104", "", conc_p, round(c_ant, 2), 0.00, ""], ["110602", "", conc_p, 0.00, round(c_ant, 2), ""]]
                        conc_d = f"DIFERIMIENTO DE COSTO EN PROCESO CAFETERIA {label_m} {anio_cierre}"
                        f_d = [["110602", "", conc_d, round(c_dif, 2), 0.00, ""], ["410104", "", conc_d, 0.00, round(c_dif, 2), ""]]

                        p1, p2, p3 = st.columns(3)
                        with p1: st.download_button(f"⬇️ P1 {label_m}", generar_excel_bytes(f_v), f"1_Costo_{label_m}.xlsx", key=f"gen_p1_{key_suffix}", use_container_width=True)
                        with p2: st.download_button(f"⬇️ P2 {label_m}", generar_excel_bytes(f_p), f"2_Ant_{label_m}.xlsx", key=f"gen_p2_{key_suffix}", use_container_width=True)
                        with p3: st.download_button(f"⬇️ P3 {label_m}", generar_excel_bytes(f_d), f"3_Dif_{label_m}.xlsx", key=f"gen_p3_{key_suffix}", use_container_width=True)

                    if not es_consolidado:
                        mostrar_descargas_logic(costo_operativo, costo_dif_mes, costo_diferido_anterior, meses_texto[mes_cierre], consumo_por_cuenta, costo_operativo, "std")
                    else:
                        for i in range(len(pesos)):
                            mostrar_descargas_logic(costo_operativo*pesos[i], costo_dif_mes*pesos[i], costo_diferido_anterior*pesos[i], nombres_meses[i], consumo_por_cuenta, costo_operativo, f"cons_{i}")

                    if st.button("💾 Guardar Cierre", type="primary", use_container_width=True):
                        with st.spinner("Guardando..."):
                            st.cache_data.clear()
                            ws_res = conectar_hoja("Cierres_Costos"); ws_det = conectar_hoja("Detalle_Cuentas")
                            fecha_hoy = date.today().strftime('%d/%m/%Y')
                            if ws_res and ws_det:
                                ws_res.append_row([fecha_hoy, mes_cierre, anio_cierre, unidad_cierre, round(grp_ini.sum(),2), round(grp_comp.sum(),2), round(grp_fin.sum(),2), round(costo_diferido_anterior,2), round(costo_dif_mes,2), round(costo_real,2)])
                                # --- Lógica de Detalle (Audit Trail) ---
                                df_det_c = pd.concat([df_ini_m[['Codigo','Cuenta_Contable','Valor','ORIGEN_ARCHIVO']].rename(columns={'Valor':'Inicial'}),
                                                   df_com_m[['Codigo','Cuenta_Contable','Valor','ORIGEN_ARCHIVO']].rename(columns={'Valor':'Compra'}),
                                                   df_fin_m[['Codigo','Cuenta_Contable','Valor','ORIGEN_ARCHIVO']].rename(columns={'Valor':'Final'})]).fillna(0)
                                df_det_c = df_det_c.groupby(['Codigo','Cuenta_Contable','ORIGEN_ARCHIVO']).sum().reset_index()
                                df_det_c['Consumo'] = df_det_c['Inicial'] + df_det_c['Compra'] - df_det_c['Final']
                                filas_g = []
                                for _, r in df_det_c.iterrows():
                                    if r['Inicial']!=0 or r['Compra']!=0 or r['Final']!=0:
                                        u_r = extraer_subunidad(r['ORIGEN_ARCHIVO'], unidad_cierre)
                                        filas_g.append([fecha_hoy, mes_cierre, anio_cierre, u_r, str(r['Cuenta_Contable']), round(r['Inicial'],2), round(r['Compra'],2), round(r['Final'],2), round(r['Consumo'],2), r['Codigo'], r['ORIGEN_ARCHIVO']])
                                if filas_g: ws_det.append_rows(filas_g)
                                st.success("✅ Conformidad guardada.")
            except Exception as e: st.error(f"Error: {e}")

    # ==========================================
    # PESTAÑA 2: TRASLADOS
    # ==========================================
    with tab2:
        st.subheader("🚚 Registro de Traslados Nexus")
        ct1, ct2 = st.columns(2)
        with ct1:
            m_t = st.selectbox("Mes:", range(1, 13), index=date.today().month - 1, key="mt_reg")
            origen_t = st.text_input("Bodega Origen:", "ABASTECIMIENTO").upper()
        with ct2:
            a_t = st.number_input("Año:", min_value=2024, value=2026, key="at_reg")
            destino_t = st.text_input("Bodega Destino:", "CAFETERIA").upper()
        arch_t = st.file_uploader("Reporte Nexus (I, N, AA)", type=["xlsx"], key="atf_reg")
        if arch_t:
            df_t_raw = pd.read_excel(arch_t, usecols="I,N,AA", names=['Codigo', 'Monto', 'Categoria'], skiprows=1)
            df_t_raw['Codigo'] = df_t_raw['Codigo'].astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
            df_dic_t = obtener_dataframe("Categorias_Costos")
            df_dic_t['Codigo'] = df_dic_t['Codigo'].astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
            df_f_t = pd.merge(df_t_raw, df_dic_t[['Codigo', 'Cuenta_Contable']], on='Codigo', how='left')
            st.dataframe(df_f_t, use_container_width=True)
            if st.button("💾 Guardar Traslados en Historial"):
                ws_t = conectar_hoja("Historico_Traslados")
                fecha_h = date.today().strftime('%d/%m/%Y')
                filas_t = [[fecha_h, m_t, a_t, origen_t, destino_t, str(r['Cuenta_Contable']), round(r['Monto'],2), r['Codigo'], "Traslado Nexus"] for _, r in df_f_t.iterrows()]
                ws_t.append_rows(filas_t)
                st.success("✅ Traslados registrados satisfactoriamente.")

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
                        if m > 0: fv_h.append([str(c).replace(".0",""), "", cv_h, 0, round(m, 2), ""])
                    st.download_button(f"⬇️ Bajar P1 {m_f}", generar_excel_bytes(fv_h), f"H_P1_{m_f}.xlsx", key="hp1_std")