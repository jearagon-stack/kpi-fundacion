import streamlit as st
import pandas as pd
from datetime import date
from utils import obtener_dataframe, conectar_hoja
import io
from validacion import ejecutar_auditoria_costos

def mostrar_modulo_costos():
    st.title("Contabilidad de Costos")

    # --- 1. ORDEN DE PESTAÑAS: Histórico es la 3 ---
    tab1, tab2, tab3 = st.tabs(["📝 Generar Cierre", "🚚 Partidas de Traslados", "🔍 Consultar Histórico"])

    # ==========================================
    # FUNCIONES DE APOYO (RESTAURADAS AL 100%)
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
        with col2: anio_cierre = st.number_input("Año:", min_value=2024, max_value=2030, value=date.today().year)
        with col3: unidad_cierre = st.selectbox("Unidad:", ["CAFETERIA"])

        meses_texto = {1:"ENERO", 2:"FEBRERO", 3:"MARZO", 4:"ABRIL", 5:"MAYO", 6:"JUNIO", 7:"JULIO", 8:"AGOSTO", 9:"SEPTIEMBRE", 10:"OCTUBRE", 11:"NOVIEMBRE", 12:"DICIEMBRE"}
        mes_txt = meses_texto[mes_cierre]

        es_consolidado = (tipo_cierre == "Consolidación Especial (Multi-mes)")
        pesos, nombres_meses = [], []

        if es_consolidado:
            st.warning("⚠️ Modo Consolidación: El costo se distribuirá en las partidas de cada mes.")
            num_meses = st.radio("¿Dividir en cuántos meses?", [2, 3], horizontal=True)
            cols_dist = st.columns(num_meses)
            for i in range(num_meses):
                with cols_dist[i]:
                    n = st.text_input(f"Mes {i+1}:", meses_texto.get(mes_cierre-num_meses+i+1, "MES"), key=f"n_gen_{i}").upper()
                    p = st.number_input(f"% Venta {n}", min_value=0.0, max_value=100.0, value=100.0/num_meses, key=f"p_gen_{i}")
                    pesos.append(p / 100); nombres_meses.append(n)

        st.divider()
        df_ventas = obtener_dataframe("Historico_Ventas")
        ventas_mes = subsidio_mes = 0.0
        if not df_ventas.empty:
            df_ventas['Fecha_DT'] = pd.to_datetime(df_ventas['Fecha'], format='%d/%m/%Y', errors='coerce')
            filtro = (df_ventas['Fecha_DT'].dt.year == anio_cierre) & (df_ventas['Unidad'] == unidad_cierre)
            if es_consolidado:
                filtro &= df_ventas['Fecha_DT'].dt.month.isin(range(mes_cierre-len(pesos)+1, mes_cierre+1))
            else:
                filtro &= (df_ventas['Fecha_DT'].dt.month == mes_cierre)
            ventas_mes = pd.to_numeric(df_ventas[filtro]['Venta_Real'], errors='coerce').sum()
            subsidio_mes = pd.to_numeric(df_ventas[filtro]['Subsidio_UCA'], errors='coerce').sum()

        porcentaje_subsidio = (subsidio_mes / ventas_mes) if ventas_mes > 0 else 0.0
        st.info(f"📊 Información de Ingresos: Ventas ${ventas_mes:,.2f} | Subsidio ${subsidio_mes:,.2f} | % Subsidio: {porcentaje_subsidio:.2%}")

        costo_diferido_anterior = st.number_input("Costo Diferido de Arrastre (110602):", min_value=0.0, value=0.0)
        
        col_u1, col_u2, col_u3 = st.columns(3)
        with col_u1: arch_ini = st.file_uploader("1. Inv. Inicial", type=["xlsx", "xls"], accept_multiple_files=True)
        with col_u2: arch_com = st.file_uploader("2. Compras", type=["xlsx", "xls"], accept_multiple_files=True)
        with col_u3: arch_fin = st.file_uploader("3. Inv. Final", type=["xlsx", "xls"], accept_multiple_files=True)

        if arch_ini and arch_com and arch_fin:
            try:
                df_inicial = consolidar(arch_ini); df_compras = consolidar(arch_com); df_final = consolidar(arch_fin)
                df_diccionario = obtener_dataframe("Categorias_Costos")
                def limpiar_cod(s): return s.astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
                for df in [df_diccionario, df_inicial, df_compras, df_final]: df['Codigo'] = limpiar_cod(df['Codigo'])
                
                basura = ['G222', 'G231', '21455979']
                df_inicial = df_inicial[~df_inicial['Codigo'].isin(basura)]
                df_compras = df_compras[~df_compras['Codigo'].isin(basura)]
                df_final = df_final[~df_final['Codigo'].isin(basura)]

                df_ini_m = pd.merge(df_inicial, df_diccionario, on='Codigo', how='left')
                df_com_m = pd.merge(df_compras, df_diccionario, on='Codigo', how='left')
                df_fin_m = pd.merge(df_final, df_diccionario, on='Codigo', how='left')

                df_ini_m['Valor'] = get_num(df_ini_m, [['EXISTENCIAS'], ['SALDO']]) * get_num(df_ini_m, [['COSTO', 'U']])
                df_fin_m['Valor'] = get_num(df_fin_m, [['EXISTENCIAS'], ['SALDO']]) * get_num(df_fin_m, [['COSTO', 'U']])
                df_com_m['Valor'] = get_num(df_com_m, [['TOTAL'], ['MONTO']])

                grp_ini = df_ini_m.groupby('Cuenta_Contable')['Valor'].sum()
                grp_comp = df_com_m.groupby('Cuenta_Contable')['Valor'].sum()
                grp_fin = df_fin_m.groupby('Cuenta_Contable')['Valor'].sum()

                todas_cuentas = set(grp_ini.index).union(grp_comp.index).union(grp_fin.index)
                consumo_por_cuenta = {}; costo_operativo = 0.0
                for cta in todas_cuentas:
                    val = grp_ini.get(cta, 0) + grp_comp.get(cta, 0) - grp_fin.get(cta, 0)
                    if val > 0: consumo_por_cuenta[cta] = val; costo_operativo += val
                
                costo_dif_mes = costo_operativo * porcentaje_subsidio
                costo_real = costo_operativo - costo_dif_mes + costo_diferido_anterior

                # Métricas
                r1, r2, r3, r4, r5 = st.columns(5)
                r1.metric("Inicial", f"${df_ini_m['Valor'].sum():,.2f}"); r2.metric("Compras", f"${df_com_m['Valor'].sum():,.2f}")
                r3.metric("Final", f"${df_fin_m['Valor'].sum():,.2f}"); r4.metric("Diferido", f"${costo_dif_mes:,.2f}"); r5.metric("Real", f"${costo_real:,.2f}")
                # --- NUEVO: MÓDULO DE AUDITORÍA ---
                ejecutar_auditoria_costos(df_ini_m, df_com_m, df_fin_m, consumo_por_cuenta, ventas_mes, costo_real)

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
                    mostrar_descargas_logic(costo_operativo, costo_dif_mes, costo_diferido_anterior, mes_txt, consumo_por_cuenta, costo_operativo, "std")
                else:
                    for i in range(len(pesos)):
                        mostrar_descargas_logic(costo_operativo*pesos[i], costo_dif_mes*pesos[i], costo_diferido_anterior*pesos[i], nombres_meses[i], consumo_por_cuenta, costo_operativo, f"cons_{i}")

                if st.button("💾 Guardar Cierre", type="primary", use_container_width=True):
                    with st.spinner("Guardando..."):
                        st.cache_data.clear()
                        ws_res = conectar_hoja("Cierres_Costos"); ws_det = conectar_hoja("Detalle_Cuentas")
                        fecha_hoy = date.today().strftime('%d/%m/%Y')
                        if ws_res and ws_det:
                            ws_res.append_row([fecha_hoy, mes_cierre, anio_cierre, unidad_cierre, round(df_ini_m['Valor'].sum(),2), round(df_com_m['Valor'].sum(),2), round(df_fin_m['Valor'].sum(),2), round(costo_diferido_anterior,2), round(costo_dif_mes,2), round(costo_real,2)])
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
                            st.success("✅ Guardado exitoso.")
            except Exception as e: st.error(f"Error: {e}")

        # ==========================================
    # PESTAÑA 2: TRASLADOS (Gestión de Entradas/Salidas)
    # ==========================================
    with tab2:
        st.subheader("🚚 Gestión y Registro de Traslados")
        ct1, ct2 = st.columns(2)
        with ct1:
            mes_t = st.selectbox("Mes del Traslado:", range(1, 13), index=date.today().month-1, key="mt_reg")
            origen_t = st.text_input("Bodega Origen:", "ABASTECIMIENTO").upper()
        with ct2:
            anio_t = st.number_input("Año:", min_value=2024, value=date.today().year, key="at_reg")
            destino_t = st.text_input("Bodega Destino:", "CAFETERIA").upper()

        arch_t = st.file_uploader("Subir Reporte Nexus (I, N, AA)", type=["xlsx"], accept_multiple_files=True, key="atf_reg")
        
        if arch_t:
            try:
                # Lectura de columnas específicas según Nexus
                dfs_t = [pd.read_excel(a, usecols="I,N,AA", names=['Codigo','Monto','Categoria'], header=None, skiprows=1) for a in arch_t]
                df_t_raw = pd.concat(dfs_t, ignore_index=True)
                df_t_raw['Monto'] = pd.to_numeric(df_t_raw['Monto'], errors='coerce').fillna(0)
                
                # Limpieza de códigos y cruce con diccionario
                df_t_raw['Codigo'] = df_t_raw['Codigo'].astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
                df_dic = obtener_dataframe("Categorias_Costos")
                df_dic['Codigo'] = df_dic['Codigo'].astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
                
                df_final_t = pd.merge(df_t_raw, df_dic[['Codigo', 'Cuenta_Contable']], on='Codigo', how='left')

                st.dataframe(df_final_t, use_container_width=True)

                if st.button("💾 Confirmar y Guardar Traslados en Historial", type="primary"):
                    with st.spinner("Registrando traslados..."):
                        ws_t = conectar_hoja("Historico_Traslados")
                        fecha_hoy = date.today().strftime('%d/%m/%Y')
                        filas_t = []
                        for _, r in df_final_t.iterrows():
                            filas_t.append([
                                fecha_hoy, mes_t, anio_t, origen_t, destino_t, 
                                str(r['Cuenta_Contable']), round(r['Monto'], 2), r['Codigo'], f"Traslado {origen_t}->{destino_t}"
                            ])
                        if filas_t:
                            ws_t.append_rows(filas_t)
                            st.success(f"✅ {len(filas_t)} traslados registrados correctamente.")
            except Exception as e:
                st.error(f"Error al procesar traslados: {e}")

    # ==========================================
    # PESTAÑA 3: CONSULTA DE HISTORIAL
    # ==========================================
    with tab3:
        st.subheader("🔍 Consulta de Cierres Anteriores")
        if st.button("🔄 Actualizar Base", key="update_hist"): st.cache_data.clear()

        df_resumen = obtener_dataframe("Cierres_Costos"); df_detalle = obtener_dataframe("Detalle_Cuentas")

        if not df_resumen.empty:
            df_resumen['Periodo'] = df_resumen['Mes'].astype(str).str.replace('.0','') + "/" + df_resumen['Año'].astype(str).str.replace('.0','') + " - " + df_resumen['Unidad']
            per_sel = st.selectbox("Seleccione Cierre:", df_resumen['Periodo'].unique().tolist(), index=len(df_resumen['Periodo'].unique())-1)

            if per_sel:
                st.divider()
                fila = df_resumen[df_resumen['Periodo'] == per_sel].iloc[-1]
                m_f, a_f = str(fila['Mes']).strip(), str(fila['Año']).strip()
                
                try:
                    v_ini, v_com, v_fin = float(fila.iloc[4]), float(fila.iloc[5]), float(fila.iloc[6])
                    v_ant, v_dif, v_real = float(fila.iloc[7]), float(fila.iloc[8]), float(fila.iloc[9])
                except: v_ini = v_com = v_fin = v_ant = v_dif = v_real = 0.0

                c1, c2, c3, c4, c5 = st.columns(5)
                c1.metric("Inicial", f"${v_ini:,.2f}"); c2.metric("Compras", f"${v_com:,.2f}")
                c3.metric("Final", f"${v_fin:,.2f}"); c4.metric("Diferido", f"${v_dif:,.2f}"); c5.metric("Real", f"${v_real:,.2f}")

                df_det_h = df_detalle[(df_detalle['Mes'].astype(str).str.replace('.0','') == m_f) & (df_detalle['Año'].astype(str).str.replace('.0','') == a_f)]
                df_det_h['Consumo'] = pd.to_numeric(df_det_h['Consumo'], errors='coerce').fillna(0)

                if st.checkbox("📂 Ver Detalle de Movimientos", key="chk_det"):
                    st.dataframe(df_det_h, use_container_width=True)

                if st.checkbox("📊 Ver Resumen por Cuentas", key="chk_res"):
                    st.table(df_det_h.groupby('Cuenta')['Consumo'].sum().reset_index())

                st.divider(); st.markdown("#### 📥 Partidas Nexus Reconstruidas")
                
                # REPLICAR LÓGICA DE CONSOLIDACIÓN EN HISTÓRICO
                tipo_h = st.radio("¿Como reconstruir estas partidas?", ["Cierre Estándar", "Cierre Consolidado (Dividir)"], key="th")
                
                consumo_h = df_det_h.groupby('Cuenta')['Consumo'].sum().to_dict()
                op_h = sum(consumo_h.values())

                if tipo_h == "Cierre Estándar":
                    cv_h = f"RECONOCIMIENTO DE COSTO DE VENTA DE CAFETERIA, {m_f}/{a_f}."
                    fv_h = [["410104", "", cv_h, round(op_h, 2), 0, ""]]
                    for c, m in consumo_h.items():
                        if m > 0: fv_h.append([str(c).replace(".0",""), "", cv_h, 0, round(m, 2), ""])
                    
                    cp_h = f"RECONOCIMIENTO DE COSTO DE LO VENDIDO EN PROCESO CAFETERIA {m_f}/{a_f}"
                    fp_h = [["410104", "", cp_h, round(v_ant, 2), 0, ""], ["110602", "", cp_h, 0, round(v_ant, 2), ""]]
                    cd_h = f"DIFERIMIENTO DE COSTO EN PROCESO CAFETERIA {m_f}/{a_f}"
                    fd_h = [["110602", "", cd_h, round(v_dif, 2), 0, ""], ["410104", "", cd_h, 0, round(v_dif, 2), ""]]
                    
                    h1, h2, h3 = st.columns(3)
                    with h1: st.download_button(f"⬇️ P1 {m_f}", generar_excel_bytes(fv_h), f"H_P1_{m_f}.xlsx", key="hp1_std")
                    with h2: st.download_button(f"⬇️ P2 {m_f}", generar_excel_bytes(fp_h), f"H_P2_{m_f}.xlsx", key="hp2_std")
                    with h3: st.download_button(f"⬇️ P3 {m_f}", generar_excel_bytes(fd_h), f"H_P3_{m_f}.xlsx", key="hp3_std")
                else:
                    # NOTA: En este punto, como el histórico no guarda el porcentaje individual, 
                    # te permito visualizar los meses, pero de forma estática si tuvieras la data.
                    # Para esta versión, lo dejo informativo pero funcional para reconstruir.
                    n_mes_h = st.number_input("¿En cuántos meses dividir?", 2, 3, 3, key="n_split")
                    cols_h = st.columns(n_mes_h); p_h, n_h = [], []
                    for i in range(n_mes_h):
                        with cols_h[i]:
                            nm = st.text_input(f"Mes {i+1}:", key=f"nm_h_{i}").upper()
                            ph = st.number_input(f"% {nm}", 0.0, 100.0, 100.0/n_mes_h, key=f"ph_h_{i}")
                            p_h.append(ph/100); n_h.append(nm)
                    
                    for i in range(n_mes_h):
                        st.markdown(f"**Partidas para {n_h[i]}**")
                        cv_c = f"RECONOCIMIENTO DE COSTO DE VENTA DE CAFETERIA, {n_h[i]} {a_f}."
                        fv_c = [["410104", "", cv_c, round(op_h*p_h[i], 2), 0, ""]]
                        for c, m in consumo_h.items():
                            mp_c = (m * p_h[i])
                            if mp_c > 0: fv_c.append([str(c).replace(".0",""), "", cv_c, 0, round(mp_c, 2), ""])
                        
                        cp_c = f"RECONOCIMIENTO DE COSTO DE LO VENDIDO EN PROCESO CAFETERIA {n_h[i]} {a_f}"
                        fp_c = [["410104", "", cp_c, round(v_ant*p_h[i], 2), 0, ""], ["110602", "", cp_c, 0, round(v_ant*p_h[i], 2), ""]]
                        cd_c = f"DIFERIMIENTO DE COSTO EN PROCESO CAFETERIA {n_h[i]} {a_f}"
                        fd_c = [["110602", "", cd_c, round(v_dif*p_h[i], 2), 0, ""], ["410104", "", cd_c, 0, round(v_dif*p_h[i], 2), ""]]
                        
                        h1c, h2c, h3c = st.columns(3)
                        with h1c: st.download_button(f"P1 {n_h[i]}", generar_excel_bytes(fv_c), f"H_P1_{n_h[i]}.xlsx", key=f"hp1_c_{i}")
                        with h2c: st.download_button(f"P2 {n_h[i]}", generar_excel_bytes(fp_c), f"H_P2_{n_h[i]}.xlsx", key=f"hp2_c_{i}")
                        with h3c: st.download_button(f"P3 {n_h[i]}", generar_excel_bytes(fd_c), f"H_P3_{n_h[i]}.xlsx", key=f"hp3_c_{i}")
                        st.divider()