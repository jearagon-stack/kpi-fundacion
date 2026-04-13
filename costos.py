import streamlit as st
import pandas as pd
from datetime import date
from utils import obtener_dataframe, conectar_hoja
import io

def mostrar_modulo_costos():
    st.title("Contabilidad de Costos")

    # --- CREACIÓN DE PESTAÑAS ---
    tab1, tab2, tab3 = st.tabs(["📝 Generar Cierre", "🔍 Consultar Histórico", "🚚 Partidas de Traslados"])

    # ==========================================
    # PESTAÑA 1: LÓGICA DE CÁLCULO Y GUARDADO
    # ==========================================
    with tab1:
        st.subheader("1. Cierre Contable")
        
        col_config1, col_config2 = st.columns([2, 1])
        with col_config1:
            tipo_cierre = st.radio("Tipo de proceso:", ["Mensual Estándar", "Consolidación Especial (Multi-mes)"], horizontal=True)

        col1, col2, col3 = st.columns(3)
        with col1:
            mes_cierre = st.selectbox("Mes a costear (o mes final):", range(1, 13), index=date.today().month - 1)
        with col2:
            anio_cierre = st.number_input("Año:", min_value=2024, max_value=2030, value=date.today().year)
        with col3:
            unidad_cierre = st.selectbox("Unidad a consolidar:", ["CAFETERIA"])

        meses_texto = {1:"ENERO", 2:"FEBRERO", 3:"MARZO", 4:"ABRIL", 5:"MAYO", 6:"JUNIO", 7:"JULIO", 8:"AGOSTO", 9:"SEPTIEMBRE", 10:"OCTUBRE", 11:"NOVIEMBRE", 12:"DICIEMBRE"}
        mes_txt = meses_texto[mes_cierre]

        es_consolidado = (tipo_cierre == "Consolidación Especial (Multi-mes)")
        pesos = []
        nombres_meses = []

        if es_consolidado:
            st.warning("⚠️ Modo Consolidación: El costo total se agrupará y luego se distribuirá en las partidas de cada mes.")
            num_meses = st.radio("¿En cuántos meses vas a dividir el costo?", [2, 3], horizontal=True)
            
            cols_dist = st.columns(num_meses)
            if num_meses == 2:
                with cols_dist[0]:
                    n1 = st.text_input("Mes 1:", "FEBRERO").upper()
                    p1 = st.number_input(f"% Venta {n1}", min_value=0.0, max_value=100.0, value=50.0)
                    pesos.append(p1 / 100)
                    nombres_meses.append(n1)
                with cols_dist[1]:
                    n2 = st.text_input("Mes 2:", "MARZO").upper()
                    st.info(f"% {n2}: {100.0 - p1}%")
                    pesos.append((100.0 - p1) / 100)
                    nombres_meses.append(n2)
            elif num_meses == 3:
                with cols_dist[0]:
                    n1 = st.text_input("Mes 1:", "ENERO").upper()
                    p1 = st.number_input(f"% Venta {n1}", min_value=0.0, max_value=100.0, value=33.3)
                    pesos.append(p1 / 100)
                    nombres_meses.append(n1)
                with cols_dist[1]:
                    n2 = st.text_input("Mes 2:", "FEBRERO").upper()
                    p2 = st.number_input(f"% Venta {n2}", min_value=0.0, max_value=100.0, value=33.3)
                    pesos.append(p2 / 100)
                    nombres_meses.append(n2)
                with cols_dist[2]:
                    n3 = st.text_input("Mes 3:", "MARZO").upper()
                    p3 = max(0.0, 100.0 - p1 - p2)
                    st.info(f"% {n3}: {round(p3, 2)}%")
                    pesos.append(p3 / 100)
                    nombres_meses.append(n3)

        st.divider()

        # --- CÁLCULO DE INGRESOS Y SUBSIDIOS ---
        df_ventas = obtener_dataframe("Historico_Ventas")
        ventas_mes = 0.0
        subsidio_mes = 0.0
        
        if not df_ventas.empty:
            df_ventas['Fecha_DT'] = pd.to_datetime(df_ventas['Fecha'], format='%d/%m/%Y', errors='coerce')
            df_ventas['Venta_Real'] = pd.to_numeric(df_ventas['Venta_Real'], errors='coerce').fillna(0)
            df_ventas['Subsidio_UCA'] = pd.to_numeric(df_ventas.get('Subsidio_UCA', 0), errors='coerce').fillna(0)
            
            if es_consolidado:
                meses_cons = st.multiselect("Selecciona los meses incluidos en esta consolidación:", range(1, 13), default=[1, 2, 3])
                filtro = (df_ventas['Fecha_DT'].dt.month.isin(meses_cons)) & (df_ventas['Fecha_DT'].dt.year == anio_cierre) & (df_ventas['Unidad'] == unidad_cierre)
            else:
                filtro = (df_ventas['Fecha_DT'].dt.month == mes_cierre) & (df_ventas['Fecha_DT'].dt.year == anio_cierre) & (df_ventas['Unidad'] == unidad_cierre)
            
            ventas_mes = df_ventas[filtro]['Venta_Real'].sum()
            subsidio_mes = df_ventas[filtro]['Subsidio_UCA'].sum()

        porcentaje_subsidio = (subsidio_mes / ventas_mes) if ventas_mes > 0 else 0.0

        if ventas_mes == 0:
            st.warning(f"Nota: No hay ventas registradas para {unidad_cierre} en el periodo seleccionado.")
        else:
            st.info(f"📊 Información de Ingresos: Ventas ${ventas_mes:,.2f} | Subsidio ${subsidio_mes:,.2f} | % Subsidio: {porcentaje_subsidio:.2%}")

        # 2. CARGA DE VARIABLES Y ARCHIVOS 
        st.subheader("2. Variables Operativas")
        costo_diferido_anterior = st.number_input("Costo Diferido de Arrastre (Cuenta 110602):", min_value=0.0, value=0.0, step=100.0)

        with st.expander("Carga de Documentos del Periodo", expanded=True):
            col_a1, col_a2, col_a3 = st.columns(3)
            with col_a1: archivos_iniciales = st.file_uploader("1. Inv. Inicial (Kardex)", type=["xlsx", "xls"], accept_multiple_files=True)
            with col_a2: archivos_compras = st.file_uploader("2. Compras del Mes", type=["xlsx", "xls"], accept_multiple_files=True)
            with col_a3: archivos_finales = st.file_uploader("3. Inv. Final (Kardex)", type=["xlsx", "xls"], accept_multiple_files=True)

        # 3. PROCESAMIENTO
        if archivos_iniciales and archivos_compras and archivos_finales:
            st.divider()
            st.subheader("3. Resultados y Cierre")
            try:
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

                df_inicial = consolidar(archivos_iniciales)
                df_compras = consolidar(archivos_compras)
                df_final = consolidar(archivos_finales)

                df_diccionario = obtener_dataframe("Categorias_Costos")
                def limpiar(s): return s.astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
                
                for df in [df_diccionario, df_inicial, df_compras, df_final]: 
                    df['Codigo'] = limpiar(df['Codigo'])
                
                basura = ['G222', 'G231', '21455979']
                df_inicial = df_inicial[~df_inicial['Codigo'].isin(basura)]
                df_compras = df_compras[~df_compras['Codigo'].isin(basura)]
                df_final = df_final[~df_final['Codigo'].isin(basura)]

                df_ini_m = pd.merge(df_inicial, df_diccionario, on='Codigo', how='left')
                df_com_m = pd.merge(df_compras, df_diccionario, on='Codigo', how='left')
                df_fin_m = pd.merge(df_final, df_diccionario, on='Codigo', how='left')

                df_ini_m = df_ini_m[df_ini_m['Categoria_Nexus'].str.strip() != 'SERVICIO']
                df_com_m = df_com_m[df_com_m['Categoria_Nexus'].str.strip() != 'SERVICIO']
                df_fin_m = df_fin_m[df_fin_m['Categoria_Nexus'].str.strip() != 'SERVICIO']

                def get_num(df, keys):
                    for k in keys:
                        for col in df.columns:
                            c_norm = str(col).upper().replace(' ', '').replace('.', '')
                            if all(p in c_norm for p in k): return pd.to_numeric(df[col], errors='coerce').fillna(0)
                    return pd.Series(0, index=df.index)

                df_ini_m['Valor'] = get_num(df_ini_m, [['EXISTENCIAS', 'FINAL'], ['SALDO', 'FINAL'], ['CANTIDAD']]) * get_num(df_ini_m, [['COSTO', 'U'], ['COSTO', 'UNITARIO'], ['PRECIO']])
                df_fin_m['Valor'] = get_num(df_fin_m, [['EXISTENCIAS', 'FINAL'], ['SALDO', 'FINAL'], ['CANTIDAD']]) * get_num(df_fin_m, [['COSTO', 'U'], ['COSTO', 'UNITARIO'], ['PRECIO']])
                df_com_m['Valor'] = get_num(df_com_m, [['TOTAL'], ['MONTO']])

                grp_ini = df_ini_m.groupby('Cuenta_Contable')['Valor'].sum()
                grp_comp = df_com_m.groupby('Cuenta_Contable')['Valor'].sum()
                grp_fin = df_fin_m.groupby('Cuenta_Contable')['Valor'].sum()

                todas_cuentas = set(grp_ini.index).union(grp_comp.index).union(grp_fin.index)
                consumo_por_cuenta = {}
                costo_operativo = 0.0
                
                for cta in todas_cuentas:
                    val = grp_ini.get(cta, 0) + grp_comp.get(cta, 0) - grp_fin.get(cta, 0)
                    if val > 0:
                        consumo_por_cuenta[cta] = val
                        costo_operativo += val
                
                costo_dif_mes = costo_operativo * porcentaje_subsidio
                costo_real = costo_operativo - costo_dif_mes + costo_diferido_anterior

                r1, r2, r3, r4, r5 = st.columns(5)
                r1.metric("Inv. Inicial", f"${df_ini_m['Valor'].sum():,.2f}")
                r2.metric("Compras", f"${df_com_m['Valor'].sum():,.2f}")
                r3.metric("Inv. Final", f"${df_fin_m['Valor'].sum():,.2f}")
                r4.metric("Diferido Total", f"${costo_dif_mes:,.2f}")
                r5.metric("Costo Real Total", f"${costo_real:,.2f}")

                def generar_excel_bytes(filas):
                    df_p = pd.DataFrame(filas)
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_p.to_excel(writer, index=False, header=False, sheet_name='Hoja1')
                    return output.getvalue()

                def mostrar_descargas(c_op, c_dif, c_ant, label_mes):
                    st.markdown(f"#### 📥 Descargas para: **{label_mes}**")
                    conc_v = f"RECONOCIMIENTO DE COSTO DE VENTA DE CAFETERIA, {label_mes} {anio_cierre}."
                    f_v = [["410104", "", conc_v, round(c_op, 2), 0.00, ""]]
                    for c, m in consumo_por_cuenta.items():
                        m_perc = m * (c_op / costo_operativo) if costo_operativo > 0 else 0
                        if m_perc > 0:
                            f_v.append([str(c).replace(".0",""), "", conc_v, 0.00, round(m_perc, 2), ""])
                    conc_p = f"RECONOCIMIENTO DE COSTO DE LO VENDIDO EN PROCESO CAFETERIA {label_mes} {anio_cierre}"
                    f_p = [["410104", "", conc_p, round(c_ant, 2), 0.00, ""], ["110602", "", conc_p, 0.00, round(c_ant, 2), ""]]
                    conc_d = f"DIFERIMIENTO DE COSTO EN PROCESO CAFETERIA {label_mes} {anio_cierre}"
                    f_d = [["110602", "", conc_d, round(c_dif, 2), 0.00, ""], ["410104", "", conc_d, 0.00, round(c_dif, 2), ""]]
                    p1, p2, p3 = st.columns(3)
                    with p1: st.download_button(f"⬇️ P1: Costo Venta", generar_excel_bytes(f_v), f"1_CostoVenta_{label_mes}.xlsx", use_container_width=True)
                    with p2: st.download_button(f"⬇️ P2: Proceso Ant.", generar_excel_bytes(f_p), f"2_ProcesoAnt_{label_mes}.xlsx", use_container_width=True)
                    with p3: st.download_button(f"⬇️ P3: Diferimiento", generar_excel_bytes(f_d), f"3_Diferimiento_{label_mes}.xlsx", use_container_width=True)

                st.subheader("Generación de Partidas Nexus")
                if not es_consolidado:
                    mostrar_descargas(costo_operativo, costo_dif_mes, costo_diferido_anterior, mes_txt)
                else:
                    for i in range(num_meses):
                        mostrar_descargas(costo_operativo * pesos[i], costo_dif_mes * pesos[i], costo_diferido_anterior * pesos[i], nombres_meses[i])
                        if i < num_meses - 1: st.divider()

                def extraer_subunidad(nombre_archivo, unidad_base):
                    nombre_upper = str(nombre_archivo).upper()
                    subunidades = ["POLIDEPORTIVO", "ICAS", "JARDINES", "EVENTOS", "CENTRAL", "ABASTECIMIENTO"]
                    for sub in subunidades:
                        if sub in nombre_upper: return f"{unidad_base} {sub}"
                    return unidad_base

                st.write("---")
                if st.button("💾 Guardar Cierre e Historial Detallado", type="primary", use_container_width=True):
                    with st.spinner("Guardando..."):
                        try:
                            st.cache_data.clear()
                            ws_res = conectar_hoja("Cierres_Costos")
                            ws_det = conectar_hoja("Detalle_Cuentas")
                            fecha = date.today().strftime('%d/%m/%Y')
                            if ws_res and ws_det:
                                ws_res.append_row([fecha, mes_cierre, anio_cierre, unidad_cierre, round(df_ini_m['Valor'].sum(),2), round(df_com_m['Valor'].sum(),2), round(df_fin_m['Valor'].sum(),2), round(costo_diferido_anterior,2), round(costo_dif_mes,2), round(costo_real,2)])
                                df_ini_g = df_ini_m[['Codigo', 'Cuenta_Contable', 'Valor', 'ORIGEN_ARCHIVO']].rename(columns={'Valor': 'Inicial'})
                                df_com_g = df_com_m[['Codigo', 'Cuenta_Contable', 'Valor', 'ORIGEN_ARCHIVO']].rename(columns={'Valor': 'Compra'})
                                df_fin_g = df_fin_m[['Codigo', 'Cuenta_Contable', 'Valor', 'ORIGEN_ARCHIVO']].rename(columns={'Valor': 'Final'})
                                df_det = pd.concat([df_ini_g, df_com_g, df_fin_g]).fillna(0)
                                df_det = df_det.groupby(['Codigo', 'Cuenta_Contable', 'ORIGEN_ARCHIVO']).sum().reset_index()
                                df_det['Consumo'] = df_det['Inicial'] + df_det['Compra'] - df_det['Final']
                                filas = []
                                for _, r in df_det.iterrows():
                                    if r['Inicial']!=0 or r['Compra']!=0 or r['Final']!=0:
                                        unidad_real = extraer_subunidad(r['ORIGEN_ARCHIVO'], unidad_cierre)
                                        filas.append([fecha, mes_cierre, anio_cierre, unidad_real, str(r['Cuenta_Contable']), round(r['Inicial'],2), round(r['Compra'],2), round(r['Final'],2), round(r['Consumo'],2), r['Codigo'], r['ORIGEN_ARCHIVO']])
                                if filas: ws_det.append_rows(filas)
                                st.success("✅ Guardado exitoso.")
                        except Exception as e: st.error(f"Error: {e}")
            except Exception as e: st.error(f"Error: {e}")

    # ==========================================
    # PESTAÑA 2: CONSULTA DE HISTORIAL (CORREGIDA)
    # ==========================================
    with tab2:
        st.subheader("🔍 Consulta de Cierres Anteriores")
        col_btn, _ = st.columns([1, 4])
        with col_btn:
            if st.button("🔄 Actualizar Base", use_container_width=True): st.cache_data.clear()

        df_resumen = obtener_dataframe("Cierres_Costos")
        df_detalle = obtener_dataframe("Detalle_Cuentas")

        if not df_resumen.empty and not df_detalle.empty:
            df_resumen['Mes'] = df_resumen['Mes'].astype(str).str.replace('.0', '', regex=False)
            df_resumen['Año'] = df_resumen['Año'].astype(str).str.replace('.0', '', regex=False)
            df_resumen['Periodo'] = df_resumen['Mes'] + "/" + df_resumen['Año'] + " - " + df_resumen['Unidad']
            periodos_disponibles = df_resumen['Periodo'].unique().tolist()
            periodo_sel = st.selectbox("Periodo cerrado:", periodos_disponibles, index=len(periodos_disponibles)-1)

            if periodo_sel:
                st.divider()
                fila_resumen = df_resumen[df_resumen['Periodo'] == periodo_sel].iloc[-1]
                mes_f, anio_f = str(fila_resumen['Mes']).strip(), str(fila_resumen['Año']).strip()
                
                try:
                    i_ini, i_com, i_fin = float(fila_resumen.iloc[4]), float(fila_resumen.iloc[5]), float(fila_resumen.iloc[6])
                    c_ant_h, dif_h, c_real_h = float(fila_resumen.iloc[7]), float(fila_resumen.iloc[8]), float(fila_resumen.iloc[9])
                except: i_ini = i_com = i_fin = c_ant_h = dif_h = c_real_h = 0.0

                r1, r2, r3, r4, r5 = st.columns(5)
                r1.metric("Inv. Inicial", f"${i_ini:,.2f}"); r2.metric("Compras", f"${i_com:,.2f}"); r3.metric("Inv. Final", f"${i_fin:,.2f}"); r4.metric("Diferido", f"${dif_h:,.2f}"); r5.metric("Costo Real", f"${c_real_h:,.2f}")

                df_detalle['Mes'] = df_detalle['Mes'].astype(str).str.replace('.0', '', regex=False).str.strip()
                df_detalle['Año'] = df_detalle['Año'].astype(str).str.replace('.0', '', regex=False).str.strip()
                df_det_f = df_detalle[(df_detalle['Mes'] == mes_f) & (df_detalle['Año'] == anio_f)]
                
                if not df_det_f.empty:
                    st.markdown("#### 📑 Detalle Operativo")
                    st.dataframe(df_det_f, use_container_width=True)
                    df_det_f['Consumo'] = pd.to_numeric(df_det_f['Consumo'], errors='coerce').fillna(0)
                    df_agrupado = df_det_f.groupby('Cuenta')['Consumo'].sum().reset_index()
                    st.table(df_agrupado)

                    # Partidas Históricas
                    st.divider()
                    st.markdown("#### 📥 Partidas Nexus Reconstruidas")
                    c_op_h = df_agrupado['Consumo'].sum()
                    cons_h = dict(zip(df_agrupado['Cuenta'], df_agrupado['Consumo']))
                    
                    conc_v = f"RECONOCIMIENTO DE COSTO DE VENTA CAFETERIA, {mes_f}/{anio_f}."
                    f_v_h = [["410104", "", conc_v, round(c_op_h, 2), 0.00, ""]]
                    for c, m in cons_h.items():
                        if m > 0: f_v_h.append([str(c).replace(".0",""), "", conc_v, 0.00, round(m, 2), ""])
                    
                    f_p_h = [["410104", "", "PROCESO ANT", round(c_ant_h, 2), 0.00, ""], ["110602", "", "PROCESO ANT", 0.00, round(c_ant_h, 2), ""]]
                    f_d_h = [["110602", "", "DIFERIMIENTO", round(dif_h, 2), 0.00, ""], ["410104", "", "DIFERIMIENTO", 0.00, round(dif_h, 2), ""]]

                    def gen_h(filas):
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer: pd.DataFrame(filas).to_excel(writer, index=False, header=False)
                        return output.getvalue()

                    ch1, ch2, ch3 = st.columns(3)
                    with ch1: st.download_button("P1 Costo", gen_h(f_v_h), f"H_P1_{mes_f}_{anio_f}.xlsx", use_container_width=True)
                    with ch2: st.download_button("P2 Proc.Ant", gen_h(f_p_h), f"H_P2_{mes_f}_{anio_f}.xlsx", use_container_width=True)
                    with ch3: st.download_button("P3 Diferir", gen_h(f_d_h), f"H_P3_{mes_f}_{anio_f}.xlsx", use_container_width=True)

    # ==========================================
    # PESTAÑA 3: PARTIDAS DE TRASLADOS
    # ==========================================
    with tab3:
        st.subheader("🚚 Partidas de Traslados")
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            mes_t = st.selectbox("Mes:", range(1, 13), index=date.today().month - 1, key="mt")
            orig_t = st.text_input("Origen:", "ABASTECIMIENTO").upper()
        with col_t2:
            anio_t = st.number_input("Año:", min_value=2024, max_value=2030, value=date.today().year, key="at")
            dest_t = st.text_input("Destino:", "DESPENSA").upper()

        st.info("💡 Sube el reporte de traslados Nexus (I: Código, N: Monto, AA: Categoría).")
        arch_t = st.file_uploader("Excel Traslados", type=["xlsx", "xls"], accept_multiple_files=True, key="atf")

        if arch_t:
            try:
                dfs_t = [pd.read_excel(a, usecols="I,N,AA", names=['Codigo', 'Monto', 'Categoria'], header=None, skiprows=1) for a in arch_t]
                df_t = pd.concat(dfs_t, ignore_index=True)
                df_t['Monto'] = pd.to_numeric(df_t['Monto'], errors='coerce').fillna(0)
                df_t = df_t[(df_t['Monto'] > 0) & (df_t['Categoria'].astype(str).str.upper().str.strip() != 'SERVICIO')]
                df_t['Codigo'] = df_t['Codigo'].astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
                
                df_dic = obtener_dataframe("Categorias_Costos")
                df_dic['Codigo'] = df_dic['Codigo'].astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
                df_f_t = pd.merge(df_t, df_dic, on='Codigo', how='left')
                
                huerf = df_f_t[df_f_t['Cuenta_Contable'].isna()]['Codigo'].unique().tolist()
                if huerf: st.error(f"Cod. no encontrados: {huerf}"); st.stop()

                grp_t = df_f_t.groupby('Cuenta_Contable')['Monto'].sum()
                st.success(f"Total: ${grp_t.sum():,.2f}")
                st.dataframe(grp_t.reset_index(), use_container_width=True)

                filas_p = []
                conc_t = f"TRASLADO {orig_t} A {dest_t}, {meses_texto[mes_t]} {anio_t}"
                for c, m in grp_t.items(): filas_p.append([str(c).replace(".0",""), "", conc_t, round(m, 2), 0.00, ""])
                for c, m in grp_t.items(): filas_p.append([str(c).replace(".0",""), "", conc_t, 0.00, round(m, 2), ""])
                
                st.download_button("⬇️ Descargar Traslado", generar_excel_bytes(filas_p), f"Traslado_{orig_t}.xlsx", use_container_width=True, type="primary")
            except Exception as e: st.error(f"Error: {e}")