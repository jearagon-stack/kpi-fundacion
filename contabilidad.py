import streamlit as st
import pandas as pd
import plotly.express as px
import io
from utils import obtener_dataframe

# --- FUNCIONES AUXILIARES ---
def limpiar_texto(texto):
    return str(texto).strip().upper()

def buscar_columna(df, palabras_clave, todas=False, excluir=None, anti_palabras=None):
    """Busca columnas de forma inteligente y permite bloquear palabras específicas (ej. 'INICIAL')."""
    if excluir is None:
        excluir = []
    if anti_palabras is None:
        anti_palabras = []
        
    # 1. Búsqueda por coincidencia exacta normalizada (ignora espacios y guiones)
    for col in df.columns:
        if col in excluir: continue
        c_norm = str(col).strip().upper().replace(' ', '').replace('_', '').replace('-', '')
        
        # Filtro de bloqueo de anti-palabras
        if any(ap.strip().upper() in c_norm for ap in anti_palabras):
            continue
            
        if any(c_norm == p.strip().upper().replace(' ', '').replace('_', '').replace('-', '') for p in palabras_clave):
            return col
            
    # 2. Búsqueda por coincidencia parcial normalizada
    for col in df.columns:
        if col in excluir: continue
        c_norm = str(col).strip().upper().replace(' ', '').replace('_', '').replace('-', '')
        
        # Filtro de bloqueo de anti-palabras
        if any(ap.strip().upper() in c_norm for ap in anti_palabras):
            continue
            
        if todas:
            if all(p.strip().upper().replace(' ', '').replace('_', '').replace('-', '') in c_norm for p in palabras_clave):
                return col
        else:
            if any(p.strip().upper().replace(' ', '').replace('_', '').replace('-', '') in c_norm for p in palabras_clave):
                return col
    return None

def generar_excel_multi(dict_dfs):
    """Genera un archivo Excel con múltiples pestañas a partir de un diccionario de DataFrames."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for nombre_hoja, df in dict_dfs.items():
            df.to_excel(writer, index=False, sheet_name=nombre_hoja[:31])
    return output.getvalue()

def cargar_archivo_contable(archivo):
    """Escanea el archivo saltando encabezados basura de sistemas ERP y purga columnas clonadas."""
    if archivo.name.lower().endswith('.csv'):
        archivo.seek(0)
        linea = archivo.readline().decode('utf-8', errors='ignore')
        separador = ';' if ';' in linea else ','
        archivo.seek(0)
        df = pd.read_csv(archivo, sep=separador, dtype=str)
    else:
        try:
            df = pd.read_excel(archivo, dtype=str)
        except Exception:
            archivo.seek(0)
            dfs = pd.read_html(archivo.read().decode('utf-8'))
            df = dfs[0].astype(str)
            
    df.columns = df.columns.astype(str).str.strip().str.upper().str.replace('\ufeff', '')
    cols_upper = df.columns.str.replace(' ', '').str.replace('.', '')
    
    if not (any('CUENTA' in c or 'ID' in c for c in cols_upper) and any('SALDO' in c for c in cols_upper)):
        for i in range(min(20, len(df))):
            row_str = df.iloc[i].astype(str).str.upper().str.replace(' ', '').str.replace('.', '')
            if (any('CUENTA' in val or 'ID' in val for val in row_str) and any('SALDO' in val for val in row_str)):
                df.columns = df.iloc[i].astype(str).str.strip().str.upper()
                df = df.iloc[i+1:].reset_index(drop=True)
                break
                
    df = df.loc[:, ~df.columns.duplicated()].copy()
    return df

def procesar_prorrateo_matriz(df_base, col_cta, col_nom, col_est, col_tipo, unidades_oficiales, unidad_gerencia="Gerencias"):
    df_res = df_base.copy()
    
    ventas = {}
    total_ventas = 0
    for u in unidades_oficiales:
        v = df_res[df_res[col_tipo].isin(["INGRESOS", "INGRESO"])][u].sum()
        ventas[u] = v
        total_ventas += v
        
    pct_ventas = {u: (ventas[u] / total_ventas if total_ventas > 0 else 0) for u in unidades_oficiales}
    gasto_gerencia = df_res[df_res[col_tipo].isin(["COSTO FIJO", "COSTO VARIABLE"])][unidad_gerencia].sum()
    resumen_data = []
    
    def crear_fila(nombre_etiqueta, tipo_indicador="RESUMEN"):
        f = {col_cta: "", col_nom: nombre_etiqueta, col_est: "", col_tipo: tipo_indicador}
        for u in unidades_oficiales + ['TOTAL CONSOLIDADO']:
            f[u] = 0
        return f

    fila_pct = crear_fila("% PARTICIPACIÓN (VENTAS)", "INDICADOR")
    for u in unidades_oficiales: fila_pct[u] = pct_ventas[u] * 100
    fila_pct['TOTAL CONSOLIDADO'] = 100 if total_ventas > 0 else 0
    resumen_data.append(fila_pct)
    
    fila_ing = crear_fila("INGRESOS TOTALES")
    for u in unidades_oficiales: fila_ing[u] = ventas[u]
    fila_ing['TOTAL CONSOLIDADO'] = total_ventas
    resumen_data.append(fila_ing)
    
    fila_cfp = crear_fila("COSTOS FIJOS (PROPIOS)")
    tot_cfp = 0
    for u in unidades_oficiales:
        cfp = 0 if u == unidad_gerencia else df_res[df_res[col_tipo] == "COSTO FIJO"][u].sum()
        fila_cfp[u] = cfp
        tot_cfp += cfp
    fila_cfp['TOTAL CONSOLIDADO'] = tot_cfp
    resumen_data.append(fila_cfp)
    
    fila_cv = crear_fila("COSTOS VARIABLES")
    tot_cv = 0
    for u in unidades_oficiales:
        cv = 0 if u == unidad_gerencia else df_res[df_res[col_tipo] == "COSTO VARIABLE"][u].sum()
        fila_cv[u] = cv
        tot_cv += cv
    fila_cv['TOTAL CONSOLIDADO'] = tot_cv
    resumen_data.append(fila_cv)
    
    fila_cif = crear_fila("GASTO ADM. ASIGNADO (GERENCIAS)")
    tot_cif = 0
    for u in unidades_oficiales:
        cif_asignado = gasto_gerencia * pct_ventas[u]
        fila_cif[u] = cif_asignado
        tot_cif += cif_asignado
    fila_cif['TOTAL CONSOLIDADO'] = tot_cif
    resumen_data.append(fila_cif)
    
    fila_pe = crear_fila("PUNTO DE EQUILIBRIO (CON CIF)", "INDICADOR")
    tot_pe_global = 0
    for u in unidades_oficiales:
        margen = (1 - (fila_cv[u] / ventas[u])) if ventas[u] > 0 else 0
        pe = ((fila_cfp[u] + fila_cif[u]) / margen) if margen > 0 else 0
        fila_pe[u] = pe
        tot_pe_global += pe
    fila_pe['TOTAL CONSOLIDADO'] = tot_pe_global
    resumen_data.append(fila_pe)
    
    fila_ms = crear_fila("SUPERÁVIT / DÉFICIT ($)", "INDICADOR")
    tot_ms = 0
    for u in unidades_oficiales:
        ms = fila_ing[u] - fila_pe[u]
        fila_ms[u] = ms
        tot_ms += ms
    fila_ms['TOTAL CONSOLIDADO'] = tot_ms
    resumen_data.append(fila_ms)

    fila_sep = {col_cta: "", col_nom: "-"*30, col_est: "", col_tipo: ""}
    for u in unidades_oficiales + ['TOTAL CONSOLIDADO']: fila_sep[u] = None
    
    df_resumen = pd.DataFrame([fila_sep] + resumen_data)
    df_final = pd.concat([df_res, df_resumen], ignore_index=True)
    return df_final, pct_ventas


def mostrar_modulo_contabilidad():
    st.title("📊 Análisis Financiero y Punto de Equilibrio")
    
    tab_carga, tab_pe, tab_dash, tab_matriz, tab_proy, tab_comp = st.tabs([
        "📥 1. Carga", 
        "🎛️ 2. Simulador", 
        "📊 3. Dashboard",
        "📑 4. Matriz 2025",
        "🚀 5. Proyecciones 26",
        "⚖️ 6. Comparativo"
    ])

    unidades_oficiales = ["Cafetería", "Librería", "Centro Soho", "CID Campus", "Talleres Gráfico", "Despensa", "Terraza", "Servicios Generales", "Gerencias"]

    # ==========================================
    # PESTAÑA 1: CARGA GENERAL
    # ==========================================
    with tab_carga:
        st.subheader("Configuración para el Análisis Global")
        archivos_subidos = st.file_uploader("Sube los archivos Excel mensuales (Consolidado):", type=["xlsx", "xls"], accept_multiple_files=True, key="up_general")

        if st.button("🔄 Procesar Consolidado Global", type="primary", use_container_width=True):
            if archivos_subidos:
                with st.spinner("Procesando datos y asegurando columnas..."):
                    try:
                        df_map = obtener_dataframe("Balance_Mapeado")
                        if df_map is None or df_map.empty:
                            st.error("❌ No se pudo cargar el catálogo mapeado de la base de datos.")
                            st.stop()
                            
                        df_map.columns = df_map.columns.astype(str).str.strip().str.upper()
                        df_map = df_map.loc[:, ~df_map.columns.duplicated()].copy()

                        c_map_cta = buscar_columna(df_map, ["IDCUENTA", "CUENTA", "CODIGO", "ID"])
                        c_map_nom = buscar_columna(df_map, ["NOMBRE", "DESCRIP"], excluir=[c_map_cta]) 
                        c_map_tipo = buscar_columna(df_map, ["TIPO", "CLASIFICACION"], excluir=[c_map_cta, c_map_nom])
                        c_map_est = buscar_columna(df_map, ["ESTADO", "GRUPO"], excluir=[c_map_cta, c_map_nom, c_map_tipo])
                        c_map_cat = buscar_columna(df_map, ["CATEGORIA", "CATEGOR"], excluir=[c_map_cta, c_map_nom, c_map_tipo, c_map_est])
                        
                        if c_map_nom is None:
                            df_map["NOMBRE DE CUENTA"] = df_map[c_map_est] if c_map_est else df_map[c_map_cta]
                            c_map_nom = "NOMBRE DE CUENTA"

                        df_map[c_map_cta] = df_map[c_map_cta].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                        df_map[c_map_tipo] = df_map[c_map_tipo].apply(limpiar_texto)
                        
                        df_map_valido = df_map.copy() 
                        
                        dfs_cons = []
                        for arch in archivos_subidos:
                            df_temp = cargar_archivo_contable(arch)

                            c_arch_cta = buscar_columna(df_temp, ["IDCUENTA", "CUENTA", "CODIGO", "ID"])
                            # Exclusión estricta de palabras INICIAL y ANTERIOR
                            c_arch_sld = buscar_columna(df_temp, ["SALDOFINAL", "FINAL", "SALDO_FINAL", "TOTALFINAL"], anti_palabras=["INICIAL", "ANTERIOR"], excluir=[c_arch_cta])
                            
                            if not c_arch_sld:
                                c_arch_sld = buscar_columna(df_temp, ["SALDO"], anti_palabras=["INICIAL", "ANTERIOR"], excluir=[c_arch_cta])
                                
                            if c_arch_sld and c_arch_cta:
                                df_temp[c_arch_cta] = df_temp[c_arch_cta].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                                df_temp[c_arch_sld] = pd.to_numeric(df_temp[c_arch_sld].astype(str).str.replace(r'[^\d\.\-]', '', regex=True), errors='coerce').fillna(0)
                                df_agrup = df_temp.groupby(c_arch_cta, as_index=False)[c_arch_sld].sum()
                                df_agrup.rename(columns={c_arch_sld: "SALDO_FINAL_GLOBAL"}, inplace=True)
                                dfs_cons.append(df_agrup)
                            else:
                                st.warning(f"⚠️ El archivo {arch.name} no posee encabezados de 'SALDO FINAL' identificables.")

                        if dfs_cons:
                            df_total = pd.concat(dfs_cons, ignore_index=True)
                            df_total_agrup = df_total.groupby(c_arch_cta, as_index=False)["SALDO_FINAL_GLOBAL"].sum()
                            df_final = pd.merge(df_map_valido, df_total_agrup, left_on=c_map_cta, right_on=c_arch_cta, how="inner")
                            
                            st.session_state['cont_df'] = df_final
                            st.session_state['cols_dict'] = {'cta': c_map_cta, 'nom': c_map_nom, 'tipo': c_map_tipo, 'est': c_map_est, 'cat': c_map_cat, 'sld': "SALDO_FINAL_GLOBAL"}
                            st.session_state['sync_flag'] = True 
                            st.success("✅ Procesado para Simulador Global.")
                        else:
                            st.error("❌ Ningún archivo pudo ser procesado. Verifica el formato de la exportación.")
                    except Exception as e:
                        st.error(f"🚨 Ocurrió un error técnico general: {e}")

    # ==========================================
    # PESTAÑA 2: SIMULADOR GLOBAL
    # ==========================================
    with tab_pe:
        if 'cont_df' in st.session_state:
            df = st.session_state['cont_df']
            cd = st.session_state['cols_dict']
            if 'df_master' not in st.session_state or st.session_state.get('sync_flag', False):
                st.session_state['df_master'] = df[[cd['cta'], cd['nom'], cd['est'], cd['cat'], cd['tipo'], cd['sld']]].copy()
                st.session_state['sync_flag'] = False
                
            st.subheader("🎛️ Simulador Financiero Rápido")
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                filtro_tipo = st.selectbox("A. Filtrar por Tipo:", options=["Todos"] + st.session_state['df_master'][cd['tipo']].unique().tolist())
            df_filtro = st.session_state['df_master'].copy()
            if filtro_tipo != "Todos": df_filtro = df_filtro[df_filtro[cd['tipo']] == filtro_tipo]
                
            with col_f2:
                filtro_cat = st.selectbox("B. Filtrar por Categoría:", options=["Todas"] + df_filtro[cd['cat']].unique().tolist())
            if filtro_cat != "Todas": df_filtro = df_filtro[df_filtro[cd['cat']] == filtro_cat]
                
            with col_f3:
                df_filtro['display_name'] = df_filtro[cd['cta']].astype(str) + " - " + df_filtro[cd['nom']].astype(str)
                cuenta_sel = st.selectbox("C. Seleccionar Cuenta:", options=["Seleccione una cuenta..."] + df_filtro['display_name'].tolist())

            if cuenta_sel != "Seleccione una cuenta...":
                id_cuenta_sel = cuenta_sel.split(" - ")[0]
                idx = st.session_state['df_master'].index[st.session_state['df_master'][cd['cta']].astype(str) == id_cuenta_sel].tolist()[0]
                col_ed1, col_ed2, col_ed3 = st.columns([2, 2, 1])
                with col_ed1:
                    t_act = st.session_state['df_master'].at[idx, cd['tipo']]
                    opts = ["COSTO FIJO", "COSTO VARIABLE", "INGRESOS", "NO APLICA", "CUENTA DE MAYOR"]
                    if t_act not in opts: opts.append(t_act)
                    n_tipo = st.selectbox("Tipo:", options=opts, index=opts.index(t_act))
                with col_ed2:
                    n_monto = st.number_input("Monto ($):", value=float(st.session_state['df_master'].at[idx, cd['sld']]), min_value=0.0)
                with col_ed3:
                    st.write(""); st.write("")
                    if st.button("✅ Aplicar", use_container_width=True):
                        st.session_state['df_master'].at[idx, cd['tipo']] = n_tipo
                        st.session_state['df_master'].at[idx, cd['sld']] = n_monto
                        st.rerun()
            
            df_m = st.session_state['df_master']
            v = df_m[df_m[cd['tipo']].isin(["INGRESOS", "INGRESO"])][cd['sld']].abs().sum()
            cf = df_m[df_m[cd['tipo']] == "COSTO FIJO"][cd['sld']].abs().sum()
            cv = df_m[df_m[cd['tipo']] == "COSTO VARIABLE"][cd['sld']].abs().sum()
            m_pct = (1 - (cv / v)) if v > 0 else 0
            pe = (cf / m_pct) if m_pct > 0 else 0
            
            c1, c2, c3 = st.columns(3)
            c1.metric("Ingresos", f"${v:,.2f}"); c2.metric("CV", f"${cv:,.2f}"); c3.metric("CF", f"${cf:,.2f}")
            st.write(f"**Punto Equilibrio:** ${pe:,.2f} | **Margen:** {m_pct*100:.2f}%")
        else: st.info("Sincroniza en Pestaña 1.")

    # ==========================================
    # PESTAÑA 3: DASHBOARD
    # ==========================================
    with tab_dash:
        if 'df_master' in st.session_state:
            df_g = st.session_state['df_master']
            df_g = df_g[df_g[st.session_state['cols_dict']['tipo']].isin(["COSTO FIJO", "COSTO VARIABLE"])]
            if not df_g.empty:
                c1, c2 = st.columns(2)
                with c1: st.plotly_chart(px.pie(df_g, values=st.session_state['cols_dict']['sld'], names=st.session_state['cols_dict']['tipo']), use_container_width=True)
                with c2: 
                    df_c = df_g.groupby(st.session_state['cols_dict']['cat'], as_index=False)[st.session_state['cols_dict']['sld']].sum()
                    st.plotly_chart(px.bar(df_c, x=st.session_state['cols_dict']['cat'], y=st.session_state['cols_dict']['sld']), use_container_width=True)
        else: st.info("Sincroniza en Pestaña 1.")

    # ==========================================
    # PESTAÑA 4: MATRIZ 2025 (HISTÓRICA)
    # ==========================================
    with tab_matriz:
        st.subheader("📑 Reporte Histórico: Matriz por Unidades")
        archivos_matriz = st.file_uploader("Sube archivos Excel por unidad:", type=["xlsx", "xls"], accept_multiple_files=True, key="up_m4")
        mapeo_m4 = {}
        if archivos_matriz:
            for arch in archivos_matriz:
                mapeo_m4[arch.name] = {"file": arch, "unidad": st.selectbox(f"Unidad para '{arch.name}':", unidades_oficiales, key=f"s4_{arch.name}")}
                
            if st.button("🔄 Armar Matriz Histórica", type="primary"):
                with st.spinner("Construyendo matriz y prorrateando Gerencias..."):
                    try:
                        df_map_m = obtener_dataframe("Balance_Mapeado")
                        df_map_m.columns = df_map_m.columns.astype(str).str.strip().str.upper()
                        df_map_m = df_map_m.loc[:, ~df_map_m.columns.duplicated()].copy()
                        
                        cm_cta = buscar_columna(df_map_m, ["IDCUENTA", "CUENTA", "CODIGO", "ID"])
                        cm_nom = buscar_columna(df_map_m, ["NOMBRE", "DESCRIP"], excluir=[cm_cta]) 
                        cm_tipo = buscar_columna(df_map_m, ["TIPO", "CLASIFICACION"], excluir=[cm_cta, cm_nom])
                        cm_est = buscar_columna(df_map_m, ["ESTADO", "GRUPO"], excluir=[cm_cta, cm_nom, cm_tipo])
                        
                        if cm_nom is None:
                            df_map_m["NOMBRE DE CUENTA"] = df_map_m[cm_est] if cm_est else df_map_m[cm_cta]
                            cm_nom = "NOMBRE DE CUENTA"

                        df_map_m[cm_cta] = df_map_m[cm_cta].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                        df_map_m[cm_tipo] = df_map_m[cm_tipo].apply(limpiar_texto)
                        
                        df_base_m4 = df_map_m.copy()
                        
                        for u in unidades_oficiales: df_base_m4[u] = 0.0

                        for nombre_arch, info in mapeo_m4.items():
                            df_tm = cargar_archivo_contable(info["file"])
                            
                            c_a_cta = buscar_columna(df_tm, ["IDCUENTA", "CUENTA", "CODIGO", "ID"])
                            # Exclusión estricta de palabras INICIAL y ANTERIOR
                            c_a_sld = buscar_columna(df_tm, ["SALDOFINAL", "FINAL", "SALDO_FINAL", "TOTALFINAL"], anti_palabras=["INICIAL", "ANTERIOR"], excluir=[c_a_cta])
                            
                            if not c_a_sld:
                                c_a_sld = buscar_columna(df_tm, ["SALDO"], anti_palabras=["INICIAL", "ANTERIOR"], excluir=[c_a_cta])
                            
                            if c_a_sld and c_a_cta:
                                df_tm[c_a_cta] = df_tm[c_a_cta].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                                df_tm[c_a_sld] = pd.to_numeric(df_tm[c_a_sld].astype(str).str.replace(r'[^\d\.\-]', '', regex=True), errors='coerce').fillna(0)
                                df_ag = df_tm.groupby(c_a_cta, as_index=False)[c_a_sld].sum()
                                
                                u_obj = info["unidad"]
                                for _, row in df_ag.iterrows():
                                    idx_c = df_base_m4[df_base_m4[cm_cta] == row[c_a_cta]].index
                                    if not idx_c.empty: df_base_m4.loc[idx_c, u_obj] += row[c_a_sld]

                        df_base_m4['TOTAL CONSOLIDADO'] = df_base_m4[unidades_oficiales].sum(axis=1)
                        st.session_state['matriz_2025_cruda'] = df_base_m4.copy()
                        
                        mask_excluir = (
                            df_base_m4[cm_tipo].str.upper().isin(["NO APLICA", ""]) |
                            df_base_m4[cm_est].str.upper().isin(["CUENTA DE MAYOR", "CUENTA DE MENOR", "CUENTA GENERAL"])
                        )
                        df_base_prorrateo = df_base_m4[~mask_excluir].copy()
                        
                        df_prorrateada, pct = procesar_prorrateo_matriz(df_base_prorrateo, cm_cta, cm_nom, cm_est, cm_tipo, unidades_oficiales)
                        st.session_state['matriz_2025_prorrateada'] = df_prorrateada
                        st.session_state['cols_matriz'] = {'cta': cm_cta, 'nom': cm_nom, 'est': cm_est, 'tipo': cm_tipo}
                        
                        st.success("✅ Matriz 2025 generada de manera exitosa.")
                    except Exception as e:
                        st.error(f"🚨 Ocurrió un error en el armado de la matriz: {e}")

        if 'matriz_2025_prorrateada' in st.session_state:
            df_m4_disp = st.session_state['matriz_2025_prorrateada'].copy()
            for col in unidades_oficiales + ['TOTAL CONSOLIDADO']:
                df_m4_disp[col] = df_m4_disp[col].apply(lambda x: f"{x:,.2f}" if pd.notnull(x) and isinstance(x, (int, float)) else x)
            st.dataframe(df_m4_disp, use_container_width=True, hide_index=True)
            
            dict_export_2025 = {
                "Matriz Cruda": st.session_state['matriz_2025_cruda'],
                "Prorrateo y PE": st.session_state['matriz_2025_prorrateada']
            }
            
            st.download_button(
                label="📥 Descargar Reporte Completo 2025 (Excel)", 
                data=generar_excel_multi(dict_export_2025), 
                file_name='Matriz_Directiva_2025.xlsx',
                use_container_width=True
            )

    # ==========================================
    # PESTAÑA 5: PROYECCIONES 2026
    # ==========================================
    with tab_proy:
        st.subheader("🚀 Simulador de Proyecciones (Fases)")
        if 'matriz_2025_cruda' not in st.session_state:
            st.warning("⚠️ Primero debes armar la Matriz Histórica en la Pestaña 4.")
        else:
            if 'lista_ajustes' not in st.session_state:
                st.session_state['lista_ajustes'] = []

            with st.expander("➕ Añadir Ajuste (Micro o Macro)", expanded=True):
                ca1, ca2, ca3, ca4 = st.columns(4)
                with ca1:
                    unidad_ajuste = st.selectbox("1. Unidad a afectar:", ["Todas las operativas (Prorrateo)"] + unidades_oficiales)
                with ca2:
                    df_cmb = st.session_state['matriz_2025_cruda']
                    c_cta = st.session_state['cols_matriz']['cta']
                    c_nom = st.session_state['cols_matriz']['nom']
                    list_ctas = (df_cmb[c_cta].astype(str) + " - " + df_cmb[c_nom]).tolist()
                    cuenta_ajuste = st.selectbox("2. Cuenta:", list_ctas)
                with ca3:
                    tipo_ajuste = st.selectbox("3. Tipo de Ajuste:", ["Monto Fijo ($)", "Porcentaje (%)"])
                with ca4:
                    val_ajuste = st.number_input("4. Valor (+ o -):", value=0.0, format="%.2f")
                
                if st.button("Añadir a Lista de Espera"):
                    st.session_state['lista_ajustes'].append({
                        "Unidad": unidad_ajuste, "Cuenta": cuenta_ajuste.split(" - ")[0], "Nombre": cuenta_ajuste.split(" - ")[1],
                        "Tipo": tipo_ajuste, "Valor": val_ajuste
                    })
                    st.rerun()

            if st.session_state['lista_ajustes']:
                st.markdown("##### Ajustes en Cola")
                df_ajustes = pd.DataFrame(st.session_state['lista_ajustes'])
                st.table(df_ajustes)
                if st.button("Limpiar Ajustes"):
                    st.session_state['lista_ajustes'] = []
                    st.rerun()

            if st.button("⚙️ Calcular Proyección 2026", type="primary", use_container_width=True):
                with st.spinner("Fase 1: Ajustes Directos | Fase 2: Recálculo de Base | Fase 3: Prorrateo..."):
                    df_proy_cruda = st.session_state['matriz_2025_cruda'].copy()
                    c_cta = st.session_state['cols_matriz']['cta']
                    
                    for aj in st.session_state['lista_ajustes']:
                        id_c = aj["Cuenta"]
                        val = aj["Valor"]
                        idx_c = df_proy_cruda[df_proy_cruda[c_cta] == id_c].index
                        
                        if not idx_c.empty:
                            if aj["Unidad"] == "Todas las operativas (Prorrateo)":
                                mask_temp = (
                                    df_proy_cruda[st.session_state['cols_matriz']['tipo']].str.upper().isin(["NO APLICA", ""]) |
                                    df_proy_cruda[st.session_state['cols_matriz']['est']].str.upper().isin(["CUENTA DE MAYOR", "CUENTA DE MENOR", "CUENTA GENERAL"])
                                )
                                df_temp_prorrateo = df_proy_cruda[~mask_temp].copy()
                                
                                _, pct_temp = procesar_prorrateo_matriz(
                                    df_temp_prorrateo, c_cta, st.session_state['cols_matriz']['nom'], 
                                    st.session_state['cols_matriz']['est'], st.session_state['cols_matriz']['tipo'], unidades_oficiales
                                )
                                for u, p in pct_temp.items():
                                    if aj["Tipo"] == "Monto Fijo ($)": df_proy_cruda.loc[idx_c, u] += (val * p)
                                    else: df_proy_cruda.loc[idx_c, u] *= (1 + (val/100))
                            else:
                                u = aj["Unidad"]
                                if aj["Tipo"] == "Monto Fijo ($)": df_proy_cruda.loc[idx_c, u] += val
                                else: df_proy_cruda.loc[idx_c, u] *= (1 + (val/100))

                    df_proy_cruda['TOTAL CONSOLIDADO'] = df_proy_cruda[unidades_oficiales].sum(axis=1)
                    st.session_state['matriz_2026_cruda'] = df_proy_cruda
                    
                    mask_excluir_26 = (
                        df_proy_cruda[st.session_state['cols_matriz']['tipo']].str.upper().isin(["NO APLICA", ""]) |
                        df_proy_cruda[st.session_state['cols_matriz']['est']].str.upper().isin(["CUENTA DE MAYOR", "CUENTA DE MENOR", "CUENTA GENERAL"])
                    )
                    df_proy_prorrateo = df_proy_cruda[~mask_excluir_26].copy()
                    
                    df_proy_prorrateada, _ = procesar_prorrateo_matriz(
                        df_proy_prorrateo, c_cta, st.session_state['cols_matriz']['nom'], 
                        st.session_state['cols_matriz']['est'], st.session_state['cols_matriz']['tipo'], unidades_oficiales
                    )
                    st.session_state['matriz_2026_prorrateada'] = df_proy_prorrateada
                    st.success("✅ Proyección 2026 completada.")

            if 'matriz_2026_prorrateada' in st.session_state:
                df_p5_disp = st.session_state['matriz_2026_prorrateada'].copy()
                for col in unidades_oficiales + ['TOTAL CONSOLIDADO']:
                    df_p5_disp[col] = df_p5_disp[col].apply(lambda x: f"{x:,.2f}" if pd.notnull(x) and isinstance(x, (int, float)) else x)
                st.dataframe(df_p5_disp, use_container_width=True, hide_index=True)
                
                dict_export_2026 = {
                    "Matriz Cruda 2026": st.session_state['matriz_2026_cruda'],
                    "Prorrateo y PE 2026": st.session_state['matriz_2026_prorrateada']
                }
                
                st.download_button(
                    label="📥 Descargar Proyección 2026 (Excel)", 
                    data=generar_excel_multi(dict_export_2026), 
                    file_name='Proyeccion_Directiva_2026.xlsx',
                    use_container_width=True
                )

    # ==========================================
    # PESTAÑA 6: COMPARATIVO 25 vs 26
    # ==========================================
    with tab_comp:
        st.subheader("⚖️ Análisis Horizontal y Vertical (2025 vs 2026)")
        
        if 'matriz_2025_cruda' in st.session_state and 'matriz_2026_cruda' in st.session_state:
            unidad_comp = st.selectbox("Seleccionar Vista:", ["TOTAL CONSOLIDADO"] + unidades_oficiales)
            
            if st.button("Generar Reporte Comparativo"):
                df_25 = st.session_state['matriz_2025_cruda']
                df_26 = st.session_state['matriz_2026_cruda']
                c_cta = st.session_state['cols_matriz']['cta']
                c_nom = st.session_state['cols_matriz']['nom']
                c_est = st.session_state['cols_matriz']['est']
                c_tipo = st.session_state['cols_matriz']['tipo']
                
                ingresos_25 = df_25[df_25[c_tipo].isin(["INGRESOS", "INGRESO"])][unidad_comp].sum()
                ingresos_26 = df_26[df_26[c_tipo].isin(["INGRESOS", "INGRESO"])][unidad_comp].sum()
                
                df_comp = pd.DataFrame({
                    "ID CUENTA": df_25[c_cta],
                    "NOMBRE CUENTA": df_25[c_nom],
                    "ESTADO": df_25[c_est],
                    "TIPO": df_25[c_tipo],
                    "MONTO 2025 ($)": df_25[unidad_comp],
                    "MONTO 2026 ($)": df_26[unidad_comp]
                })
                
                df_comp["AV 2025 (%)"] = (df_comp["MONTO 2025 ($)"] / ingresos_25 * 100) if ingresos_25 > 0 else 0
                df_comp["AV 2026 (%)"] = (df_comp["MONTO 2026 ($)"] / ingresos_26 * 100) if ingresos_26 > 0 else 0
                df_comp["VARIACIÓN ($)"] = df_comp["MONTO 2026 ($)"] - df_comp["MONTO 2025 ($)"]
                df_comp["VARIACIÓN (%)"] = (df_comp["VARIACIÓN ($)"] / df_comp["MONTO 2025 ($)"].replace(0, pd.NA) * 100).fillna(0)
                
                st.session_state['df_comparativo'] = df_comp
                st.success(f"Reporte generado para: {unidad_comp}")

        if 'df_comparativo' in st.session_state:
            df_out = st.session_state['df_comparativo'].copy()
            
            cols_dinero = ["MONTO 2025 ($)", "MONTO 2026 ($)", "VARIACIÓN ($)"]
            cols_pct = ["AV 2025 (%)", "AV 2026 (%)", "VARIACIÓN (%)"]
            for c in cols_dinero: df_out[c] = df_out[c].apply(lambda x: f"${x:,.2f}" if pd.notnull(x) else "")
            for c in cols_pct: df_out[c] = df_out[c].apply(lambda x: f"{x:,.2f}%" if pd.notnull(x) else "")
            
            st.dataframe(df_out, use_container_width=True, hide_index=True)
            
            st.download_button(
                "📥 Descargar Comparativo (Excel)", 
                data=generar_excel_multi({"Análisis Comparativo": st.session_state['df_comparativo']}), 
                file_name='Analisis_Comparativo.xlsx'
            )
        elif 'matriz_2025_cruda' not in st.session_state or 'matriz_2026_cruda' not in st.session_state:
            st.info("👈 Debes generar la Matriz 2025 (Pestaña 4) y la Proyección 2026 (Pestaña 5) antes de comparar.")