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
        "CAFETERIA": ["TERRAZA", "CENTRO SOHO", "CAFETERIA CENTRAL", "CAFETERIA", "CAFETERIA ABASTECIMIENTO"],
        "DESPENSA":  ["DESPENSA"]
    }

    # LA LISTA NEGRA: Destinos que el sistema ignorará por completo
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

    def cargar_y_marcar(archivo):
        df = pd.read_excel(archivo, dtype=str)
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
                if all(p in c_norm for p in k): return pd.to_numeric(df[col], errors='coerce').fillna(0.0)
        return pd.Series(0.0, index=df.index)

    def proteger_cuentas_nulas(df_m, fallback="SIN CUENTA REGISTRADA"):
        if 'Cuenta_Contable' in df_m.columns:
            df_m['Cuenta_Contable'] = df_m['Cuenta_Contable'].fillna(fallback)
            mask = df_m['Cuenta_Contable'].astype(str).str.strip().str.upper().isin(["", "NAN", "NAT", "NONE"])
            df_m.loc[mask, 'Cuenta_Contable'] = fallback
        return df_m

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
            traslados_neta = 0.0
            f_t_in = None; f_t_out = None
            if not df_hist_tras.empty:
                df_hist_tras['Monto'] = pd.to_numeric(df_hist_tras['Monto'], errors='coerce').fillna(0.0)
                
                if es_consolidado:
                    meses_indices = [list(meses_texto.keys())[list(meses_texto.values()).index(m)] for m in nombres_meses]
                    filtro_base = (pd.to_numeric(df_hist_tras['Mes'], errors='coerce').isin(meses_indices)) & \
                                  (pd.to_numeric(df_hist_tras['Año'], errors='coerce') == anio_cierre)
                else:
                    filtro_base = (pd.to_numeric(df_hist_tras['Mes'], errors='coerce') == mes_cierre) & \
                                  (pd.to_numeric(df_hist_tras['Año'], errors='coerce') == anio_cierre)
                
                mask_dest = df_hist_tras['Destino'].apply(lambda x: es_de_unidad(x, unidad_cierre))
                mask_orig = df_hist_tras['Origen'].apply(lambda x: es_de_unidad(x, unidad_cierre))
                
                f_t_in = filtro_base & mask_dest & (~mask_orig)
                f_t_out = filtro_base & mask_orig & (~mask_dest)
                
                monto_in = df_hist_tras[f_t_in]['Monto'].sum()
                monto_out = df_hist_tras[f_t_out]['Monto'].sum()
                traslados_neta = monto_in - monto_out

            porcentaje_subsidio = (subsidio_mes / ventas_mes) if ventas_mes > 0 else 0.0
            st.info(f"📊 Ingresos Totales Periodo: Ventas ${ventas_mes:,.2f} | Subsidio ${subsidio_mes:,.2f} | Traslados Netos (Entran - Salen): ${traslados_neta:,.2f}")

            costo_diferido_anterior = st.number_input("Costo Diferido de Arrastre (110602):", min_value=0.0, value=0.0)
            
            col_u1, col_u2, col_u3 = st.columns(3)
            with col_u1: arch_ini = st.file_uploader("1. Inv. Inicial", type=["xlsx"], accept_multiple_files=True)
            with col_u2: arch_com = st.file_uploader("2. Compras", type=["xlsx"], accept_multiple_files=True)
            with col_u3: arch_fin = st.file_uploader("3. Inv. Final", type=["xlsx"], accept_multiple_files=True)

            if arch_ini and arch_com and arch_fin:
                if 'huerfanos_df' not in st.session_state:
                    forzar_calculo = st.checkbox("⚠️ Forzar cálculo ciego (Omitir revisión de cuentas)")
                    
                    if st.button("⚙️ Procesar Archivos y Guardar en Memoria", type="primary", use_container_width=True):
                        with st.spinner("Realizando validación cruzada y buscando anomalías..."):
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

                                def detectar_huerfanos(df):
                                    if df.empty: return pd.DataFrame()
                                    mask = df['Cuenta_Contable'].isna() | df['Cuenta_Contable'].astype(str).str.strip().str.upper().isin(["", "NAN", "NAT", "NONE"])
                                    return df[mask][['Codigo', 'Categoria', 'ORIGEN_ARCHIVO']]

                                df_faltantes = pd.concat([detectar_huerfanos(df_ini_m), detectar_huerfanos(df_com_m), detectar_huerfanos(df_fin_m)]).drop_duplicates(subset=['Codigo'])

                                if not df_hist_tras.empty:
                                    grp_tras_in = df_hist_tras[f_t_in].groupby('Cuenta_Contable')['Monto'].sum()
                                    grp_tras_out = df_hist_tras[f_t_out].groupby('Cuenta_Contable')['Monto'].sum()
                                else:
                                    grp_tras_in = pd.Series(dtype=float); grp_tras_out = pd.Series(dtype=float)

                                if not df_faltantes.empty and not forzar_calculo:
                                    st.session_state['pre_proceso'] = {
                                        'ini': df_ini_m, 'com': df_com_m, 'fin': df_fin_m, 
                                        'grp_tras_in': grp_tras_in, 'grp_tras_out': grp_tras_out
                                    }
                                    st.session_state['huerfanos_df'] = df_faltantes
                                    st.rerun()

                                df_ini_m = proteger_cuentas_nulas(df_ini_m)
                                df_com_m = proteger_cuentas_nulas(df_com_m)
                                df_fin_m = proteger_cuentas_nulas(df_fin_m)

                                df_ini_m['Valor'] = pd.to_numeric(get_num(df_ini_m, [['EXISTENCIAS'], ['SALDO']]), errors='coerce').fillna(0.0) * pd.to_numeric(get_num(df_ini_m, [['COSTO', 'U']]), errors='coerce').fillna(0.0)
                                df_fin_m['Valor'] = pd.to_numeric(get_num(df_fin_m, [['EXISTENCIAS'], ['SALDO']]), errors='coerce').fillna(0.0) * pd.to_numeric(get_num(df_fin_m, [['COSTO', 'U']]), errors='coerce').fillna(0.0)
                                df_com_m['Valor'] = pd.to_numeric(get_num(df_com_m, [['TOTAL'], ['MONTO']]), errors='coerce').fillna(0.0)

                                grp_ini = df_ini_m.groupby('Cuenta_Contable')['Valor'].sum()
                                grp_comp = df_com_m.groupby('Cuenta_Contable')['Valor'].sum()
                                grp_fin = df_fin_m.groupby('Cuenta_Contable')['Valor'].sum()
                                
                                todas_cuentas = set(grp_ini.index).union(grp_comp.index).union(grp_fin.index).union(grp_tras_in.index).union(grp_tras_out.index)
                                consumo_por_cuenta = {}
                                total_ini_val = total_com_val = total_tras_val = total_fin_val = costo_operativo = 0.0
                                
                                cuentas_invalidas = ["NO APLICA", "0", "0.0", "OMITIDO_MANUAL"]
                                for cta in todas_cuentas:
                                    if pd.isna(cta): continue
                                    cta_str = str(cta).strip().upper().replace(".0", "")
                                    
                                    v_ini = float(grp_ini.get(cta, 0.0))
                                    v_com = float(grp_comp.get(cta, 0.0))
                                    v_tras_entrada = float(grp_tras_in.get(cta, 0.0))
                                    v_tras_salida = float(grp_tras_out.get(cta, 0.0))
                                    v_fin = float(grp_fin.get(cta, 0.0))
                                    
                                    v_tras_neta = v_tras_entrada - v_tras_salida
                                    
                                    if cta_str not in cuentas_invalidas: 
                                        total_ini_val += v_ini; total_com_val += v_com
                                        total_tras_val += v_tras_neta; total_fin_val += v_fin
                                        
                                        val = v_ini + v_com + v_tras_neta - v_fin
                                        if val != 0:
                                            consumo_por_cuenta[cta_str] = val
                                            costo_operativo += val
                                
                                costo_dif_mes = float(costo_operativo) * float(porcentaje_subsidio)
                                costo_real = float(costo_operativo) - float(costo_dif_mes) + float(costo_diferido_anterior)

                                st.session_state['memoria_cierre'] = {
                                    'df_ini_m': df_ini_m, 'df_com_m': df_com_m, 'df_fin_m': df_fin_m,
                                    'grp_ini_sum': total_ini_val, 'grp_comp_sum': total_com_val, 
                                    'grp_tras_sum': total_tras_val, 'grp_fin_sum': total_fin_val,
                                    'costo_dif_mes': costo_dif_mes, 'costo_real': costo_real, 'costo_operativo': costo_operativo,
                                    'costo_diferido_anterior': costo_diferido_anterior, 'consumo_por_cuenta': consumo_por_cuenta,
                                    'mes_cierre': mes_cierre, 'anio_cierre': anio_cierre, 'unidad_cierre': unidad_cierre,
                                    'es_consolidado': es_consolidado, 'pesos': pesos, 'nombres_meses': nombres_meses
                                }
                                st.session_state['datos_auditoria'] = {'consumo': consumo_por_cuenta, 'ventas': ventas_mes, 'costo_real': costo_real}
                                st.rerun() 
                            except Exception as e: st.error(f"Error procesando: {e}")

                else:
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
                            "ORIGEN_ARCHIVO": "Archivo Origen",
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
                            df_ini_m = st.session_state['pre_proceso']['ini']
                            df_com_m = st.session_state['pre_proceso']['com']
                            df_fin_m = st.session_state['pre_proceso']['fin']
                            grp_tras_in = st.session_state['pre_proceso']['grp_tras_in']
                            grp_tras_out = st.session_state['pre_proceso']['grp_tras_out']

                            codigos_omitir = edited_df[edited_df['Accion'] == 'Omitir (No sumar al costo)']['Codigo'].tolist()
                            df_asignar = edited_df[edited_df['Accion'] == 'Escribir Cuenta Manual']
                            df_categoria = edited_df[edited_df['Accion'] == 'Usar Categoría Nativa']

                            def aplicar_reglas_auditor(df_obj):
                                df = df_obj.copy()
                                mask_omitir = df['Codigo'].isin(codigos_omitir)
                                if mask_omitir.any(): df.loc[mask_omitir, 'Cuenta_Contable'] = 'OMITIDO_MANUAL'
                                
                                for _, row in df_asignar.iterrows():
                                    c_manual = str(row['Cuenta_Manual']).strip()
                                    if c_manual == "": c_manual = "SIN CUENTA REGISTRADA"
                                    df.loc[df['Codigo'] == row['Codigo'], 'Cuenta_Contable'] = c_manual
                                
                                for _, row in df_categoria.iterrows():
                                    cat_val = str(row['Categoria']).strip()
                                    df.loc[df['Codigo'] == row['Codigo'], 'Cuenta_Contable'] = cat_val
                                
                                return df

                            df_ini_m = aplicar_reglas_auditor(df_ini_m)
                            df_com_m = aplicar_reglas_auditor(df_com_m)
                            df_fin_m = aplicar_reglas_auditor(df_fin_m)

                            df_ini_m = proteger_cuentas_nulas(df_ini_m)
                            df_com_m = proteger_cuentas_nulas(df_com_m)
                            df_fin_m = proteger_cuentas_nulas(df_fin_m)

                            df_ini_m['Valor'] = pd.to_numeric(get_num(df_ini_m, [['EXISTENCIAS'], ['SALDO']]), errors='coerce').fillna(0.0) * pd.to_numeric(get_num(df_ini_m, [['COSTO', 'U']]), errors='coerce').fillna(0.0)
                            df_fin_m['Valor'] = pd.to_numeric(get_num(df_fin_m, [['EXISTENCIAS'], ['SALDO']]), errors='coerce').fillna(0.0) * pd.to_numeric(get_num(df_fin_m, [['COSTO', 'U']]), errors='coerce').fillna(0.0)
                            df_com_m['Valor'] = pd.to_numeric(get_num(df_com_m, [['TOTAL'], ['MONTO']]), errors='coerce').fillna(0.0)

                            grp_ini = df_ini_m.groupby('Cuenta_Contable')['Valor'].sum()
                            grp_comp = df_com_m.groupby('Cuenta_Contable')['Valor'].sum()
                            grp_fin = df_fin_m.groupby('Cuenta_Contable')['Valor'].sum()
                            
                            todas_cuentas = set(grp_ini.index).union(grp_comp.index).union(grp_fin.index).union(grp_tras_in.index).union(grp_tras_out.index)
                            consumo_por_cuenta = {}
                            total_ini_val = total_com_val = total_tras_val = total_fin_val = costo_operativo = 0.0
                            
                            cuentas_invalidas = ["NO APLICA", "0", "0.0", "OMITIDO_MANUAL"]
                            for cta in todas_cuentas:
                                if pd.isna(cta): continue
                                cta_str = str(cta).strip().upper().replace(".0", "")
                                
                                v_ini = float(grp_ini.get(cta, 0.0))
                                v_com = float(grp_comp.get(cta, 0.0))
                                v_tras_entrada = float(grp_tras_in.get(cta, 0.0))
                                v_tras_salida = float(grp_tras_out.get(cta, 0.0))
                                v_fin = float(grp_fin.get(cta, 0.0))
                                
                                v_tras_neta = v_tras_entrada - v_tras_salida
                                
                                if cta_str not in cuentas_invalidas: 
                                    total_ini_val += v_ini; total_com_val += v_com
                                    total_tras_val += v_tras_neta; total_fin_val += v_fin
                                    val = v_ini + v_com + v_tras_neta - v_fin
                                    if val != 0:
                                        consumo_por_cuenta[cta_str] = val
                                        costo_operativo += val
                            
                            costo_dif_mes = float(costo_operativo) * float(porcentaje_subsidio)
                            costo_real = float(costo_operativo) - float(costo_dif_mes) + float(costo_diferido_anterior)

                            st.session_state['memoria_cierre'] = {
                                'df_ini_m': df_ini_m, 'df_com_m': df_com_m, 'df_fin_m': df_fin_m,
                                'grp_ini_sum': total_ini_val, 'grp_comp_sum': total_com_val, 
                                'grp_tras_sum': total_tras_val, 'grp_fin_sum': total_fin_val,
                                'costo_dif_mes': costo_dif_mes, 'costo_real': costo_real, 'costo_operativo': costo_operativo,
                                'costo_diferido_anterior': costo_diferido_anterior, 'consumo_por_cuenta': consumo_por_cuenta,
                                'mes_cierre': mes_cierre, 'anio_cierre': anio_cierre, 'unidad_cierre': unidad_cierre,
                                'es_consolidado': es_consolidado, 'pesos': pesos, 'nombres_meses': nombres_meses
                            }
                            st.session_state['datos_auditoria'] = {'consumo': consumo_por_cuenta, 'ventas': ventas_mes, 'costo_real': costo_real}
                            
                            del st.session_state['huerfanos_df']
                            del st.session_state['pre_proceso']
                            st.rerun()

                    if col_b2.button("❌ Cancelar y Limpiar Archivos"):
                        del st.session_state['huerfanos_df']
                        del st.session_state['pre_proceso']
                        st.rerun()

        else:
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
     # PESTAÑA 2: REGISTRO DE TRASLADOS (SELECTOR DE INGRESOS/SALIDAS)
    # =========================================================================
    with tab2:
        st.subheader("🚚 Registro Automático de Traslados Nexus")

        col_t1, col_t2, col_t3 = st.columns(3)
        with col_t1:
            mes_reg = st.selectbox("Mes:", range(1, 13), index=date.today().month - 1, key="mt_reg")
        with col_t2:
            anio_reg = st.number_input("Año:", min_value=2024, value=2026, key="at_reg")
        with col_t3:
            u_responsable = st.selectbox("Módulo Responsable:", ["CAFETERIA", "DESPENSA"], key="uni_reg")

        # NUEVO SELECTOR DE TIPO DE MOVIMIENTO
        tipo_movimiento = st.radio("¿Qué tipo de traslado vas a registrar?", ["Ingresos (Entradas a mi unidad)", "Salidas (Envíos desde mi unidad)"], horizontal=True)

        st.info(f"💡 Filtro Activo: Procesando **{tipo_movimiento.split(' ')[0]}** para {u_responsable}.")

        archivo_nexus = st.file_uploader("Reporte Nexus (I:Cod, K:Cant, N:Monto, AA:Cat, AB:Origen, AC:Destino)", type=["xlsx"], key="atf_reg")

        if archivo_nexus:
            try:
                df_raw_t = pd.read_excel(archivo_nexus, usecols="I,K,N,AA,AB,AC", header=0, names=['Codigo', 'Cantidad', 'Monto', 'Categoria', 'Origen', 'Destino'], dtype=str)
                df_raw_t['Codigo'] = df_raw_t['Codigo'].astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
                df_raw_t['Monto'] = pd.to_numeric(df_raw_t['Monto'], errors='coerce').fillna(0.0)
                df_raw_t['Cantidad'] = pd.to_numeric(df_raw_t['Cantidad'], errors='coerce').fillna(0.0)

                condicion_destino = df_raw_t['Destino'].apply(lambda x: es_de_unidad(x, u_responsable))
                condicion_origen = df_raw_t['Origen'].apply(lambda x: es_de_unidad(x, u_responsable))

                # LÓGICA SEGÚN EL SELECTOR
                if "Ingresos" in tipo_movimiento:
                    filtro_direccion = condicion_destino & (~condicion_origen)
                else:
                    filtro_direccion = condicion_origen & (~condicion_destino)
                
                mask_no_es_fantasma = ~df_raw_t['Destino'].isin(destinos_ignorados)
                mask_base_tecnica = (df_raw_t['Monto'] > 0) & (df_raw_t['Categoria'] != 'SERVICIO')

                df_tras_filtrados = df_raw_t[filtro_direccion & mask_no_es_fantasma & mask_base_tecnica]

                if df_tras_filtrados.empty:
                    st.warning(f"⚠️ No se detectaron {tipo_movimiento.split(' ')[0].lower()} válidos para {u_responsable} en este archivo.")
                else:
                    df_maestro_cta = obtener_dataframe("Categorias_Costos")
                    df_maestro_cta['Codigo'] = df_maestro_cta['Codigo'].astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
                    df_tras_final = pd.merge(df_tras_filtrados, df_maestro_cta[['Codigo', 'Cuenta_Contable']], on='Codigo', how='left')

                    c_nulas = df_tras_final['Cuenta_Contable'].isna() | df_tras_final['Cuenta_Contable'].astype(str).str.strip().str.upper().isin(["", "NAN", "NAT", "NONE"])

                    if c_nulas.any():
                        st.warning("⚠️ CÓDIGOS SIN CUENTA DETECTADOS: Asignales una acción para poder guardar.")
                        huerfanos_t = df_tras_final[c_nulas][['Codigo', 'Categoria', 'Origen', 'Destino']].drop_duplicates(subset=['Codigo'])
                        huerfanos_t['Accion'] = "Usar Categoría Nativa"
                        huerfanos_t['Cuenta_Manual'] = ""

                        ed_tras_h = st.data_editor(huerfanos_t, hide_index=True, use_container_width=True, key="ed_huerfanos_tras")

                        for _, r_ed in ed_tras_h.iterrows():
                            if r_ed['Accion'] == 'Omitir (No guardar)':
                                df_tras_final.loc[df_tras_final['Codigo'] == r_ed['Codigo'], 'Cuenta_Contable'] = 'OMITIDO_MANUAL'
                            elif r_ed['Accion'] == 'Escribir Cuenta Manual':
                                df_tras_final.loc[df_tras_final['Codigo'] == r_ed['Codigo'], 'Cuenta_Contable'] = r_ed['Cuenta_Manual']
                            else:
                                df_tras_final.loc[df_tras_final['Codigo'] == r_ed['Codigo'], 'Cuenta_Contable'] = r_ed['Categoria']

                    df_tras_final = df_tras_final[~df_tras_final['Cuenta_Contable'].astype(str).str.upper().isin(["OMITIDO_MANUAL", "NO APLICA", "0", "0.0"])]

                    if not df_tras_final.empty:
                        st.success(f"✅ {len(df_tras_final)} Movimientos validados.")
                        st.dataframe(df_tras_final, use_container_width=True)

                        st.markdown("#### 📥 Partidas de Traslado para Descargar:")
                        grupos_descarga = df_tras_final.groupby(['Origen', 'Destino'])
                        
                        for idx_g, ((o_v, d_v), df_g) in enumerate(grupos_descarga):
                            resumen_cta = df_g.groupby('Cuenta_Contable')['Monto'].sum()
                            glosa = f"TRASLADO DE {o_v} A {d_v}, {meses_texto[mes_reg]} {anio_reg}"
                            
                            datos_partida = []
                            for c_c, m_v in resumen_cta.items():
                                if m_v > 0.01: datos_partida.append([str(c_c).replace(".0",""), "", glosa, round(m_v, 2), 0.00, ""])
                            for c_c, m_v in resumen_cta.items():
                                if m_v > 0.01: datos_partida.append([str(c_c).replace(".0",""), "", glosa, 0.00, round(m_v, 2), ""])
                            
                            st.download_button(f"⬇️ {d_v} (Origen: {o_v})", generar_excel_bytes(datos_partida), f"Partida_{d_v}_desde_{o_v}.xlsx", key=f"dl_t_btn_{idx_g}")

                        if st.button("💾 Guardar registros en la base histórica", type="primary", use_container_width=True):
                            with st.spinner("Procesando guardado seguro..."):
                                st.cache_data.clear()
                                df_hist_v = obtener_dataframe("Historico_Traslados")
                                
                                def gen_key(r, m, a): return f"{m}|{a}|{r['Origen']}|{r['Destino']}|{r['Codigo']}"
                                df_tras_final['llave_auditoria'] = df_tras_final.apply(lambda r: gen_key(r, mes_reg, anio_reg), axis=1)

                                if not df_hist_v.empty:
                                    df_hist_v['llave'] = df_hist_v.apply(lambda r: f"{r['Mes']}|{r['Año']}|{r['Origen']}|{r['Destino']}|{r.get('Codigo', r.get('Código',''))}", axis=1)
                                    df_insertar = df_tras_final[~df_tras_final['llave_auditoria'].isin(df_hist_v['llave'].values)]
                                else:
                                    df_insertar = df_tras_final

                                if not df_insertar.empty:
                                    ws_historico = conectar_hoja("Historico_Traslados")
                                    f_actual = date.today().strftime('%d/%m/%Y')
                                    filas_batch = [[f_actual, mes_reg, anio_reg, r['Origen'], r['Destino'], str(r['Cuenta_Contable']), round(r['Monto'],2), r['Codigo'], "Traslado Nexus", r['Cantidad']] for _, r in df_insertar.iterrows()]
                                    ws_historico.append_rows(filas_batch)
                                    st.success(f"✅ Éxito: {len(df_insertar)} nuevos registros almacenados.")
                                else:
                                    st.warning("⚠️ No hay datos nuevos que guardar (Duplicados detectados).")
                                
                                st.cache_data.clear()

            except Exception as e_t:
                st.error(f"Error en el proceso de traslados: {e_t}")

    # =========================================================================
    # PESTAÑA 3: CONSULTA DE HISTORIAL OPERATIVO
    # =========================================================================
    with tab3:
        st.subheader("🔍 Consulta de Historial de Cierres")
        
        if st.button("🔄 Actualizar Base de Datos", key="update_hist_btn"):
            st.cache_data.clear()

        df_resumen = obtener_dataframe("Cierres_Costos")
        df_detalle = obtener_dataframe("Detalle_Cuentas")

        if not df_resumen.empty:
            # Creación del selector dinámico de periodos
            df_resumen['Periodo'] = df_resumen['Mes'].astype(str).str.replace('.0','') + "/" + df_resumen['Año'].astype(str).str.replace('.0','') + " - " + df_resumen['Unidad']
            periodo_seleccionado = st.selectbox("Seleccione el Cierre a Consultar:", df_resumen['Periodo'].unique().tolist(), index=len(df_resumen['Periodo'].unique())-1)

            if periodo_seleccionado:
                fila_resumen = df_resumen[df_resumen['Periodo'] == periodo_seleccionado].iloc[-1]
                
                mes_f, anio_f = str(fila_resumen['Mes']).strip(), str(fila_resumen['Año']).strip()
                
                # Conversión de valores a formato moneda para visualización
                v_ini = float(fila_resumen.iloc[4])
                v_com = float(fila_resumen.iloc[5])
                v_fin = float(fila_resumen.iloc[6])
                v_ant = float(fila_resumen.iloc[7])
                v_dif = float(fila_resumen.iloc[8])
                v_real = float(fila_resumen.iloc[9])

                # Layout de métricas principales del cierre
                m1, m2, m3, m4, m5 = st.columns(5)
                m1.metric("Inicial", f"${v_ini:,.2f}")
                m2.metric("Compras", f"${v_com:,.2f}")
                m3.metric("Final", f"${v_fin:,.2f}")
                m4.metric("Diferido", f"${v_dif:,.2f}")
                m5.metric("Costo Real", f"${v_real:,.2f}")

                # Sección para ver el detalle por cuenta contable
                df_det_h = df_detalle[(df_detalle['Mes'].astype(str).str.replace('.0','') == mes_f) & (df_detalle['Año'].astype(str).str.replace('.0','') == anio_f)].copy()
                
                if not df_det_h.empty:
                    df_det_h['Consumo'] = pd.to_numeric(df_det_h['Consumo'], errors='coerce').fillna(0.0)
                    
                    if st.checkbox("📂 Ver Detalle de Cuentas / Movimientos", key="chk_ver_det"):
                        st.dataframe(df_det_h, use_container_width=True)

                    st.divider()
                    st.markdown("#### 📥 Reconstruir Partidas Contables")
                    
                    formato_partida = st.radio("Tipo de Partida:", ["Cierre Estándar", "Cierre Consolidado"], key="radio_h")
                    
                    # Regeneración de diccionarios de consumo para archivos descargables
                    dict_consumo_h = df_det_h.groupby('Cuenta')['Consumo'].sum().to_dict()
                    total_op_h = sum(dict_consumo_h.values())

                    if formato_partida == "Cierre Estándar":
                        concepto_h = f"RECONOCIMIENTO DE COSTO DE VENTA DE CAFETERIA, {mes_f}/{anio_f}."
                        partida_v = [["410104", "", concepto_h, round(total_op_h, 2), 0.00, ""]]
                        
                        for cta, monto in dict_consumo_h.items():
                            cta_limpia = str(cta).replace(".0","").strip()
                            if cta_limpia != "" and cta_limpia.lower() not in ["nan", "nat", "no aplica"]:
                                if abs(monto) > 0.01:
                                    partida_v.append([cta_limpia, "", concepto_h, 0.00, round(monto, 2), ""])
                        
                        st.download_button(f"⬇️ Descargar P1 ({mes_f}/{anio_f})", generar_excel_bytes(partida_v), f"H_P1_{mes_f}_{anio_f}.xlsx", key="btn_h_p1")
        else:
            st.info("Aún no existen cierres guardados en la base histórica.")