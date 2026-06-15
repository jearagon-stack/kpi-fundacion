import streamlit as st
import pandas as pd
from datetime import date
from utils import obtener_dataframe, conectar_hoja
import io
import re

def mostrar_modulo_costos():
    st.title("Contabilidad de Costos - Cafetería")

    tab1, tab2, tab3 = st.tabs(["📝 Generar Cierre", "🚚 Partidas de Traslados", "🔍 Consultar Histórico"])

    # ==========================================
    # FUNCIONES DE APOYO Y FILTROS ESTRICTOS
    # ==========================================
    
    BODEGAS_CAFETERIA = [
        "CAFETERIA JARDINES", 
        "CAFETERIA POLIDEPORTIVO", 
        "CAFETERIA ICAS", 
        "CAFETERIA EVENTOS", 
        "CAFETERIA ABASTECIMIENTO", 
        "CAFETERIA CENTRAL",
        "PRODUCCION CAFETERIA CENTRAL", 
        "GENERAL"
    ]

    def normalizar_texto(t):
        if pd.isna(t): return ""
        return str(t).strip().upper().replace('Í', 'I').replace('ÍA', 'IA')

    def es_cafeteria(b):
        b_norm = normalizar_texto(b)
        if "DESPENSA" in b_norm: return False
        claves_cafeteria = ["CAFETERIA", "JARDINES", "POLIDEPORTIVO", "ICAS", "EVENTOS", "ABASTECIMIENTO", "CENTRAL", "GENERAL"]
        return any(clave in b_norm for clave in claves_cafeteria)

    def es_despensa(b):
        return "DESPENSA" in normalizar_texto(b)

    def generar_excel_bytes(filas):
        df_p = pd.DataFrame(filas)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_p.to_excel(writer, index=False, header=False, sheet_name='Hoja1')
        return output.getvalue()

    def extraer_subunidad(nombre_archivo, unidad_base):
        nombre_upper = normalizar_texto(nombre_archivo)
        subunidades = ["POLIDEPORTIVO", "ICAS", "JARDINES", "EVENTOS", "CENTRAL", "ABASTECIMIENTO"]
        for sub in subunidades:
            if sub in nombre_upper: return f"{unidad_base} {sub}"
        return unidad_base

    def cargar_y_marcar(archivo):
        if archivo.name.lower().endswith('.csv'):
            archivo.seek(0)
            linea = archivo.readline().decode('utf-8', errors='ignore')
            separador = ';' if ';' in linea else ','
            archivo.seek(0)
            df = pd.read_csv(archivo, sep=separador, dtype=str)
        else:
            df = pd.read_excel(archivo, dtype=str)
            
        df.columns = df.columns.astype(str).str.strip().str.replace('\ufeff', '')
        cols_upper = df.columns.str.upper().str.replace(' ', '').str.replace('.', '')
        
        if not (any('COD' in c and 'PROD' in c for c in cols_upper) or any('IDPRODUCTO' in c for c in cols_upper)):
            for i in range(min(20, len(df))):
                row_str = df.iloc[i].astype(str).str.upper().str.replace(' ', '').str.replace('.', '')
                if any('COD' in val and 'PROD' in val for val in row_str) or any('IDPRODUCTO' in val for val in row_str):
                    df.columns = df.iloc[i].astype(str).str.strip()
                    df = df.iloc[i+1:].reset_index(drop=True)
                    break
                    
        col_rename = {}
        col_desc = None
        for col in df.columns:
            c_norm = str(col).upper().replace(' ', '').replace('.', '')
            if ('COD' in c_norm and 'PROD' in c_norm) or 'IDPRODUCTO' in c_norm:
                col_rename[col] = 'Codigo'
            if 'PROD' in c_norm or 'DESC' in c_norm or 'NOMBRE' in c_norm:
                col_desc = col
                
        df.rename(columns=col_rename, inplace=True)
        df['ORIGEN_ARCHIVO'] = archivo.name
        
        if 'Codigo' in df.columns:
            df = df[df['Codigo'].notna()]
            df = df[df['Codigo'].astype(str).str.strip() != '']
            df = df[~df['Codigo'].astype(str).str.upper().str.contains('TOTAL', na=False)]
            
        if col_desc:
            df = df[~df[col_desc].astype(str).str.upper().str.contains('TOTAL ', na=False)]
            df = df[~df[col_desc].astype(str).str.upper().str.startswith('TOTAL')]
            
        return df

    def limpiar_nativos_nexus(df):
        col_cat = next((c for c in df.columns if 'CATEGORIA' in str(c).upper()), None)
        if col_cat:
            df[col_cat] = df[col_cat].astype(str).str.upper().str.strip()
            df = df[df[col_cat] != 'SERVICIO']
        return df

    def consolidar(lista):
        if not lista: return pd.DataFrame()
        return pd.concat([limpiar_nativos_nexus(cargar_y_marcar(a)) for a in lista], ignore_index=True)

    def get_num(df, keys):
        for k in keys:
            for col in df.columns:
                c_norm = str(col).upper().replace(' ', '').replace('.', '')
                if all(p in c_norm for p in k): 
                    val_limpio = df[col].astype(str).str.replace(',', '', regex=False).str.replace('$', '', regex=False)
                    return pd.to_numeric(val_limpio, errors='coerce').fillna(0.0)
        return pd.Series(0.0, index=df.index)

    def get_col_exacta(df, opciones_exactas, opciones_parciales):
        for op in opciones_exactas:
            for c in df.columns:
                if c.strip().upper() == op: return c
        for op in opciones_parciales:
            for c in df.columns:
                if op in c.strip().upper(): return c
        return None

    def limpiar_cod(s): 
        return s.astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)

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
                
                suma_actual = sum(pesos) * 100
                if round(suma_actual, 2) != 100.0:
                    st.error(f"❌ Error de Distribución: Los porcentajes suman {suma_actual:.2f}%. Deben sumar exactamente 100.00%.")
                    st.stop()

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
                ventas_mes = pd.to_numeric(df_ventas[filtro_v]['Venta_Real'].astype(str).str.replace(',', ''), errors='coerce').sum()
                subsidio_mes = pd.to_numeric(df_ventas[filtro_v]['Subsidio_UCA'].astype(str).str.replace(',', ''), errors='coerce').sum()

            porcentaje_subsidio = (subsidio_mes / ventas_mes) if ventas_mes > 0 else 0.0
            st.info(f"📊 Ingresos Totales Periodo: Ventas ${ventas_mes:,.2f} | Subsidio ${subsidio_mes:,.2f}")

            costo_diferido_anterior = st.number_input("Costo Diferido de Arrastre (110602):", min_value=0.0, value=0.0)
            
            st.markdown("#### 📁 Archivos Base de Cierre (Cálculo de Consumo)")
            col_u1, col_u2, col_u3 = st.columns(3)
            with col_u1: arch_inv_maestro = st.file_uploader("1. Inventario (Ini/Fin)", type=["xlsx", "csv"], accept_multiple_files=True)
            with col_u2: arch_com = st.file_uploader("2. Compras", type=["xlsx", "csv"], accept_multiple_files=True)
            with col_u3: arch_tras_mes = st.file_uploader("3. Traslados del Mes", type=["xlsx", "csv"], accept_multiple_files=True)

            st.markdown("---")
            st.markdown("#### 🛡️ Auditoría Inteligente (Validación de Costos)")
            col_k1, col_k2 = st.columns(2)
            with col_k1: arch_kardex_aud = st.file_uploader("4. Kardex Valuado (Opcional)", type=["xlsx", "csv"], accept_multiple_files=True)
            with col_k2: arch_kardex_res = st.file_uploader("5. Kardex Resumen (Para Unificar Costos)", type=["xlsx", "csv"])

            if arch_inv_maestro and arch_com:
                if 'huerfanos_df' not in st.session_state:
                    forzar_calculo = st.checkbox("⚠️ Forzar cálculo ciego (Omitir revisión de cuentas)")
                    
                    if st.button("⚙️ Procesar Archivos y Calcular Consumo", type="primary", use_container_width=True):
                        with st.spinner("Agrupando productos y consolidando estructura matemática..."):
                            try:
                                mapa_costo_unificado = {}
                                if arch_kardex_res:
                                    try:
                                        df_resumen_g = leer_archivo_mixto(arch_kardex_res)
                                        df_resumen_g.columns = df_resumen_g.columns.astype(str).str.strip().str.upper()
                                        c_cod_res_g = next((c for c in df_resumen_g.columns if 'IDPRODUCTO' in c or 'COD' in c), df_resumen_g.columns[0])
                                        c_costo_res_g = next((c for c in df_resumen_g.columns if 'COSTOPROM' in c.replace(' ', '')), None)
                                        if not c_costo_res_g: c_costo_res_g = next((c for c in df_resumen_g.columns if 'COSTO' in c), None)
                                        
                                        if c_costo_res_g:
                                            costos_limpios = df_resumen_g[c_costo_res_g].astype(str).str.replace(',', '')
                                            mapa_costo_unificado = dict(zip(
                                                df_resumen_g[c_cod_res_g].str.strip().str.upper().str.replace(r'\.0$', '', regex=True),
                                                pd.to_numeric(costos_limpios, errors='coerce').fillna(0.0)
                                            ))
                                    except Exception as e:
                                        st.warning(f"No se pudo unificar costo del Kardex Resumen: {e}")

                                df_dic = obtener_dataframe("Categorias_Costos")
                                if not df_dic.empty and 'Codigo' in df_dic.columns:
                                    df_dic['Codigo'] = limpiar_cod(df_dic['Codigo'])
                                    df_dic = df_dic.drop_duplicates(subset=['Codigo'])
                                else:
                                    df_dic = pd.DataFrame(columns=['Codigo', 'Cuenta_Contable', 'Categoria'])

                                basura = ['G222', 'G231', '21455979']

                                # 1. PROCESAMIENTO DE INVENTARIOS
                                df_inv_raw = consolidar(arch_inv_maestro)
                                df_inv_grp = pd.DataFrame(columns=['Codigo', 'Valor_Ini', 'Valor_Fin'])
                                
                                if not df_inv_raw.empty:
                                    df_inv_raw['Codigo'] = limpiar_cod(df_inv_raw['Codigo'])
                                    df_inv_raw = df_inv_raw[~df_inv_raw['Codigo'].isin(basura)]
                                    
                                    val_ini_col = get_num(df_inv_raw, [['VALORINICIAL', 'SALDOINICIALVALOR'], ['TOTALINICIAL']])
                                    if val_ini_col.sum() > 0:
                                        df_inv_raw['Valor_Ini'] = val_ini_col.round(2)
                                    else:
                                        df_inv_raw['Cant_Ini'] = pd.to_numeric(get_num(df_inv_raw, [['EXISTENCIASINIC'], ['EXISTENCIA', 'INIC'], ['SALDO', 'INIC']]), errors='coerce').fillna(0.0)
                                        df_inv_raw['Costo_Ini'] = pd.to_numeric(get_num(df_inv_raw, [['COSTOUNITARIO'], ['COSTO', 'U'], ['PRECIO', 'U']]), errors='coerce').fillna(0.0).round(4)
                                        df_inv_raw['Valor_Ini'] = (df_inv_raw['Cant_Ini'] * df_inv_raw['Costo_Ini']).round(2)
                                    
                                    val_fin_col = get_num(df_inv_raw, [['VALORFINAL', 'SALDOFINALVALOR', 'TOTALFINAL']])
                                    if val_fin_col.sum() > 0:
                                        df_inv_raw['Valor_Fin'] = val_fin_col.round(2)
                                    else:
                                        df_inv_raw['Cant_Fin'] = pd.to_numeric(get_num(df_inv_raw, [['EXISTENCIASFIN'], ['EXISTENCIA', 'FIN'], ['SALDO', 'FIN']]), errors='coerce').fillna(0.0)
                                        if mapa_costo_unificado:
                                            df_inv_raw['Costo_Fin'] = df_inv_raw['Codigo'].map(mapa_costo_unificado).fillna(0.0).round(4)
                                        else:
                                            df_inv_raw['Costo_Fin'] = pd.to_numeric(get_num(df_inv_raw, [['COSTOUNITARIO'], ['COSTO', 'U'], ['PRECIO', 'U']]), errors='coerce').fillna(0.0).round(4)
                                        df_inv_raw['Valor_Fin'] = (df_inv_raw['Cant_Fin'] * df_inv_raw['Costo_Fin']).round(2)
                                    
                                    df_inv_grp = df_inv_raw.groupby('Codigo', as_index=False).agg({'Valor_Ini': 'sum', 'Valor_Fin': 'sum'})
                                    df_categoria_nat = df_inv_raw.groupby('Codigo', as_index=False).first()[['Codigo', 'Categoria']] if 'Categoria' in df_inv_raw.columns else pd.DataFrame(columns=['Codigo', 'Categoria'])
                                else:
                                    df_categoria_nat = pd.DataFrame(columns=['Codigo', 'Categoria'])

                                # 2. PROCESAMIENTO DE COMPRAS
                                df_com_raw = consolidar(arch_com)
                                df_com_grp = pd.DataFrame(columns=['Codigo', 'Valor_Com'])
                                
                                if not df_com_raw.empty:
                                    df_com_raw['Codigo'] = limpiar_cod(df_com_raw['Codigo'])
                                    df_com_raw = df_com_raw[~df_com_raw['Codigo'].isin(basura)]
                                    df_com_raw['Valor_Com'] = pd.to_numeric(get_num(df_com_raw, [['TOTAL'], ['MONTO'], ['VALOR']]), errors='coerce').fillna(0.0).round(2)
                                    df_com_grp = df_com_raw.groupby('Codigo', as_index=False).agg({'Valor_Com': 'sum'})

                                # 3. PROCESAMIENTO DE TRASLADOS (Lógica estricta de Origen/Destino)
                                df_tras_in_grp = pd.DataFrame(columns=['Codigo', 'Valor_Tras_In'])
                                df_tras_out_grp = pd.DataFrame(columns=['Codigo', 'Valor_Tras_Out'])
                                
                                if arch_tras_mes:
                                    df_tras_raw = consolidar(arch_tras_mes)
                                    if not df_tras_raw.empty:
                                        df_tras_raw['Codigo'] = limpiar_cod(df_tras_raw['Codigo'])
                                        df_tras_raw['Monto_T'] = pd.to_numeric(get_num(df_tras_raw, [['COSTOTOTAL', 'VALORTOTAL'], ['TOTAL'], ['MONTO'], ['VALOR']]), errors='coerce').fillna(0.0).round(2)
                                        
                                        c_orig = get_col_exacta(df_tras_raw, ['ORIGEN', 'SALIDA', 'BODEGAORIGEN', 'BODEGASALIDA'], ['ORIG', 'SALID'])
                                        c_dest = get_col_exacta(df_tras_raw, ['DESTINO', 'INGRESO', 'BODEGADESTINO', 'BODEGAINGRESO'], ['DEST', 'INGR'])
                                        
                                        if not c_orig or not c_dest:
                                            st.error("⚠️ En el archivo de Traslados no se detectaron columnas válidas de Origen y Destino. Se omitirán.")
                                        else:
                                            mask_orig_caf = df_tras_raw[c_orig].apply(es_cafeteria)
                                            mask_dest_caf = df_tras_raw[c_dest].apply(es_cafeteria)
                                            mask_orig_desp = df_tras_raw[c_orig].apply(es_despensa)
                                            mask_dest_desp = df_tras_raw[c_dest].apply(es_despensa)

                                            if unidad_cierre == "CAFETERIA":
                                                f_in = mask_dest_caf & (~mask_orig_caf)
                                                f_out = mask_orig_caf & (~mask_dest_caf)
                                            else:
                                                f_in = mask_dest_desp & (~mask_orig_desp)
                                                f_out = mask_orig_desp & (~mask_dest_desp)

                                            df_tras_in_grp = df_tras_raw[f_in].groupby('Codigo', as_index=False).agg({'Monto_T': 'sum'}).rename(columns={'Monto_T': 'Valor_Tras_In'})
                                            df_tras_out_grp = df_tras_raw[f_out].groupby('Codigo', as_index=False).agg({'Monto_T': 'sum'}).rename(columns={'Monto_T': 'Valor_Tras_Out'})

                                # 4. MATRIZ MAESTRA UNIFICADA
                                all_codes = set()
                                for d in [df_inv_grp, df_com_grp, df_tras_in_grp, df_tras_out_grp]:
                                    if not d.empty: all_codes.update(d['Codigo'].tolist())
                                
                                df_master = pd.DataFrame({'Codigo': list(all_codes)})
                                df_master = pd.merge(df_master, df_dic[['Codigo', 'Cuenta_Contable']], on='Codigo', how='left')
                                df_master = pd.merge(df_master, df_categoria_nat, on='Codigo', how='left')
                                df_master = pd.merge(df_master, df_inv_grp, on='Codigo', how='left').fillna(0)
                                df_master = pd.merge(df_master, df_com_grp, on='Codigo', how='left').fillna(0)
                                df_master = pd.merge(df_master, df_tras_in_grp, on='Codigo', how='left').fillna(0)
                                df_master = pd.merge(df_master, df_tras_out_grp, on='Codigo', how='left').fillna(0)

                                mask_huerf = df_master['Cuenta_Contable'].isna() | df_master['Cuenta_Contable'].astype(str).str.strip().str.upper().isin(["", "NAN", "NAT", "NONE"])
                                df_faltantes = df_master[mask_huerf][['Codigo', 'Categoria']].copy()
                                if 'Categoria' not in df_faltantes.columns or df_faltantes['Categoria'].isna().all():
                                    df_faltantes['Categoria'] = 'SIN CATEGORIA NATIVA'

                                if not df_faltantes.empty and not forzar_calculo:
                                    st.session_state['pre_proceso'] = {
                                        'df_master': df_master,
                                        'mapa_costo_unificado': mapa_costo_unificado
                                    }
                                    st.session_state['huerfanos_df'] = df_faltantes
                                    st.rerun()

                                # Si no hay huérfanos o se forzó el cálculo:
                                if 'Cuenta_Contable' in df_master.columns:
                                    df_master['Cuenta_Contable'] = df_master['Cuenta_Contable'].fillna("SIN CUENTA REGISTRADA")
                                    mask_null = df_master['Cuenta_Contable'].astype(str).str.strip().str.upper().isin(["", "NAN", "NAT", "NONE"])
                                    df_master.loc[mask_null, 'Cuenta_Contable'] = "SIN CUENTA REGISTRADA"
                                
                                df_master = df_master[~df_master['Cuenta_Contable'].isin(["OMITIDO_MANUAL", "NO APLICA", "0", "0.0"])]

                                df_master['Consumo'] = (df_master['Valor_Ini'] + df_master['Valor_Com'] + df_master['Valor_Tras_In'] - df_master['Valor_Tras_Out'] - df_master['Valor_Fin']).round(2)
                                
                                grp_ini_sum = round(df_master['Valor_Ini'].sum(), 2)
                                grp_comp_sum = round(df_master['Valor_Com'].sum(), 2)
                                grp_tras_sum = round(df_master['Valor_Tras_In'].sum() - df_master['Valor_Tras_Out'].sum(), 2)
                                grp_fin_sum = round(df_master['Valor_Fin'].sum(), 2)
                                
                                consumo_por_cuenta = df_master[df_master['Consumo'] != 0].groupby('Cuenta_Contable')['Consumo'].sum().to_dict()
                                costo_operativo = round(sum(consumo_por_cuenta.values()), 2)
                                
                                costo_dif_mes = round(float(costo_operativo) * float(porcentaje_subsidio), 2)
                                costo_real = round(float(costo_operativo) - float(costo_dif_mes) + float(costo_diferido_anterior), 2)

                                st.session_state['memoria_cierre'] = {
                                    'df_master': df_master,
                                    'grp_ini_sum': grp_ini_sum, 'grp_comp_sum': grp_comp_sum, 
                                    'grp_tras_sum': grp_tras_sum, 'grp_fin_sum': grp_fin_sum,
                                    'costo_dif_mes': costo_dif_mes, 'costo_real': costo_real, 'costo_operativo': costo_operativo,
                                    'costo_diferido_anterior': costo_diferido_anterior, 'consumo_por_cuenta': consumo_por_cuenta,
                                    'mes_cierre': mes_cierre, 'anio_cierre': anio_cierre, 'unidad_cierre': unidad_cierre,
                                    'es_consolidado': es_consolidado, 'pesos': pesos, 'nombres_meses': nombres_meses,
                                    'anomalias_kardex': pd.DataFrame(), 'anomalias_antiguas': pd.DataFrame()
                                }
                                
                                if 'huerfanos_df' in st.session_state: del st.session_state['huerfanos_df']
                                if 'pre_proceso' in st.session_state: del st.session_state['pre_proceso']
                                
                                st.rerun()

                            except Exception as e: 
                                st.error(f"Error procesando la matriz maestra: {e}")

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
                            "Accion": st.column_config.SelectboxColumn(
                                "¿Qué hacer?",
                                options=["Omitir (No sumar al costo)", "Usar Categoría Nativa", "Escribir Cuenta Manual"],
                                required=True
                            ),
                            "Cuenta_Manual": st.column_config.TextColumn("Cuenta (Si es Manual)")
                        },
                        hide_index=True, disabled=["Codigo", "Categoria"], use_container_width=True
                    )

                    col_b1, col_b2 = st.columns(2)
                    if col_b1.button("✅ Aplicar Decisiones y Generar Cierre", type="primary", use_container_width=True):
                        with st.spinner("Aplicando reglas y calculando matriz maestra..."):
                            df_master = st.session_state['pre_proceso']['df_master'].copy()

                            codigos_omitir = edited_df[edited_df['Accion'] == 'Omitir (No sumar al costo)']['Codigo'].tolist()
                            df_asignar = edited_df[edited_df['Accion'] == 'Escribir Cuenta Manual']
                            df_categoria = edited_df[edited_df['Accion'] == 'Usar Categoría Nativa']

                            mask_omitir = df_master['Codigo'].isin(codigos_omitir)
                            if mask_omitir.any(): df_master.loc[mask_omitir, 'Cuenta_Contable'] = 'OMITIDO_MANUAL'
                            
                            for _, row in df_asignar.iterrows():
                                c_manual = str(row['Cuenta_Manual']).strip()
                                if c_manual == "": c_manual = "SIN CUENTA REGISTRADA"
                                df_master.loc[df_master['Codigo'] == row['Codigo'], 'Cuenta_Contable'] = c_manual
                            
                            for _, row in df_categoria.iterrows():
                                cat_val = str(row['Categoria']).strip()
                                df_master.loc[df_master['Codigo'] == row['Codigo'], 'Cuenta_Contable'] = cat_val

                            df_master['Cuenta_Contable'] = df_master['Cuenta_Contable'].fillna("SIN CUENTA REGISTRADA")
                            mask_null = df_master['Cuenta_Contable'].astype(str).str.strip().str.upper().isin(["", "NAN", "NAT", "NONE"])
                            df_master.loc[mask_null, 'Cuenta_Contable'] = "SIN CUENTA REGISTRADA"
                            
                            df_master = df_master[~df_master['Cuenta_Contable'].isin(["OMITIDO_MANUAL", "NO APLICA", "0", "0.0"])]

                            df_master['Consumo'] = (df_master['Valor_Ini'] + df_master['Valor_Com'] + df_master['Valor_Tras_In'] - df_master['Valor_Tras_Out'] - df_master['Valor_Fin']).round(2)
                            
                            grp_ini_sum = round(df_master['Valor_Ini'].sum(), 2)
                            grp_comp_sum = round(df_master['Valor_Com'].sum(), 2)
                            grp_tras_sum = round(df_master['Valor_Tras_In'].sum() - df_master['Valor_Tras_Out'].sum(), 2)
                            grp_fin_sum = round(df_master['Valor_Fin'].sum(), 2)
                            
                            consumo_por_cuenta = df_master[df_master['Consumo'] != 0].groupby('Cuenta_Contable')['Consumo'].sum().to_dict()
                            costo_operativo = round(sum(consumo_por_cuenta.values()), 2)
                            
                            costo_dif_mes = round(float(costo_operativo) * float(porcentaje_subsidio), 2)
                            costo_real = round(float(costo_operativo) - float(costo_dif_mes) + float(costo_diferido_anterior), 2)

                            st.session_state['memoria_cierre'] = {
                                'df_master': df_master,
                                'grp_ini_sum': grp_ini_sum, 'grp_comp_sum': grp_comp_sum, 
                                'grp_tras_sum': grp_tras_sum, 'grp_fin_sum': grp_fin_sum,
                                'costo_dif_mes': costo_dif_mes, 'costo_real': costo_real, 'costo_operativo': costo_operativo,
                                'costo_diferido_anterior': costo_diferido_anterior, 'consumo_por_cuenta': consumo_por_cuenta,
                                'mes_cierre': mes_cierre, 'anio_cierre': anio_cierre, 'unidad_cierre': unidad_cierre,
                                'es_consolidado': es_consolidado, 'pesos': pesos, 'nombres_meses': nombres_meses,
                                'anomalias_kardex': pd.DataFrame(), 'anomalias_antiguas': pd.DataFrame()
                            }
                            
                            if 'huerfanos_df' in st.session_state: del st.session_state['huerfanos_df']
                            if 'pre_proceso' in st.session_state: del st.session_state['pre_proceso']
                            
                            st.rerun()

        else:
            mem = st.session_state['memoria_cierre']
            st.success(f"📦 **DATOS EN MEMORIA:** Cierre de {mem['unidad_cierre']} - Periodo: {meses_texto[mem['mes_cierre']]} {mem['anio_cierre']}")
            
            r1, r2, r3, r4, r5, r6 = st.columns(6)
            r1.metric("Inicial (+)", f"${mem['grp_ini_sum']:,.2f}")
            r2.metric("Compras (+)", f"${mem['grp_comp_sum']:,.2f}")
            r3.metric("Traslados Netos (+/-)", f"${mem.get('grp_tras_sum', 0.0):,.2f}")
            r4.metric("Final (-)", f"${mem['grp_fin_sum']:,.2f}")
            r5.metric("Diferido (-)", f"${mem['costo_dif_mes']:,.2f}")
            r6.metric("Real (=)", f"${mem['costo_real']:,.2f}")

            st.divider()
            if st.checkbox("📂 Ver Detalle de Consumo por Cuentas"):
                df_det_view = pd.DataFrame(list(mem['consumo_por_cuenta'].items()), columns=['Cuenta Contable', 'Consumo (Impacto)'])
                st.dataframe(df_det_view.style.format({'Consumo (Impacto)':'${:.2f}'}), use_container_width=True)

            if not st.session_state.get('auditoria_aprobada', False):
                st.warning("⚠️ Debes revisar los resultados operativos antes de dar el cierre por bueno.")
                col_b1, col_b2 = st.columns(2)
                if col_b1.button("🗑️ Descartar Memoria y Subir Archivos Nuevos"):
                    del st.session_state['memoria_cierre']
                    st.rerun()
                if col_b2.button("✅ Aprobar Auditoría y Proceder"):
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
                    diferido_para_liquidar = mem['costo_diferido_anterior'] 
                    
                    for i in range(len(mem['pesos'])):
                        consumo_mes = mem['costo_operativo'] * mem['pesos'][i]
                        nuevo_diferido_mes = mem['costo_dif_mes'] * mem['pesos'][i]
                        
                        mostrar_descargas_logic(
                            consumo_mes, 
                            nuevo_diferido_mes, 
                            diferido_para_liquidar, 
                            mem['nombres_meses'][i], 
                            mem['consumo_por_cuenta'], 
                            mem['costo_operativo'], 
                            f"cons_{i}"
                        )
                        diferido_para_liquidar = nuevo_diferido_mes

                if st.button("💾 Cerrar Periodo y Guardar Base", type="primary", use_container_width=True):
                    with st.spinner("Guardando en la nube..."):
                        ws_res = conectar_hoja("Cierres_Costos"); ws_det = conectar_hoja("Detalle_Cuentas")
                        fecha_hoy = date.today().strftime('%d/%m/%Y')
                        if ws_res and ws_det:
                            ws_res.append_row([fecha_hoy, mem['mes_cierre'], mem['anio_cierre'], mem['unidad_cierre'], round(mem['grp_ini_sum'],2), round(mem['grp_comp_sum'],2), round(mem['grp_fin_sum'],2), round(mem['costo_diferido_anterior'],2), round(mem['costo_dif_mes'],2), round(mem['costo_real'],2)])
                            
                            df_guardar = mem['df_master'][mem['df_master']['Consumo'] != 0].copy()
                            filas_g = []
                            for _, r in df_guardar.iterrows():
                                filas_g.append([fecha_hoy, mem['mes_cierre'], mem['anio_cierre'], mem['unidad_cierre'], str(r['Cuenta_Contable']), round(r['Valor_Ini'],2), round(r['Valor_Com'],2), round(r['Valor_Fin'],2), round(r['Consumo'],2), r['Codigo'], "Archivo_Consolidado"])
                            if filas_g: ws_det.append_rows(filas_g)
                            
                        del st.session_state['memoria_cierre']
                        st.session_state['auditoria_aprobada'] = False
                        st.cache_data.clear()
                        st.rerun()

    # =========================================================================
    # PESTAÑA 2: REGISTRO DE TRASLADOS (FILTRO LÓGICO ESTRICTO)
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

        st.divider()
        st.info("💡 Filtro Activo: El sistema omitirá cualquier movimiento que se dé internamente entre las bodegas base establecidas.")

        archivo_nexus = st.file_uploader("Reporte Nexus (A:Tipo, I:Cod, K:Cant, N:Monto, AA:Cat, AC:Salida, AD:Ingreso)", type=["xlsx", "csv"], key="atf_reg")

        if archivo_nexus:
            try:
                df_raw_t = leer_archivo_mixto(archivo_nexus)
                
                c_cod = get_col_exacta(df_raw_t, ['CODIGO', 'IDPRODUCTO'], ['COD'])
                c_cant = get_col_exacta(df_raw_t, ['CANTIDAD', 'UNIDADES'], ['CANT'])
                c_monto = get_col_exacta(df_raw_t, ['MONTO', 'TOTAL', 'VALOR', 'COSTOTOTAL'], ['MONT', 'TOT'])
                c_cat = get_col_exacta(df_raw_t, ['CATEGORIA'], ['CAT'])
                c_orig = get_col_exacta(df_raw_t, ['ORIGEN', 'SALIDA', 'BODEGAORIGEN'], ['ORIG', 'SALID'])
                c_dest = get_col_exacta(df_raw_t, ['DESTINO', 'INGRESO', 'BODEGADESTINO'], ['DEST', 'INGR'])

                if not all([c_cod, c_cant, c_monto, c_cat, c_orig, c_dest]):
                    st.error("🚨 Faltan columnas en el reporte de traslados. Asegúrate de tener: Código, Cantidad, Monto, Categoría, Origen y Destino.")
                    st.stop()
                
                df_raw_t = df_raw_t.rename(columns={c_cod:'Codigo', c_cant:'Cantidad', c_monto:'Monto', c_cat:'Categoria', c_orig:'Origen', c_dest:'Destino'})
                
                df_raw_t['Codigo'] = df_raw_t['Codigo'].astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
                df_raw_t['Monto'] = pd.to_numeric(df_raw_t['Monto'].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0)
                df_raw_t['Cantidad'] = pd.to_numeric(df_raw_t['Cantidad'].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0)

                df_raw_t = df_raw_t[df_raw_t['Codigo'].notna()]
                df_raw_t = df_raw_t[~df_raw_t['Codigo'].astype(str).str.upper().str.contains('TOTAL', na=False)]

                def es_interno(b, unidad):
                    if unidad == "CAFETERIA": return es_cafeteria(b)
                    else: return es_despensa(b)

                mask_orig_interna = df_raw_t['Origen'].apply(lambda x: es_interno(x, u_responsable))
                mask_dest_interna = df_raw_t['Destino'].apply(lambda x: es_interno(x, u_responsable))

                filtro_direccion = mask_orig_interna != mask_dest_interna
                mask_base_tecnica = (df_raw_t['Monto'] > 0) & (df_raw_t['Categoria'].astype(str).str.upper() != 'SERVICIO')

                df_tras_filtrados = df_raw_t[filtro_direccion & mask_base_tecnica]

                if df_tras_filtrados.empty:
                    st.warning("⚠️ No se encontraron movimientos aplicables para registro externo en este archivo.")
                else:
                    df_maestro_cta = obtener_dataframe("Categorias_Costos")
                    if not df_maestro_cta.empty and 'Codigo' in df_maestro_cta.columns:
                        df_maestro_cta['Codigo'] = df_maestro_cta['Codigo'].astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
                        df_maestro_cta = df_maestro_cta.drop_duplicates(subset=['Codigo'])
                    
                    df_tras_final = pd.merge(df_tras_filtrados, df_maestro_cta[['Codigo', 'Cuenta_Contable']], on='Codigo', how='left')

                    c_nulas = df_tras_final['Cuenta_Contable'].isna() | df_tras_final['Cuenta_Contable'].astype(str).str.strip().str.upper().isin(["", "NAN", "NAT", "NONE"])

                    if c_nulas.any():
                        st.warning("⚠️ CÓDIGOS SIN CUENTA DETECTADOS: Asignales una acción para poder continuar.")
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
                        st.success(f"✅ {len(df_tras_final)} Movimientos validados listos para descarga local.")
                        st.dataframe(df_tras_final, use_container_width=True)

                        st.markdown("#### 📥 Partidas para Descargar:")
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

            except Exception as e_t:
                st.error(f"Error procesando el archivo: {e_t}")

    # =========================================================================
    # PESTAÑA 3: CONSULTA DE HISTORIAL OPERATIVO
    # =========================================================================
    with tab3:
        st.subheader("🔍 Consulta de Historial de Cierres")
        
        if st.button("🔄 Actualizar Base de Datos", key="btn_update_historico"):
            st.cache_data.clear()

        df_resumen_c = obtener_dataframe("Cierres_Costos")
        df_detalle_c = obtener_dataframe("Detalle_Cuentas")

        if not df_resumen_c.empty:
            df_resumen_c['Periodo'] = df_resumen_c['Mes'].astype(str).str.replace('.0','') + "/" + df_resumen_c['Año'].astype(str).str.replace('.0','') + " - " + df_resumen_c['Unidad']
            p_sel = st.selectbox("Seleccione el Cierre:", df_resumen_c['Periodo'].unique().tolist(), index=len(df_resumen_c['Periodo'].unique())-1)

            if p_sel:
                f_res = df_resumen_c[df_resumen_c['Periodo'] == p_sel].iloc[-1]
                m_cons, a_cons = str(f_res['Mes']).strip(), str(f_res['Año']).strip()
                
                v_i = float(str(f_res.iloc[4]).replace(',', ''))
                v_c = float(str(f_res.iloc[5]).replace(',', ''))
                v_f = float(str(f_res.iloc[6]).replace(',', ''))
                v_a = float(str(f_res.iloc[7]).replace(',', ''))
                v_d = float(str(f_res.iloc[8]).replace(',', ''))
                v_r = float(str(f_res.iloc[9]).replace(',', ''))

                c_m1, c_m2, c_m3, c_m4, c_m5 = st.columns(5)
                c_m1.metric("Inicial", f"${v_i:,.2f}")
                c_m2.metric("Compras", f"${v_c:,.2f}")
                c_m3.metric("Final", f"${v_f:,.2f}")
                c_m4.metric("Diferido", f"${v_d:,.2f}")
                c_m5.metric("Costo Real", f"${v_r:,.2f}")

                df_det_huerf = df_detalle_c[(df_detalle_c['Mes'].astype(str).str.replace('.0','') == m_cons) & (df_detalle_c['Año'].astype(str).str.replace('.0','') == a_cons)].copy()
                
                if not df_det_huerf.empty:
                    df_det_huerf['Consumo'] = pd.to_numeric(df_det_huerf['Consumo'].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0)
                    
                    if st.checkbox("📂 Mostrar Detalle de Cuentas", key="chk_ver_detalle_h"):
                        st.dataframe(df_det_huerf, use_container_width=True)

                    st.divider()
                    st.markdown("#### 📥 Regenerar Partidas Contables")
                    
                    f_partida = st.radio("Formato de Partida:", ["Cierre Estándar", "Cierre Consolidado"], key="r_tipo_h")
                    dict_c_h = df_det_huerf.groupby('Cuenta')['Consumo'].sum().to_dict()
                    total_h = sum(dict_c_h.values())

                    if f_partida == "Cierre Estándar":
                        con_h = f"RECONOCIMIENTO DE COSTO DE VENTA DE {f_res['Unidad']}, {m_cons}/{a_cons}."
                        p_v_h = [["410104", "", con_h, round(total_h, 2), 0.00, ""]]
                        
                        for ct_h, mt_h in dict_c_h.items():
                            cl_h = str(ct_h).replace(".0","").strip()
                            if cl_h != "" and cl_h.lower() not in ["nan", "nat", "no aplica"]:
                                if abs(mt_h) > 0.01:
                                    p_v_h.append([cl_h, "", con_h, 0.00, round(mt_h, 2), ""])
                        
                        st.download_button(f"⬇️ Descargar Partida ({m_cons}/{a_cons})", generar_excel_bytes(p_v_h), f"H_P1_{m_cons}_{a_cons}.xlsx", key="btn_h_descarga")

                    elif f_partida == "Cierre Consolidado":
                        num_meses_h = st.number_input("¿En cuántos meses se dividió el cierre?", min_value=2, max_value=12, value=3, key="num_m_h")
                        
                        pesos_h = []
                        nombres_meses_h = []
                        
                        st.write("Distribución de porcentajes para recrear partidas:")
                        cols_h = st.columns(num_meses_h)
                        for i in range(num_meses_h):
                            with cols_h[i]:
                                nom_m = st.text_input(f"Nombre Mes {i+1}:", value=f"Mes {i+1}", key=f"nm_h_{i}")
                                peso_m = st.number_input(f"% Mes {i+1}:", min_value=0.0, max_value=100.0, value=100.0/num_meses_h, key=f"pm_h_{i}")
                                nombres_meses_h.append(nom_m)
                                pesos_h.append(peso_m / 100.0)
                        
                        if sum(pesos_h) > 0:
                            diferido_para_liquidar = v_a 
                            
                            for i in range(num_meses_h):
                                peso_actual = pesos_h[i]
                                mes_label = nombres_meses_h[i]
                                
                                c_op_mes = total_h * peso_actual
                                nuevo_diferido_mes = v_d * peso_actual
                                
                                st.markdown(f"##### 📥 Partidas: **{mes_label}**")
                                conc_v = f"RECONOCIMIENTO DE COSTO DE VENTA DE {f_res['Unidad']}, {mes_label} {a_cons}."
                                f_v = [["410104", "", conc_v, round(c_op_mes, 2), 0.00, ""]]
                                
                                for ct_h, mt_h in dict_c_h.items():
                                    cl_h = str(ct_h).replace(".0","").strip()
                                    if cl_h != "" and cl_h.lower() not in ["nan", "nat", "no aplica"]:
                                        m_perc = mt_h * peso_actual
                                        if abs(m_perc) > 0.01:
                                            f_v.append([cl_h, "", conc_v, 0.00, round(m_perc, 2), ""])
                                
                                conc_p = f"RECONOCIMIENTO DE COSTO DE LO VENDIDO EN PROCESO {f_res['Unidad']} {mes_label} {a_cons}"
                                f_p = [["410104", "", conc_p, round(diferido_para_liquidar, 2), 0.00, ""], ["110602", "", conc_p, 0.00, round(diferido_para_liquidar, 2), ""]]
                                
                                conc_d = f"DIFERIMIENTO DE COSTO EN PROCESO {f_res['Unidad']} {mes_label} {a_cons}"
                                f_d = [["110602", "", conc_d, round(nuevo_diferido_mes, 2), 0.00, ""], ["410104", "", conc_d, 0.00, round(nuevo_diferido_mes, 2), ""]]

                                p1, p2, p3 = st.columns(3)
                                with p1: st.download_button(f"⬇️ P1 {mes_label}", generar_excel_bytes(f_v), f"H_1_Costo_{mes_label}.xlsx", key=f"dl_p1_h_{i}", use_container_width=True)
                                with p2: st.download_button(f"⬇️ P2 {mes_label}", generar_excel_bytes(f_p), f"H_2_Ant_{mes_label}.xlsx", key=f"dl_p2_h_{i}", use_container_width=True)
                                with p3: st.download_button(f"⬇️ P3 {mes_label}", generar_excel_bytes(f_d), f"H_3_Dif_{mes_label}.xlsx", key=f"dl_p3_h_{i}", use_container_width=True)

                                diferido_para_liquidar = nuevo_diferido_mes
        else:
            st.info("No hay cierres previos registrados en el historial.")

if __name__ == "__main__":
    mostrar_modulo_costos()