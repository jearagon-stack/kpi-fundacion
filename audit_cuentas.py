import streamlit as st
import pandas as pd
import io
import re
from datetime import date

# ==========================================
# FUNCIONES AUXILIARES
# ==========================================
def limpiar_codigo_cuenta(c):
    if pd.isna(c) or str(c).strip() == "": return "SIN_CUENTA"
    val = str(c).strip()
    if val.endswith('.0'): return val[:-2]
    return val

def generar_excel_auditoria(df, nombre_hoja="Auditoria_Cuentas"):
    output = io.BytesIO()
    df_descarga = df.drop(columns=['Categoria_Original'], errors='ignore')
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_descarga.to_excel(writer, index=False, sheet_name=nombre_hoja[:31])
    return output.getvalue()

def leer_archivo_mixto(file):
    """Permite leer tanto archivos Excel como CSV sin que el sistema falle."""
    if file.name.lower().endswith('.csv'):
        return pd.read_csv(file, dtype=str)
    return pd.read_excel(file, dtype=str)

# ==========================================
# MÓDULO PRINCIPAL DE AUDITORÍA
# ==========================================
def mostrar_modulo_auditoria():
    st.title("🔍 Auditoría de Cuentas y Parametrización")
    
    # --- CREACIÓN DE PESTAÑAS PARA SEPARAR MÓDULOS ---
    tab_compras, tab_inventario = st.tabs([
        "🛒 1. Auditoría de Compras (CXP)", 
        "📦 2. Auditoría de Descargas (Inventarios)"
    ])

    # ==========================================
    # PESTAÑA 1: TU CÓDIGO ORIGINAL INTACTO
    # ==========================================
    with tab_compras:
        st.info("Cruce bidireccional entre Compras Operativas (Nexus) y Partidas Contables (CxP) para detectar fugas y errores de asignación.")

        st.markdown("##### 📥 Carga de Reportes")
        col1, col2 = st.columns(2)
        with col1:
            arch_ops = st.file_uploader("1. Reporte de Compras (Operaciones)", type=["xlsx", "xls"], key="audit_ops")
        with col2:
            arch_acc = st.file_uploader("2. Reporte de Movimientos (Contabilidad)", type=["xlsx", "xls"], key="audit_acc")

        if arch_ops and arch_acc:
            if st.button("🚀 Ejecutar Cruce de Auditoría", type="primary", use_container_width=True):
                
                with st.spinner("Triangulando documentos y desglosando cuentas..."):
                    try:
                        df_ops = pd.read_excel(arch_ops, dtype=str)
                        df_acc = pd.read_excel(arch_acc, dtype=str)

                        df_ops.columns = df_ops.columns.str.strip()
                        df_acc.columns = df_acc.columns.str.strip()

                        # ==========================================
                        # 1. PROCESAMIENTO DE OPERACIONES (COMPRAS)
                        # ==========================================
                        col_num_ops = next((c for c in df_ops.columns if 'NUMERO' in str(c).upper()), None)
                        col_tot_ops = next((c for c in df_ops.columns if 'TOTAL' in str(c).upper()), None)
                        col_cat_ops = next((c for c in df_ops.columns if 'CATEGORIA' in str(c).upper()), None)
                        col_desc_ops = next((c for c in df_ops.columns if 'DESCRIPCION' in str(c).upper() or 'NOMBRE' in str(c).upper() and 'MAYOR' not in str(c).upper()), None)

                        try: 
                            col_tipo_ops = df_ops.columns[10]
                        except IndexError: 
                            col_tipo_ops = None

                        if not all([col_num_ops, col_tot_ops, col_cat_ops, col_desc_ops]):
                            st.error("🚨 Error en Operaciones: Faltan columnas clave (Numero, Total, Categoria o Descripcion).")
                            st.stop()

                        df_ops['_Cat_Upper'] = df_ops[col_cat_ops].astype(str).str.upper().str.strip()

                        permitidas = ['MATERIA PRIMA', 'PRODUCTO TERMINADO', 'EMPAQUE', 'LIMPIEZA', 'REPUESTO']
                        ignorar_silencio = ['SERVICIO']
                        
                        def es_ilegal(cat_str):
                            if pd.isna(cat_str) or cat_str in ['NAN', 'NULL', 'NONE', 'NA', '']: return True
                            return not (any(p in cat_str for p in permitidas) or any(i in cat_str for i in ignorar_silencio))

                        if df_ops['_Cat_Upper'].apply(es_ilegal).any():
                            invalidos = df_ops[df_ops['_Cat_Upper'].apply(es_ilegal)]
                            col_desc_err = 'Nombre' if 'Nombre' in df_ops.columns else col_num_ops
                            nombres = invalidos[col_desc_err].dropna().astype(str).unique()
                            lista = "\n- ".join(nombres[:10]) + ("\n- ..." if len(nombres) > 10 else "")
                            st.error(f"🚨 **PARO DE SEGURIDAD EN OPERACIONES** 🚨\nHay productos sin categoría o mal asignados:\n{lista}")
                            st.stop()

                        df_ops = df_ops[df_ops['_Cat_Upper'].apply(lambda x: any(p in x for p in permitidas))].copy()
                        
                        # Limpieza segura del Documento usando Regex
                        df_ops['Documento'] = df_ops[col_num_ops].astype(str).str.strip().str.upper()
                        df_ops['Documento'] = df_ops['Documento'].str.replace(r'^CFE-', '', regex=True).str.replace(r'^FSE-', '', regex=True)
                        
                        df_ops['Categoria'] = df_ops['_Cat_Upper']
                        df_ops['Desc_Limpia'] = df_ops[col_desc_ops].astype(str).str.upper().str.strip()
                        df_ops['Tipo_Doc'] = df_ops[col_tipo_ops].astype(str).str.upper().str.strip() if col_tipo_ops else "NO IDENTIFICADO"

                        # Convertir Monto_Ops a número y aplicar la REGLA DE LOS SIGNOS para Notas de Crédito
                        df_ops['Monto_Ops_Abs'] = pd.to_numeric(df_ops[col_tot_ops], errors='coerce').fillna(0)
                        df_ops['Monto_Ops'] = df_ops.apply(lambda r: -abs(r['Monto_Ops_Abs']) if 'NOTA DE CRÉDITO' in r['Tipo_Doc'] or 'NOTA DE CREDITO' in r['Tipo_Doc'] else r['Monto_Ops_Abs'], axis=1)

                        df_ops_grouped = df_ops.groupby(['Documento', 'Categoria']).agg({
                            'Monto_Ops': 'sum',
                            'Tipo_Doc': 'first'
                        }).reset_index()
                        
                        documentos_conocidos = df_ops['Documento'].unique().tolist()
                        documentos_conocidos.sort(key=len, reverse=True) 
                        ops_docs_desc = df_ops.groupby('Documento')['Desc_Limpia'].apply(list).to_dict()

                        # ==========================================
                        # 2. PROCESAMIENTO DE CONTABILIDAD
                        # ==========================================
                        col_tipo_acc = next((c for c in df_acc.columns if 'IDTIPO' in str(c).upper()), None)
                        col_cta_acc = next((c for c in df_acc.columns if 'IDCUENTA' in str(c).upper() and 'MAYOR' not in str(c).upper()), None)
                        col_partida_acc = next((c for c in df_acc.columns if 'NUMERO' in str(c).upper()), None)
                        col_conc_acc = next((c for c in df_acc.columns if 'CONCEPTO' in str(c).upper()), None)
                        col_debe_acc = next((c for c in df_acc.columns if 'DEBE' in str(c).upper()), None)
                        col_haber_acc = next((c for c in df_acc.columns if 'HABER' in str(c).upper()), None)
                        
                        col_nom_cta_acc = next((c for c in df_acc.columns if str(c).strip().upper() == 'NOMBRE'), None)
                        if not col_nom_cta_acc: 
                            col_nom_cta_acc = next((c for c in df_acc.columns if 'NOMBRE' in str(c).upper() and 'MAYOR' not in str(c).upper()), None)

                        if not all([col_tipo_acc, col_cta_acc, col_partida_acc, col_conc_acc, col_debe_acc, col_nom_cta_acc]):
                            st.error("🚨 Error en Contabilidad: Faltan columnas clave.")
                            st.stop()

                        df_acc = df_acc[df_acc[col_tipo_acc].astype(str).str.upper().str.strip() == 'CXP'].copy()

                        def extraer_doc(concepto):
                            c = str(concepto).upper()
                            match_uuid = re.search(r'([A-Z0-9]{8}-[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{12})', c)
                            if match_uuid: return match_uuid.group(1)
                            for doc in documentos_conocidos:
                                if doc in c: return doc
                            match_fse = re.search(r'(?:CFE|FSE)-([A-Z0-9\-]+)', c)
                            if match_fse: return match_fse.group(1)
                            return None

                        df_acc['Doc_Extraido'] = df_acc[col_conc_acc].apply(extraer_doc)
                        mapa_partidas = df_acc.dropna(subset=['Doc_Extraido']).groupby(col_partida_acc)['Doc_Extraido'].unique().to_dict()

                        def rellenar_doc_ciego(row):
                            if pd.notna(row['Doc_Extraido']): return row['Doc_Extraido']
                            docs_en_partida = mapa_partidas.get(row[col_partida_acc], [])
                            if len(docs_en_partida) == 1: 
                                return docs_en_partida[0] 
                            elif len(docs_en_partida) > 1:
                                concepto_conta = str(row[col_conc_acc]).upper()
                                for doc in docs_en_partida:
                                    for desc in ops_docs_desc.get(doc, []):
                                        if desc in concepto_conta: return doc
                            return "NO_IDENTIFICADO"

                        df_acc['Documento_Final'] = df_acc.apply(rellenar_doc_ciego, axis=1)

                        col_f_ops = next((c for c in df_ops.columns if 'FECHA' in str(c).upper()), None)
                        col_p_ops = next((c for c in df_ops.columns if 'PROV' in str(c).upper()), None)
                        col_f_acc = next((c for c in df_acc.columns if 'FECHA' in str(c).upper()), None)
                        col_p_acc = next((c for c in df_acc.columns if 'PROV' in str(c).upper() or 'REFERENCIA' in str(c).upper()), None)

                        df_acc['Debe_Num'] = pd.to_numeric(df_acc[col_debe_acc], errors='coerce').fillna(0)
                        df_acc['Haber_Num'] = pd.to_numeric(df_acc[col_haber_acc], errors='coerce').fillna(0) if col_haber_acc else 0
                        df_acc['Monto_Conta_Neto'] = df_acc['Debe_Num'] - df_acc['Haber_Num']

                        def triangulacion_ciegos(row):
                            if row['Documento_Final'] != "NO_IDENTIFICADO": return row['Documento_Final']
                            monto_acc = row['Monto_Conta_Neto']
                            if pd.isna(monto_acc) or monto_acc == 0: return "NO_IDENTIFICADO"
                            
                            concepto_acc = str(row[col_conc_acc]).upper()
                            fecha_acc = str(row[col_f_acc]).strip() if col_f_acc else ""
                            prov_acc = str(row[col_p_acc]).upper().strip() if col_p_acc else ""

                            ops_match = df_ops[abs(df_ops['Monto_Ops'] - monto_acc) < 0.05]
                            if ops_match.empty: return "NO_IDENTIFICADO"

                            posibles_docs = []
                            for _, op_row in ops_match.iterrows():
                                desc_op = str(op_row['Desc_Limpia'])
                                fecha_op = str(op_row[col_f_ops]).strip() if col_f_ops else ""
                                prov_op = str(op_row[col_p_ops]).upper().strip() if col_p_ops else ""

                                match_desc = (desc_op in concepto_acc) or (concepto_acc in desc_op)
                                if (fecha_acc and fecha_op and fecha_acc != fecha_op): continue
                                if (prov_acc and prov_op and prov_op not in prov_acc and prov_acc not in prov_op): continue
                                if match_desc: posibles_docs.append(op_row['Documento'])
                            
                            if len(set(posibles_docs)) == 1: return posibles_docs[0]
                            return "NO_IDENTIFICADO"
                        
                        df_acc['Documento_Final'] = df_acc.apply(triangulacion_ciegos, axis=1)

                        mapa_cuentas_temp = ['110601', '110603', '110608', '110609', '110610']
                        
                        mascara_ciegos = (df_acc['Documento_Final'] == 'NO_IDENTIFICADO') & (abs(df_acc['Monto_Conta_Neto']) > 0) & (df_acc[col_cta_acc].astype(str).str.strip().isin(mapa_cuentas_temp))
                        ciegos_relevantes = df_acc[mascara_ciegos]
                        
                        if not ciegos_relevantes.empty:
                            ejemplos = ciegos_relevantes[[col_partida_acc, col_conc_acc]].drop_duplicates().head(10)
                            lista = "".join([f"\n- **Partida:** {r[col_partida_acc]} | **Concepto:** {r[col_conc_acc]}" for _, r in ejemplos.iterrows()])
                            st.error(f"🚨 **ALERTA ROJA: DOCUMENTOS NO IDENTIFICADOS EN CONTABILIDAD** 🚨\n{lista}")
                            st.stop()

                        mapa_cuentas = {
                            '110601': 'MATERIA PRIMA',
                            '110603': 'PRODUCTO TERMINADO',
                            '110610': 'REPUESTO',
                            '110608': 'LIMPIEZA',
                            '110609': 'EMPAQUE'
                            
                        }
                        
                        df_acc['Cuenta_Limpia'] = df_acc[col_cta_acc].astype(str).str.strip().apply(lambda c: c[:-2] if c.endswith('.0') else c)
                        df_inv = df_acc[df_acc['Cuenta_Limpia'].isin(mapa_cuentas.keys())].copy()
                        df_inv['Categoria'] = df_inv['Cuenta_Limpia'].map(mapa_cuentas)
                        df_inv['Cuenta_Nom'] = df_inv[col_nom_cta_acc].astype(str).str.strip()

                        # AGRUPAMOS POR CUENTA PARA QUE SE DESGLOSE EN MÚLTIPLES FILAS
                        df_acc_grouped = df_inv.groupby(['Documento_Final', 'Categoria', 'Cuenta_Nom']).agg({
                            'Monto_Conta_Neto': 'sum',
                            col_partida_acc: lambda x: ', '.join(x.dropna().astype(str).unique())
                        }).reset_index()
                        df_acc_grouped.rename(columns={'Documento_Final': 'Documento', 'Monto_Conta_Neto': 'Monto_Conta', col_partida_acc: 'Partida_Conta'}, inplace=True)

                        # ==========================================
                        # 3. CRUCE BIDIRECCIONAL FINAL
                        # ==========================================
                        df_cruce = pd.merge(df_ops_grouped, df_acc_grouped, on=['Documento', 'Categoria'], how='outer')
                        df_cruce['Monto_Ops'] = df_cruce['Monto_Ops'].fillna(0.0)
                        df_cruce['Monto_Conta'] = df_cruce['Monto_Conta'].fillna(0.0)
                        
                        # EVITAR DUPLICAR MONTO DE OPERACIONES CUANDO CONTABILIDAD SE DIVIDE EN VARIAS CUENTAS
                        df_cruce['is_dup'] = df_cruce.duplicated(subset=['Documento', 'Categoria'])
                        df_cruce.loc[df_cruce['is_dup'], 'Monto_Ops'] = 0.0
                        
                        df_cruce['Diferencia ($)'] = (df_cruce['Monto_Ops'] - df_cruce['Monto_Conta']).round(2)
                        df_cruce['Cuenta_Nom'] = df_cruce['Cuenta_Nom'].fillna('SIN REGISTRO EN CONTA')
                        df_cruce['Partida_Conta'] = df_cruce['Partida_Conta'].fillna('SIN PARTIDA')
                        df_cruce['Tipo_Doc'] = df_cruce['Tipo_Doc'].fillna('NO EN OPS')

                        # Configuración de las columnas visibles y ocultas
                        df_cruce['Categoria_Original'] = df_cruce['Categoria']
                        df_cruce['Categoria Operativa'] = df_cruce.apply(lambda r: r['Categoria'] if r['Monto_Ops'] != 0 else 'SIN REGISTRO EN OPS', axis=1)

                        # CALCULAMOS EL GLOBAL DEL DOCUMENTO PARA EVALUAR SU ESTADO CORRECTAMENTE
                        df_cruce['doc_ops_total'] = df_cruce.groupby(['Documento', 'Categoria_Original'])['Monto_Ops'].transform('sum')
                        df_cruce['doc_acc_total'] = df_cruce.groupby(['Documento', 'Categoria_Original'])['Monto_Conta'].transform('sum')

                        docs_en_ops = set(df_ops_grouped['Documento'].unique())
                        docs_en_acc = set(df_acc_grouped['Documento'].unique())

                        def evaluar_auditoria(row):
                            doc = str(row['Documento'])
                            m_ops = row['doc_ops_total']
                            m_acc = row['doc_acc_total']
                            dif_global = abs(m_ops - m_acc)
                            
                            if m_ops != 0 and m_acc == 0:
                                if doc in docs_en_acc: return "🔴 ERROR DE CUENTA CONTABLE (Categoría cruzada)"
                                else: return "🔴 NO CONTABILIZADO (Falta la Partida o Inventario)"
                            elif m_acc != 0 and m_ops == 0:
                                if doc in docs_en_ops: return "🔴 ERROR DE CATEGORÍA OPERATIVA (Categoría cruzada)"
                                else: return "🟡 NO EN OPERACIONES (Puede ser un ajuste manual)"
                            else:
                                if dif_global > 0.05: return "🟠 DIFERENCIA DE MONTO"
                                else: return "🟢 CUADRADO EXACTO"

                        df_cruce['Estado de Auditoría'] = df_cruce.apply(evaluar_auditoria, axis=1)

                        orden_estado = {
                            "🔴 ERROR DE CUENTA CONTABLE (Categoría cruzada)": 1,
                            "🔴 ERROR DE CATEGORÍA OPERATIVA (Categoría cruzada)": 2,
                            "🔴 NO CONTABILIZADO (Falta la Partida o Inventario)": 3,
                            "🟠 DIFERENCIA DE MONTO": 4,
                            "🟡 NO EN OPERACIONES (Puede ser un ajuste manual)": 5,
                            "🟢 CUADRADO EXACTO": 6
                        }
                        df_cruce['Prioridad'] = df_cruce['Estado de Auditoría'].map(orden_estado).fillna(99)
                        df_cruce = df_cruce.sort_values(['Prioridad', 'Documento']).drop(columns=['Prioridad', 'is_dup', 'doc_ops_total', 'doc_acc_total', 'Categoria'])

                        columnas_ordenadas = [
                            'Partida_Conta', 'Tipo_Doc', 'Documento', 'Categoria Operativa', 'Cuenta_Nom', 
                            'Monto_Ops', 'Monto_Conta', 'Diferencia ($)', 'Estado de Auditoría', 'Categoria_Original'
                        ]
                        df_cruce = df_cruce[columnas_ordenadas]
                        df_cruce = df_cruce.rename(columns={
                            'Cuenta_Nom': 'Cuenta Contable (Conta)',
                            'Partida_Conta': 'Partida Contable (Conta)',
                            'Tipo_Doc': 'Tipo Doc (Ops)'
                        })

                        st.session_state['audit_cruce_df'] = df_cruce
                        st.session_state['audit_ejecutado'] = True

                    except Exception as e:
                        st.error(f"Error procesando los archivos. Verifica los formatos. Detalle: {e}")

        # --- SECCIÓN VISUAL DE RESULTADOS ---
        if st.session_state.get('audit_ejecutado', False):
            df_final = st.session_state['audit_cruce_df']
            
            errores_cuenta = len(df_final[df_final['Estado de Auditoría'].str.contains('ERROR DE CUENTA')])
            no_conta = len(df_final[df_final['Estado de Auditoría'].str.contains('NO CONTABILIZADO')])
            dif_monto = len(df_final[df_final['Estado de Auditoría'].str.contains('DIFERENCIA')])
            cuadrados = len(df_final[df_final['Estado de Auditoría'].str.contains('CUADRADO')])
            
            st.success("✅ Auditoría completada con éxito.")
            st.markdown("---")
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Errores de Cuenta", errores_cuenta)
            c2.metric("Falta Contabilizar", no_conta)
            c3.metric("Diferencias de Monto", dif_monto)
            c4.metric("Documentos Cuadrados", cuadrados)
            
            st.subheader("📋 Matriz de Validación de Cuentas")
            
            st.dataframe(
                df_final.drop(columns=['Categoria_Original'], errors='ignore'),
                column_config={
                    "Partida Contable (Conta)": st.column_config.TextColumn("Partida"),
                    "Tipo Doc (Ops)": st.column_config.TextColumn("Tipo"),
                    "Documento": st.column_config.TextColumn("Documento Fiscal"),
                    "Categoria Operativa": st.column_config.TextColumn("Categoría Operativa"),
                    "Cuenta Contable (Conta)": st.column_config.TextColumn("Cuenta Detectada"),
                    "Monto_Ops": st.column_config.NumberColumn("Monto Ops ($)", format="$ %.2f"),
                    "Monto_Conta": st.column_config.NumberColumn("Monto Conta ($)", format="$ %.2f"),
                    "Diferencia ($)": st.column_config.NumberColumn("Diferencia ($)", format="$ %.2f"),
                    "Estado de Auditoría": st.column_config.TextColumn("Estado de Auditoría")
                },
                use_container_width=True,
                hide_index=True
            )
            
            st.markdown("---")
            st.subheader("📊 Resumen Macro por Categoría")
            st.info("Este panel suma los dólares exactos de las columnas de arriba para asegurar que los grandes totales coincidan.")
            
            df_resumen = df_final.groupby('Categoria_Original')[['Monto_Ops', 'Monto_Conta', 'Diferencia ($)']].sum().reset_index()
            df_resumen.rename(columns={'Categoria_Original': 'Categoría General'}, inplace=True)
            
            st.dataframe(
                df_resumen,
                column_config={
                    "Categoría General": st.column_config.TextColumn("Categoría General"),
                    "Monto_Ops": st.column_config.NumberColumn("Total Operaciones ($)", format="$ %.2f"),
                    "Monto_Conta": st.column_config.NumberColumn("Total Contabilidad ($)", format="$ %.2f"),
                    "Diferencia ($)": st.column_config.NumberColumn("Descuadre Global ($)", format="$ %.2f")
                },
                use_container_width=True,
                hide_index=True
            )

            st.download_button(
                label="📥 Descargar Reporte de Auditoría Detallado (.xlsx)",
                data=generar_excel_auditoria(df_final),
                file_name=f"Auditoria_Cuentas_{date.today().strftime('%d_%m_%Y')}.xlsx",
                type="primary",
                use_container_width=True
            )

    # ==========================================
    # PESTAÑA 2: NUEVO MÓDULO DE INVENTARIOS
    # ==========================================
    with tab_inventario:
        st.info("Cruce entre Descargas Contables (Libro Mayor) y Salidas Operativas (Kardex Valuado) para detectar diferencias en el Costo de lo Vendido.")

        st.markdown("##### 📥 Carga de Reportes para Inventario")
        col_i1, col_i2, col_i3 = st.columns(3)
        with col_i1:
            arch_acc_inv = st.file_uploader("1. Libro Mayor (Conta)", type=["xlsx", "xls", "csv"], key="inv_acc")
        with col_i2:
            arch_kardex = st.file_uploader("2. Kardex Valuado (Ops)", type=["xlsx", "xls", "csv"], key="inv_kar")
        with col_i3:
            arch_kardex_res = st.file_uploader("3. Kardex Resumen (Mapeo)", type=["xlsx", "xls", "csv"], key="inv_map")

        if arch_acc_inv and arch_kardex and arch_kardex_res:
            if st.button("🚀 Ejecutar Cruce de Inventarios", type="primary", use_container_width=True):
                with st.spinner("Procesando y cruzando descargas de inventario..."):
                    try:
                        df_acc_inv = leer_archivo_mixto(arch_acc_inv)
                        df_kar = leer_archivo_mixto(arch_kardex)
                        df_map = leer_archivo_mixto(arch_kardex_res)

                        df_acc_inv.columns = df_acc_inv.columns.str.strip().str.upper()
                        df_kar.columns = df_kar.columns.str.strip().str.upper()
                        df_map.columns = df_map.columns.str.strip().str.upper()

                        # --- FASE 1: PROCESAMIENTO DE CONTABILIDAD ---
                        col_tipo_acc_inv = next((c for c in df_acc_inv.columns if 'IDTIPO' in c or 'TIPO' in c), None)
                        col_cta_acc_inv = next((c for c in df_acc_inv.columns if 'IDCUENTA' in c or 'CUENTA' in c), None)
                        col_haber_acc_inv = next((c for c in df_acc_inv.columns if 'HABER' in c), None)

                        if not all([col_tipo_acc_inv, col_cta_acc_inv, col_haber_acc_inv]):
                            st.error("🚨 Error en Libro Mayor: Faltan columnas (IdTipo, IdCuenta o Haber).")
                            st.stop()

                        # Filtrar excluyendo CXP y CONT para dejar solo las descargas operativas (DESP, TRRZ, CID, etc.)
                        filtro_tipo = ~df_acc_inv[col_tipo_acc_inv].astype(str).str.upper().str.strip().isin(['CXP', 'CONT'])
                        df_acc_inv_filt = df_acc_inv[filtro_tipo].copy()

                        df_acc_inv_filt['Haber_Num'] = pd.to_numeric(df_acc_inv_filt[col_haber_acc_inv], errors='coerce').fillna(0)
                        
                        # Mapeo idéntico de cuentas a categorías maestras
                        mapa_cuentas_inv = {
                            '110601': 'MATERIA PRIMA',
                            '110603': 'PRODUCTO TERMINADO',
                            '110610': 'REPUESTO',
                            '110608': 'LIMPIEZA',
                            '110609': 'EMPAQUE'
                        }
                        df_acc_inv_filt['Cuenta_Limpia'] = df_acc_inv_filt[col_cta_acc_inv].astype(str).str.strip().apply(lambda c: c[:-2] if c.endswith('.0') else c)
                        df_acc_inv_filt['Categoria'] = df_acc_inv_filt['Cuenta_Limpia'].map(mapa_cuentas_inv).fillna('OTRA CUENTA (' + df_acc_inv_filt['Cuenta_Limpia'] + ')')
                        
                        # Agrupar saldos contables
                        df_acc_grouped = df_acc_inv_filt.groupby('Categoria')['Haber_Num'].sum().reset_index()
                        df_acc_grouped.rename(columns={'Haber_Num': 'Descargado en Contabilidad ($)'}, inplace=True)

                        # --- FASE 2: PROCESAMIENTO DEL KARDEX Y MAPEO ---
                        col_pref = next((c for c in df_kar.columns if 'PREFIJO' in c), None)
                        col_prod_kar = next((c for c in df_kar.columns if 'IDPRODUCTO' in c or 'PRODUCTO' in c), None)
                        col_salidas = next((c for c in df_kar.columns if 'SALIDASVALOR' in c or 'SALIDAS' in c), None)

                        col_prod_map = next((c for c in df_map.columns if 'IDPRODUCTO' in c or 'PRODUCTO' in c or 'CODIGO' in c), None)
                        col_cat_map = next((c for c in df_map.columns if 'CATEGORIA' in c or 'CUENTA' in c), None)

                        if not all([col_pref, col_prod_kar, col_salidas]):
                            st.error("🚨 Error en Kardex Valuado: Faltan columnas (Prefijo, IdProducto o SalidasValor).")
                            st.stop()
                        if not all([col_prod_map, col_cat_map]):
                            st.error("🚨 Error en Kardex Resumen: Faltan columnas (IdProducto/Codigo o Categoria).")
                            st.stop()

                        # Crear diccionario para cruzar el Producto con su Categoría usando el Kardex Resumen
                        mapa_prod_cat = df_map.set_index(col_prod_map)[col_cat_map].to_dict()

                        # Filtrar solo transacciones de venta (CCF, FCF y sus variantes electrónicas CFE, FSE)
                        prefijos_validos = ['CCF', 'FCF', 'CFE', 'FSE']
                        df_kar_filt = df_kar[df_kar[col_pref].astype(str).str.upper().str.strip().isin(prefijos_validos)].copy()

                        df_kar_filt['Salidas_Num'] = pd.to_numeric(df_kar_filt[col_salidas], errors='coerce').fillna(0)
                        
                        # Asignar la categoría desde el Kardex Resumen
                        df_kar_filt['Categoria_Mapeada'] = df_kar_filt[col_prod_kar].astype(str).str.strip().map(mapa_prod_cat).str.upper().str.strip()
                        df_kar_filt['Categoria'] = df_kar_filt['Categoria_Mapeada'].fillna('SIN CATEGORIA EN MAPEO')

                        # Agrupar saldos de operaciones
                        df_kar_grouped = df_kar_filt.groupby('Categoria')['Salidas_Num'].sum().reset_index()
                        df_kar_grouped.rename(columns={'Salidas_Num': 'Salidas en Kardex ($)'}, inplace=True)

                        # --- FASE 3: CRUCE Y MATRIZ DE DIFERENCIAS ---
                        df_cruce_inv = pd.merge(df_acc_grouped, df_kar_grouped, on='Categoria', how='outer').fillna(0)
                        
                        df_cruce_inv['Diferencia Neta ($)'] = df_cruce_inv['Descargado en Contabilidad ($)'] - df_cruce_inv['Salidas en Kardex ($)']
                        
                        def accion_recomendada(dif):
                            if dif > 0.05: return "⬇️ Conta tiene más salidas (Ajustar sobrante)"
                            elif dif < -0.05: return "⬆️ Conta tiene menos salidas (Ajustar faltante)"
                            return "✅ Cuadrado"
                            
                        df_cruce_inv['Acción Recomendada'] = df_cruce_inv['Diferencia Neta ($)'].apply(accion_recomendada)

                        st.success("✅ Cruce de Inventarios ejecutado con éxito.")
                        
                        st.dataframe(
                            df_cruce_inv,
                            column_config={
                                "Categoria": st.column_config.TextColumn("Categoría de Inventario"),
                                "Descargado en Contabilidad ($)": st.column_config.NumberColumn("Descargado en Contabilidad (Haber)", format="$ %.2f"),
                                "Salidas en Kardex ($)": st.column_config.NumberColumn("Salidas en Kardex (Costo Ventas)", format="$ %.2f"),
                                "Diferencia Neta ($)": st.column_config.NumberColumn("Diferencia Neta ($)", format="$ %.2f"),
                                "Acción Recomendada": st.column_config.TextColumn("Recomendación de Ajuste")
                            },
                            use_container_width=True,
                            hide_index=True
                        )

                        st.download_button(
                            label="📥 Descargar Reporte de Diferencias de Inventario (.xlsx)",
                            data=generar_excel_auditoria(df_cruce_inv, "Auditoria_Descargas"),
                            file_name=f"Auditoria_Inventarios_{date.today().strftime('%d_%m_%Y')}.xlsx",
                            type="primary",
                            use_container_width=True
                        )

                    except Exception as e:
                        st.error(f"Error técnico durante el cruce de inventarios: {e}")