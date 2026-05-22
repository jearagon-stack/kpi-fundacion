import streamlit as st
import pandas as pd
import io
import re

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
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=nombre_hoja)
    return output.getvalue()

# ==========================================
# MÓDULO PRINCIPAL DE AUDITORÍA
# ==========================================
def mostrar_modulo_auditoria():
    st.title("🔍 Auditoría de Cuentas y Parametrización")
    st.info("Cruce bidireccional entre Compras Operativas (Nexus) y Partidas Contables (CxP) para detectar fugas y errores de asignación.")

    st.markdown("##### 📥 Carga de Reportes")
    col1, col2 = st.columns(2)
    with col1:
        arch_ops = st.file_uploader("1. Reporte de Compras (Operaciones)", type=["xlsx", "xls"], key="audit_ops")
    with col2:
        arch_acc = st.file_uploader("2. Reporte de Movimientos (Contabilidad)", type=["xlsx", "xls"], key="audit_acc")

    if arch_ops and arch_acc:
        if st.button("🚀 Ejecutar Cruce de Auditoría", type="primary", use_container_width=True):
            
            with st.spinner("Escaneando documentos, cruzando cuentas y evaluando diferencias..."):
                try:
                    # Cargar DataFrames
                    df_ops = pd.read_excel(arch_ops, dtype=str)
                    df_acc = pd.read_excel(arch_acc, dtype=str)

                    df_ops.columns = df_ops.columns.str.strip()
                    df_acc.columns = df_acc.columns.str.strip()

                    # ==========================================
                    # 1. PROCESAMIENTO DE OPERACIONES (COMPRAS)
                    # ==========================================
                    
                    # Identificar columnas Operaciones
                    col_num_ops = next((c for c in df_ops.columns if 'NUMERO' in str(c).upper()), None)
                    col_tot_ops = next((c for c in df_ops.columns if 'TOTAL' in str(c).upper()), None)
                    col_cat_ops = next((c for c in df_ops.columns if 'CATEGORIA' in str(c).upper() or 'CATEGORÍA' in str(c).upper()), None)

                    if not all([col_num_ops, col_tot_ops, col_cat_ops]):
                        st.error("🚨 Error en archivo de Operaciones: No se encontraron las columnas Numero, Total o Categoria.")
                        st.stop()

                    df_ops['_Cat_Upper'] = df_ops[col_cat_ops].astype(str).str.upper().str.strip()

                    # ADUANA: Filtro Estricto de Categorías (Igual que Producción)
                    permitidas = ['MATERIA PRIMA', 'PRODUCTO TERMINADO', 'EMPAQUE', 'LIMPIEZA']
                    ignorar_silencio = ['SERVICIO']
                    
                    def es_ilegal(cat_str):
                        if pd.isna(cat_str) or cat_str in ['NAN', 'NULL', 'NONE', 'NA', '']: return True
                        return not (any(p in cat_str for p in permitidas) or any(i in cat_str for i in ignorar_silencio))

                    if df_ops['_Cat_Upper'].apply(es_ilegal).any():
                        invalidos = df_ops[df_ops['_Cat_Upper'].apply(es_ilegal)]
                        col_desc = 'Nombre' if 'Nombre' in df_ops.columns else col_num_ops
                        nombres = invalidos[col_desc].dropna().astype(str).unique()
                        lista = "\n- ".join(nombres[:10]) + ("\n- ..." if len(nombres) > 10 else "")
                        st.error(f"🚨 **PARO DE SEGURIDAD EN OPERACIONES** 🚨\n\nHay productos sin categoría o mal asignados.\n**Corrige antes de continuar:**\n\n- {lista}")
                        st.stop()

                    # Filtrar validos y sumarizar
                    df_ops = df_ops[df_ops['_Cat_Upper'].apply(lambda x: any(p in x for p in permitidas))]
                    df_ops['Monto_Ops'] = pd.to_numeric(df_ops[col_tot_ops], errors='coerce').fillna(0)
                    df_ops['Documento'] = df_ops[col_num_ops].astype(str).str.strip()
                    df_ops['Categoria'] = df_ops['_Cat_Upper']

                    # Agrupar operaciones por Documento y Categoría
                    df_ops_grouped = df_ops.groupby(['Documento', 'Categoria'])['Monto_Ops'].sum().reset_index()
                    documentos_conocidos = df_ops_grouped['Documento'].unique().tolist()
                    documentos_conocidos.sort(key=len, reverse=True) # Ordenar por longitud para que el regex agarre el más específico primero

                    # ==========================================
                    # 2. PROCESAMIENTO DE CONTABILIDAD
                    # ==========================================
                    
                    col_tipo_acc = next((c for c in df_acc.columns if 'IDTIPO' in str(c).upper()), None)
                    col_cta_acc = next((c for c in df_acc.columns if 'IDCUENTA' in str(c).upper()), None)
                    col_partida_acc = next((c for c in df_acc.columns if 'NUMERO' in str(c).upper()), None)
                    col_conc_acc = next((c for c in df_acc.columns if 'CONCEPTO' in str(c).upper()), None)
                    col_debe_acc = next((c for c in df_acc.columns if 'DEBE' in str(c).upper()), None)

                    if not all([col_tipo_acc, col_cta_acc, col_partida_acc, col_conc_acc, col_debe_acc]):
                        st.error("🚨 Error en archivo de Contabilidad: Faltan columnas clave (IdTipo, IdCuenta, Numero, Concepto, Debe).")
                        st.stop()

                    # Filtrar solo Cuentas por Pagar (CxP)
                    df_acc = df_acc[df_acc[col_tipo_acc].astype(str).str.upper().str.strip() == 'CXP'].copy()

                    # --- EXTRACTOR INTELIGENTE DE DOCUMENTOS ---
                    def extraer_doc(concepto):
                        concepto = str(concepto).upper()
                        # 1. Buscar si el concepto contiene un documento exacto de Operaciones
                        for doc in documentos_conocidos:
                            if doc.upper() in concepto:
                                return doc
                        # 2. Si no, buscar formatos estandar de Nexus
                        match = re.search(r'(CFE-[A-Z0-9\-]+|FSE-\d+|[A-Z0-9]{8}-[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{12})', concepto)
                        if match:
                            # Quitar prefijos comunes si los atrapa el regex
                            return match.group(1).replace('CFE-', '').replace('FSE-', 'FSE-')
                        return None

                    df_acc['Doc_Extraido'] = df_acc[col_conc_acc].apply(extraer_doc)

                    # --- HEREDAR DOCUMENTOS CIEGOS POR PARTIDA (Ej. PANCITO CON POLLO) ---
                    # Crear un diccionario: Número de Partida -> Documento Encontrado
                    mapa_partidas = df_acc.dropna(subset=['Doc_Extraido']).groupby(col_partida_acc)['Doc_Extraido'].unique().to_dict()

                    def rellenar_doc_ciego(row):
                        if pd.notna(row['Doc_Extraido']):
                            return row['Doc_Extraido']
                        docs_en_partida = mapa_partidas.get(row[col_partida_acc], [])
                        if len(docs_en_partida) == 1:
                            return docs_en_partida[0] # Hereda el documento de las otras líneas
                        elif len(docs_en_partida) > 1:
                            return docs_en_partida[0] # Si hay varios, se asume el primero por defecto
                        return "DOC_NO_IDENTIFICADO"

                    df_acc['Documento_Final'] = df_acc.apply(rellenar_doc_ciego, axis=1)

                    # --- FILTRAR CUENTAS Y MAPEAR CATEGORÍAS ---
                    mapa_cuentas = {
                        '110601': 'MATERIA PRIMA',
                        '110603': 'PRODUCTO TERMINADO',
                        '110608': 'LIMPIEZA',
                        '110609': 'EMPAQUE'
                    }
                    
                    df_acc['Cuenta_Limpia'] = df_acc[col_cta_acc].apply(limpiar_codigo_cuenta)
                    df_inv = df_acc[df_acc['Cuenta_Limpia'].isin(mapa_cuentas.keys())].copy()
                    df_inv['Categoria'] = df_inv['Cuenta_Limpia'].map(mapa_cuentas)
                    df_inv['Monto_Conta'] = pd.to_numeric(df_inv[col_debe_acc], errors='coerce').fillna(0)

                    # Agrupar contabilidad por Documento y Categoría Mapeada
                    df_acc_grouped = df_inv.groupby(['Documento_Final', 'Categoria'])['Monto_Conta'].sum().reset_index()
                    df_acc_grouped.rename(columns={'Documento_Final': 'Documento'}, inplace=True)

                    # ==========================================
                    # 3. CRUCE Y AUDITORÍA BIDIRECCIONAL
                    # ==========================================
                    
                    # Unir ambos dataframes asegurando que se vean todas las caras de la moneda
                    df_cruce = pd.merge(df_ops_grouped, df_acc_grouped, on=['Documento', 'Categoria'], how='outer')
                    
                    df_cruce['Monto_Ops'] = df_cruce['Monto_Ops'].fillna(0.0).round(2)
                    df_cruce['Monto_Conta'] = df_cruce['Monto_Conta'].fillna(0.0).round(2)
                    df_cruce['Diferencia ($)'] = (df_cruce['Monto_Ops'] - df_cruce['Monto_Conta']).round(2)

                    # Listas de control para identificar dónde está el error
                    docs_en_ops = df_ops_grouped['Documento'].unique()
                    docs_en_acc = df_acc_grouped['Documento'].unique()

                    def evaluar_auditoria(row):
                        doc = row['Documento']
                        m_ops = row['Monto_Ops']
                        m_acc = row['Monto_Conta']
                        
                        if m_ops > 0 and m_acc == 0:
                            # Está en Ops, pero no bajo esta categoría en Conta
                            if doc in docs_en_acc:
                                return "🔴 ERROR DE CUENTA CONTABLE (Categoría cruzada)"
                            else:
                                return "🔴 NO CONTABILIZADO (Falta la Partida o Inventario)"
                                
                        elif m_acc > 0 and m_ops == 0:
                            # Está en Conta bajo esta categoría, pero no en Ops
                            if doc in docs_en_ops:
                                return "🔴 ERROR DE CATEGORÍA OPERATIVA (Categoría cruzada)"
                            else:
                                return "🟡 NO EN OPERACIONES (Puede ser un ajuste manual)"
                                
                        else:
                            # Ambos existen en la misma categoría, evaluar dinero
                            diff = abs(row['Diferencia ($)'])
                            if diff > 0.05: # Tolerancia de centavos
                                return "🟠 DIFERENCIA DE MONTO"
                            else:
                                return "🟢 CUADRADO EXACTO"

                    df_cruce['Estado de Auditoría'] = df_cruce.apply(evaluar_auditoria, axis=1)

                    # Ordenar para que los errores salgan arriba
                    orden_estado = {
                        "🔴 ERROR DE CUENTA CONTABLE (Categoría cruzada)": 1,
                        "🔴 ERROR DE CATEGORÍA OPERATIVA (Categoría cruzada)": 2,
                        "🔴 NO CONTABILIZADO (Falta la Partida o Inventario)": 3,
                        "🟠 DIFERENCIA DE MONTO": 4,
                        "🟡 NO EN OPERACIONES (Puede ser un ajuste manual)": 5,
                        "🟢 CUADRADO EXACTO": 6
                    }
                    df_cruce['Prioridad'] = df_cruce['Estado de Auditoría'].map(orden_estado)
                    df_cruce = df_cruce.sort_values(['Prioridad', 'Documento']).drop(columns=['Prioridad'])

                    st.session_state['audit_cruce_df'] = df_cruce
                    st.session_state['audit_ejecutado'] = True

                except Exception as e:
                    st.error(f"Error procesando los archivos. Verifica los formatos. Detalle: {e}")

        # --- SECCIÓN VISUAL DE RESULTADOS ---
        if st.session_state.get('audit_ejecutado', False):
            df_final = st.session_state['audit_cruce_df']
            
            # Estadísticas rápidas
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
                df_final,
                column_config={
                    "Documento": st.column_config.TextColumn("Documento Fiscal / Interno"),
                    "Categoria": st.column_config.TextColumn("Categoría / Cuenta"),
                    "Monto_Ops": st.column_config.NumberColumn("Monto Operaciones ($)", format="$ %.2f"),
                    "Monto_Conta": st.column_config.NumberColumn("Monto Contabilidad ($)", format="$ %.2f"),
                    "Diferencia ($)": st.column_config.NumberColumn("Diferencia ($)", format="$ %.2f"),
                    "Estado de Auditoría": st.column_config.TextColumn("Estado de Auditoría")
                },
                use_container_width=True,
                hide_index=True
            )
            
            st.download_button(
                label="📥 Descargar Reporte de Auditoría (.xlsx)",
                data=generar_excel_auditoria(df_final),
                file_name=f"Auditoria_Cuentas_{date.today().strftime('%d_%m_%Y')}.xlsx",
                type="primary",
                use_container_width=True
            )