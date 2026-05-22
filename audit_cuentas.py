import streamlit as st
import pandas as pd
import io
import re
from datetime import date # <- ¡Solución al error de pantalla roja!

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
                    
                    col_num_ops = next((c for c in df_ops.columns if 'NUMERO' in str(c).upper()), None)
                    col_tot_ops = next((c for c in df_ops.columns if 'TOTAL' in str(c).upper()), None)
                    col_cat_ops = next((c for c in df_ops.columns if 'CATEGORIA' in str(c).upper() or 'CATEGORÍA' in str(c).upper()), None)
                    col_desc_ops = next((c for c in df_ops.columns if 'DESCRIPCION' in str(c).upper() or 'NOMBRE' in str(c).upper() and 'MAYOR' not in str(c).upper()), None)

                    if not all([col_num_ops, col_tot_ops, col_cat_ops, col_desc_ops]):
                        st.error("🚨 Error en Operaciones: Faltan columnas clave (Numero, Total, Categoria o Descripcion).")
                        st.stop()

                    df_ops['_Cat_Upper'] = df_ops[col_cat_ops].astype(str).str.upper().str.strip()

                    # ADUANA: Filtro Estricto de Categorías
                    permitidas = ['MATERIA PRIMA', 'PRODUCTO TERMINADO', 'EMPAQUE', 'LIMPIEZA']
                    ignorar_silencio = ['SERVICIO']
                    
                    def es_ilegal(cat_str):
                        if pd.isna(cat_str) or cat_str in ['NAN', 'NULL', 'NONE', 'NA', '']: return True
                        return not (any(p in cat_str for p in permitidas) or any(i in cat_str for i in ignorar_silencio))

                    if df_ops['_Cat_Upper'].apply(es_ilegal).any():
                        invalidos = df_ops[df_ops['_Cat_Upper'].apply(es_ilegal)]
                        col_desc_err = 'Nombre' if 'Nombre' in df_ops.columns else col_num_ops
                        nombres = invalidos[col_desc_err].dropna().astype(str).unique()
                        lista = "\n- ".join(nombres[:10]) + ("\n- ..." if len(nombres) > 10 else "")
                        st.error(f"🚨 **PARO DE SEGURIDAD EN OPERACIONES** 🚨\n\nHay productos sin categoría o mal asignados.\n**Corrige antes de continuar:**\n\n- {lista}")
                        st.stop()

                    # Limpiar y preparar Operaciones
                    df_ops = df_ops[df_ops['_Cat_Upper'].apply(lambda x: any(p in x for p in permitidas))].copy()
                    df_ops['Monto_Ops'] = pd.to_numeric(df_ops[col_tot_ops], errors='coerce').fillna(0)
                    
                    # Normalizar los Documentos (Quitando posibles prefijos para cruce limpio)
                    df_ops['Documento'] = df_ops[col_num_ops].astype(str).str.strip().str.upper()
                    df_ops['Documento'] = df_ops['Documento'].str.replace('CFE-', '').str.replace('FSE-', '')
                    df_ops['Categoria'] = df_ops['_Cat_Upper']
                    df_ops['Desc_Limpia'] = df_ops[col_desc_ops].astype(str).str.upper().str.strip()

                    # Agrupar operaciones para el cruce final
                    df_ops_grouped = df_ops.groupby(['Documento', 'Categoria'])['Monto_Ops'].sum().reset_index()
                    
                    # Diccionario para "Adivinar" facturas múltiples por descripcion
                    documentos_conocidos = df_ops['Documento'].unique().tolist()
                    documentos_conocidos.sort(key=len, reverse=True) # Los más largos primero (UUIDs antes que "00008")
                    ops_docs_desc = df_ops.groupby('Documento')['Desc_Limpia'].apply(list).to_dict()

                    # ==========================================
                    # 2. PROCESAMIENTO DE CONTABILIDAD
                    # ==========================================
                    
                    col_tipo_acc = next((c for c in df_acc.columns if 'IDTIPO' in str(c).upper()), None)
                    col_cta_acc = next((c for c in df_acc.columns if 'IDCUENTA' in str(c).upper() and 'MAYOR' not in str(c).upper()), None)
                    col_partida_acc = next((c for c in df_acc.columns if 'NUMERO' in str(c).upper()), None)
                    col_conc_acc = next((c for c in df_acc.columns if 'CONCEPTO' in str(c).upper()), None)
                    col_debe_acc = next((c for c in df_acc.columns if 'DEBE' in str(c).upper()), None)

                    if not all([col_tipo_acc, col_cta_acc, col_partida_acc, col_conc_acc, col_debe_acc]):
                        st.error("🚨 Error en Contabilidad: Faltan columnas clave (IdTipo, IdCuenta, Numero, Concepto, Debe).")
                        st.stop()

                    # Filtrar Cuentas por Pagar (CxP)
                    df_acc = df_acc[df_acc[col_tipo_acc].astype(str).str.upper().str.strip() == 'CXP'].copy()

                    # --- EXTRACTOR INTELIGENTE Y ROBUSTO ---
                    def extraer_doc(concepto):
                        c = str(concepto).upper()
                        # 1. Buscar UUIDs formales explícitos
                        match_uuid = re.search(r'([A-Z0-9]{8}-[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{12})', c)
                        if match_uuid: return match_uuid.group(1)
                        # 2. Buscar si el string del documento (ej. "00008") está dentro del concepto ("FSE-00008")
                        for doc in documentos_conocidos:
                            if doc in c: return doc
                        # 3. Rescate por formato general CFE o FSE
                        match_fse = re.search(r'(?:CFE|FSE)-([A-Z0-9\-]+)', c)
                        if match_fse: return match_fse.group(1)
                        return None

                    df_acc['Doc_Extraido'] = df_acc[col_conc_acc].apply(extraer_doc)

                    # --- RESOLVEDOR DE PARTIDAS MULTIPLES ---
                    mapa_partidas = df_acc.dropna(subset=['Doc_Extraido']).groupby(col_partida_acc)['Doc_Extraido'].unique().to_dict()

                    def rellenar_doc_ciego(row):
                        if pd.notna(row['Doc_Extraido']):
                            return row['Doc_Extraido']
                        
                        docs_en_partida = mapa_partidas.get(row[col_partida_acc], [])
                        if len(docs_en_partida) == 1:
                            return docs_en_partida[0] # Fácil, solo hay 1 factura en la partida
                        
                        elif len(docs_en_partida) > 1:
                            # ¡Múltiples facturas! Buscar a cuál le pertenece leyendo la descripción de la zanahoria/pan
                            concepto_conta = str(row[col_conc_acc]).upper()
                            for doc in docs_en_partida:
                                descripciones_ops = ops_docs_desc.get(doc, [])
                                for desc in descripciones_ops:
                                    if desc in concepto_conta:
                                        return doc
                            # Si no se pudo enlazar, lo marcamos para que se note la mezcla
                            return f"MULTIPLE_NO_IDENTIFICADO_{row[col_partida_acc]}"
                            
                        return "DOC_NO_IDENTIFICADO"

                    df_acc['Documento_Final'] = df_acc.apply(rellenar_doc_ciego, axis=1)

                    # --- MAPEO DE CUENTAS CONTABLES ---
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

                    df_acc_grouped = df_inv.groupby(['Documento_Final', 'Categoria'])['Monto_Conta'].sum().reset_index()
                    df_acc_grouped.rename(columns={'Documento_Final': 'Documento'}, inplace=True)

                    # ==========================================
                    # 3. CRUCE BIDIRECCIONAL FINAL
                    # ==========================================
                    
                    df_cruce = pd.merge(df_ops_grouped, df_acc_grouped, on=['Documento', 'Categoria'], how='outer')
                    
                    df_cruce['Monto_Ops'] = df_cruce['Monto_Ops'].fillna(0.0).round(2)
                    df_cruce['Monto_Conta'] = df_cruce['Monto_Conta'].fillna(0.0).round(2)
                    df_cruce['Diferencia ($)'] = (df_cruce['Monto_Ops'] - df_cruce['Monto_Conta']).round(2)

                    docs_en_ops = set(df_ops_grouped['Documento'].unique())
                    docs_en_acc = set(df_acc_grouped['Documento'].unique())

                    def evaluar_auditoria(row):
                        doc = str(row['Documento'])
                        m_ops = row['Monto_Ops']
                        m_acc = row['Monto_Conta']
                        
                        if m_ops > 0 and m_acc == 0:
                            if doc in docs_en_acc: return "🔴 ERROR DE CUENTA CONTABLE (Categoría cruzada)"
                            else: return "🔴 NO CONTABILIZADO (Falta la Partida o Inventario)"
                                
                        elif m_acc > 0 and m_ops == 0:
                            if doc in docs_en_ops: return "🔴 ERROR DE CATEGORÍA OPERATIVA (Categoría cruzada)"
                            else: return "🟡 NO EN OPERACIONES (Puede ser un ajuste manual)"
                                
                        else:
                            if abs(row['Diferencia ($)']) > 0.05: return "🟠 DIFERENCIA DE MONTO"
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
                    df_cruce = df_cruce.sort_values(['Prioridad', 'Documento']).drop(columns=['Prioridad'])

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