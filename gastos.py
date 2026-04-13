import streamlit as st
import pandas as pd
import json
import base64
import re
from datetime import datetime, date, timedelta
from utils import obtener_dataframe, conectar_hoja, obtener_gmail

TIPOS_DTE = {"01": "FACTURA", "03": "CREDITO FISCAL", "05": "NOTA DE CREDITO", "07": "RETENCION"}

def mostrar_modulo_gastos():
    st.title("Panel de Control: Auditoría DTE")

    if "ultimo_archivo" not in st.session_state:
        st.session_state.ultimo_archivo = None
    if "mensaje_exito" in st.session_state:
        st.success(st.session_state.mensaje_exito)
        del st.session_state.mensaje_exito
    if "mensaje_error" in st.session_state:
        st.error(st.session_state.mensaje_error)
        del st.session_state.mensaje_error

    df_base_cruda = obtener_dataframe("Base_Proveedores")
    df_base = pd.DataFrame()
    if not df_base_cruda.empty and 'Nit' in df_base_cruda.columns and 'Unidad' in df_base_cruda.columns:
        df_base = df_base_cruda.copy()
        df_base['Nit'] = df_base['Nit'].astype(str).apply(lambda x: re.sub(r'\D', '', x)).apply(lambda x: x.zfill(14) if len(x) > 0 else x)
        df_base['Unidad'] = df_base['Unidad'].astype(str).str.upper().str.strip()

    df_historico = obtener_dataframe("Historico")
    codigos_historico = []
    if not df_historico.empty and 'Codigo' in df_historico.columns:
        codigos_historico = df_historico['Codigo'].astype(str).str.strip().str.upper().tolist()

    if 'pendientes' not in st.session_state: st.session_state.pendientes = []

    st.subheader("1. Bandeja de Entrada (Nuevos Documentos)")

    col_fecha, col_btn = st.columns([1, 3])
    with col_fecha:
        rango_sync = st.date_input("Rango a sincronizar (Inicio - Fin):", value=(datetime.now().date(), datetime.now().date()), key="date_sync")
    with col_btn:
        st.write("")
        st.write("")
        btn_sincronizar = st.button("Sincronizar Gmail")

    if btn_sincronizar:
        try:
            if isinstance(rango_sync, tuple) and len(rango_sync) == 2:
                s_inicio, s_fin = rango_sync
            elif isinstance(rango_sync, tuple) and len(rango_sync) == 1:
                s_inicio = s_fin = rango_sync[0]
            else:
                s_inicio = s_fin = rango_sync
        except:
            s_inicio = s_fin = datetime.now().date()
            
        service = obtener_gmail()
        if not service:
            st.warning("⚠️ Verifica la conexión a Gmail. No se pudo autenticar el servicio.")
        elif df_base.empty:
            st.warning("⚠️ Verifica que la pestaña 'Base_Proveedores' tenga datos cargados.")
        else:
            with st.spinner("Buscando facturas en Gmail para las fechas seleccionadas..."):
                try:
                    labels = service.users().labels().list(userId='me').execute()
                    t_id = next((l['id'] for l in labels.get('labels', []) if 'DTE_AUDITORIA' in l['name'].upper()), None)
                    
                    if t_id:
                        s_fin_gmail = s_fin + timedelta(days=5)
                        query_fecha = f"after:{s_inicio.strftime('%Y/%m/%d')} before:{s_fin_gmail.strftime('%Y/%m/%d')}"
                        
                        res = service.users().messages().list(userId='me', labelIds=[t_id], q=query_fecha, maxResults=100).execute()
                        msgs = res.get('messages', [])
                        
                        lista_nits_autorizados = df_base['Nit'].tolist()
                        temp = []
                        contador = 1
                        
                        for m in msgs:
                            msg = service.users().messages().get(userId='me', id=m['id']).execute()
                            for p in msg['payload'].get('parts', []):
                                if p.get('filename') and p['filename'].lower().endswith('.json'):
                                    att = service.users().messages().attachments().get(userId='me', messageId=m['id'], id=p['body']['attachmentId']).execute()
                                    
                                    raw_bytes = base64.urlsafe_b64decode(att['data'].encode('UTF-8'))
                                    try:
                                        json_str = raw_bytes.decode('utf-8-sig')
                                    except UnicodeDecodeError:
                                        json_str = raw_bytes.decode('latin-1')
                                    
                                    dte = json.loads(json_str)
                                    
                                    ident = dte.get('identificacion', {})
                                    codigo_gen = str(ident.get('codigoGeneracion', '')).strip().upper()
                                    numero_control = str(ident.get('numeroControl', 'N/A')).strip().upper()
                                    
                                    fecha_cruda = str(ident.get('fecEmi', 'N/A'))
                                    try:
                                        fecha_emision_dt = datetime.strptime(fecha_cruda, '%Y-%m-%d').date()
                                        fecha_emision = fecha_emision_dt.strftime('%d/%m/%Y')
                                    except:
                                        fecha_emision_dt = None
                                        fecha_emision = fecha_cruda
                                        
                                    if fecha_emision_dt:
                                        if not (s_inicio <= fecha_emision_dt <= s_fin):
                                            continue 
                                    
                                    if codigo_gen in codigos_historico:
                                        continue
                                    
                                    nit_crudo = str(dte.get('emisor', {}).get('nit', ''))
                                    nit_limpio = re.sub(r'\D', '', nit_crudo).zfill(14)
                                    
                                    if nit_limpio in lista_nits_autorizados:
                                        unidades_posibles = df_base[df_base['Nit'] == nit_limpio]['Unidad'].dropna().unique()
                                        if len(unidades_posibles) > 1: unidad_auto = "PENDIENTE"
                                        elif len(unidades_posibles) == 1: unidad_auto = unidades_posibles[0]
                                        else: unidad_auto = "PENDIENTE"
                                            
                                        resumen = dte.get('resumen', {})
                                        monto = (resumen.get('totalPagar') or resumen.get('montoTotalOperacion') or resumen.get('totalVenta') or resumen.get('valorModificacion') or 0.00)
                                        
                                        temp.append({
                                            "Descartar": False,
                                            "#": contador,
                                            "Fecha": fecha_emision,
                                            "Num_Control": numero_control,
                                            "Proveedor": dte.get('emisor', {}).get('nombre', 'N/A').upper(),
                                            "Unidad": unidad_auto,
                                            "Tipo": TIPOS_DTE.get(ident.get('tipoDte'), "OTROS"),
                                            "Monto ($)": f"{float(monto):.2f}",
                                            "Codigo": codigo_gen
                                        })
                                        contador += 1
                        
                        st.session_state.pendientes = temp
                        if len(temp) == 0:
                            st.info("ℹ️ Sincronización completada: No hay facturas nuevas pendientes en las fechas seleccionadas.")
                        else:
                            st.success(f"✅ Sincronización completada: {len(temp)} facturas nuevas encontradas.")
                    else:
                        st.error("⚠️ No se encontró la etiqueta 'DTE_AUDITORIA' en tu cuenta de Gmail.")
                except Exception as e: 
                    st.error(f"Error al sincronizar: {e}")

    if st.session_state.pendientes:
        st.write(f"**Documentos listos para clasificar ({len(st.session_state.pendientes)})**")
        
        columnas_ordenadas = ["Descartar", "#", "Fecha", "Num_Control", "Proveedor", "Unidad", "Tipo", "Monto ($)", "Codigo"]
        
        df_edit = st.data_editor(pd.DataFrame(st.session_state.pendientes), 
                                column_config={
                                    "Unidad": st.column_config.SelectboxColumn("Unidad", options=["CAFETERIA", "DESPENSA", "TERRAZA", "PENDIENTE"]),
                                    "Descartar": st.column_config.CheckboxColumn("Eliminar 🗑️", default=False)
                                },
                                column_order=columnas_ordenadas,
                                hide_index=True, use_container_width=True, key="editor_docs")
        
        col_espacio, col_btn_procesar = st.columns([4, 1])
        with col_btn_procesar:
            procesar_btn = st.button("Procesar Documentos", use_container_width=True)

        if procesar_btn:
            with st.spinner("Guardando en la nube de Google Sheets..."):
                try:
                    ws_hist = conectar_hoja("Historico")
                    if ws_hist:
                        df_a_guardar = df_edit[(df_edit['Unidad'] != "PENDIENTE") & (~df_edit['Descartar'])].copy()
                        df_eliminados = df_edit[df_edit['Descartar']].copy()
                        
                        df_final_procesar = pd.concat([df_a_guardar, df_eliminados])
                        
                        if not df_final_procesar.empty:
                            df_final_procesar['Fecha_Procesado'] = datetime.now().strftime("%Y-%m-%d")
                            
                            if not df_eliminados.empty:
                                df_final_procesar.loc[df_final_procesar['Descartar'] == True, 'Unidad'] = "ELIMINADO"
                            
                            df_final_procesar = df_final_procesar.drop(columns=['Descartar'])
                            
                            datos_hist = ws_hist.get_all_values()
                            if len(datos_hist) > 0:
                                encabezados = datos_hist[0]
                                for col in encabezados:
                                    if col not in df_final_procesar.columns:
                                        df_final_procesar[col] = 'PENDIENTE' if col == 'Estado_Nexus' else ''
                                df_final_procesar = df_final_procesar[encabezados]
                            
                            datos_a_insertar = [[str(item) for item in row] for row in df_final_procesar.values.tolist()]
                            
                            if len(datos_hist) == 0:
                                ws_hist.clear()
                                ws_hist.append_row(df_final_procesar.columns.tolist())
                            
                            ws_hist.append_rows(datos_a_insertar)
                            
                            obtener_dataframe.clear() 
                            
                            pendientes_restantes = df_edit[(df_edit['Unidad'] == "PENDIENTE") & (~df_edit['Descartar'])]
                            st.session_state.pendientes = pendientes_restantes.to_dict('records')
                            
                            mensajes_exito = []
                            if len(df_a_guardar) > 0: mensajes_exito.append(f"{len(df_a_guardar)} archivados")
                            if len(df_eliminados) > 0: mensajes_exito.append(f"{len(df_eliminados)} eliminados")
                            
                            st.session_state.mensaje_exito = f"Operación exitosa: {' y '.join(mensajes_exito)}."
                            
                            if "editor_docs" in st.session_state: del st.session_state["editor_docs"]
                            st.rerun()
                        else:
                            st.session_state.mensaje_error = "No has asignado ninguna unidad ni seleccionado documentos para eliminar."
                            st.rerun()
                except Exception as e:
                    st.session_state.mensaje_error = f"Error crítico de Google al guardar: {e}"
                    st.rerun()

    st.divider()
    st.subheader("2. Auditoría y KPI (Tiempo Real)")

    col_f1, col_f2 = st.columns(2)
    with col_f1: u_audit = st.selectbox("Unidad a auditar:", ["CAFETERIA", "DESPENSA", "TERRAZA"])
    with col_f2: 
        rango_fechas = st.date_input("Rango de fechas de documentos (Inicio - Fin):", value=(date(2025, 1, 1), datetime.now().date()))

    try:
        if isinstance(rango_fechas, tuple) and len(rango_fechas) == 2:
            f_inicio, f_fin = rango_fechas
        elif isinstance(rango_fechas, tuple) and len(rango_fechas) == 1:
            f_inicio = f_fin = rango_fechas[0]
        else:
            f_inicio = f_fin = rango_fechas
    except:
        f_inicio = f_fin = datetime.now().date()

    fecha_str = f"Del {f_inicio.strftime('%d/%m/%Y')} al {f_fin.strftime('%d/%m/%Y')}"

    st.markdown("**Alimentar Reporte de Compras**")
    archivo_excel = st.file_uploader("Sube el Excel de compras. El sistema guardará los reportados de forma PERMANENTE en la base.", type=["xlsx", "csv"])

    if archivo_excel:
        if st.session_state.ultimo_archivo != archivo_excel.name:
            try:
                archivo_excel.seek(0)
                df_excel = pd.read_csv(archivo_excel) if archivo_excel.name.endswith('.csv') else pd.read_excel(archivo_excel)
                cods_excel = [str(x).strip().upper() for x in df_excel.iloc[:, 1].dropna().unique()]
                
                with st.spinner("Actualizando registros permanentemente en Google Sheets..."):
                    ws_hist = conectar_hoja("Historico")
                    if ws_hist:
                        datos = ws_hist.get_all_values()
                        if len(datos) > 1:
                            df_h = pd.DataFrame(datos[1:], columns=datos[0])
                            if 'Estado_Nexus' not in df_h.columns: df_h['Estado_Nexus'] = 'PENDIENTE'
                            df_h['Codigo_Limpio'] = df_h['Codigo'].astype(str).str.strip().str.upper()
                            df_h.loc[df_h['Codigo_Limpio'].isin(cods_excel), 'Estado_Nexus'] = 'REPORTADO'
                            df_h = df_h.drop(columns=['Codigo_Limpio'])
                            datos_a_insertar = [[str(item) for item in row] for row in df_h.values.tolist()]
                            ws_hist.clear()
                            ws_hist.append_row(df_h.columns.tolist())
                            ws_hist.append_rows(datos_a_insertar)
                            obtener_dataframe.clear() 
                
                st.session_state.ultimo_archivo = archivo_excel.name
                st.success("✅ Base de datos actualizada con los nuevos registros. ¡Ya puedes quitar el archivo Excel sin problema!")
                st.rerun()
            except Exception as e:
                st.error(f"Error crítico al procesar el Excel: {e}")
        else:
            st.success(f"✅ Archivo '{archivo_excel.name}' ya guardado en la base. Puedes cerrarlo.")
    else:
        st.session_state.ultimo_archivo = None

    docs_en_bodega = 0
    num_encontrados = 0
    num_faltantes = 0
    porc = 0.0
    df_faltantes = pd.DataFrame()
    df_historico = obtener_dataframe("Historico")

    if not df_historico.empty and 'Unidad' in df_historico.columns and 'Fecha' in df_historico.columns:
        df_historico['Fecha_DT'] = pd.to_datetime(df_historico['Fecha'], format='%d/%m/%Y', errors='coerce')
        df_historico['Fecha_DT'] = df_historico['Fecha_DT'].fillna(pd.to_datetime(df_historico['Fecha'], errors='coerce'))
        df_filtrado = df_historico[(df_historico['Unidad'] == u_audit) & (df_historico['Fecha_DT'].dt.date >= f_inicio) & (df_historico['Fecha_DT'].dt.date <= f_fin)].copy()
        docs_en_bodega = len(df_filtrado)
        
        if docs_en_bodega > 0:
            if 'Estado_Nexus' in df_filtrado.columns: df_faltantes = df_filtrado[df_filtrado['Estado_Nexus'] != 'REPORTADO']
            else: df_faltantes = df_filtrado.copy()
            num_faltantes = len(df_faltantes)
            num_encontrados = docs_en_bodega - num_faltantes
            porc = (num_encontrados / docs_en_bodega * 100) if docs_en_bodega > 0 else 0.0

    st.markdown("### Resumen en Vivo")
    c1_k, c2_k, c3_k = st.columns(3)
    c1_k.metric("Docs en Histórico (Esperados)", docs_en_bodega)
    c2_k.metric("Documentos Reportados (Nexus)", num_encontrados)
    c3_k.metric("KPI Actual (Cumplimiento)", f"{porc:.1f}%")

    st.markdown("### Documentos Pendientes de Registrar")
    if docs_en_bodega == 0:
        st.info("No hay documentos en el Histórico para el rango de fechas seleccionado.")
    elif num_faltantes > 0:
        with st.expander(f"👁️ Ver los {num_faltantes} documentos pendientes"):
            cols_to_show = ["Proveedor", "Tipo", "Monto ($)", "Codigo"]
            if "Num_Control" in df_faltantes.columns: cols_to_show.insert(0, "Num_Control")
            if "Fecha" in df_faltantes.columns: cols_to_show.insert(0, "Fecha")
            st.dataframe(df_faltantes[cols_to_show], hide_index=True, use_container_width=True)
    else:
        st.success("¡Excelente! Todos los documentos del Histórico han sido ingresados.")

    st.write("---")
    if st.button("Guardar resultado actual en la base de datos (Registro_KPI)"):
        with st.spinner("Registrando KPI..."):
            try:
                ws_kpi_escribir = conectar_hoja("Registro_KPI")
                if ws_kpi_escribir:
                    datos_kpi_crudos = ws_kpi_escribir.get_all_values()
                    datos_kpi_reales = [fila for fila in datos_kpi_crudos if any(str(c).strip() for c in fila)]
                    if len(datos_kpi_reales) == 0:
                        ws_kpi_escribir.append_row(["Fecha", "Unidad", "Porcentaje", "Docs_Esperados", "Docs_Encontrados"])
                    ws_kpi_escribir.append_row([fecha_str, u_audit, round(porc, 1), docs_en_bodega, num_encontrados])
                    obtener_dataframe.clear()
                    st.session_state.mensaje_exito = "El porcentaje actual se ha guardado en tu base de datos."
                    st.rerun()
            except Exception as e: 
                st.session_state.mensaje_error = f"Error de Google al guardar KPI: {e}"
                st.rerun()