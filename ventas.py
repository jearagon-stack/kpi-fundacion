import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime, date, timedelta
from utils import obtener_dataframe, conectar_hoja

def generar_excel_bytes(df):
    """Genera un archivo Excel en memoria para poder descargarlo."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Ventas_Limpias')
    return output.getvalue()

def mostrar_modulo_ventas():
    st.title("📊 Panel de Control: KPI de Ventas")

    df_metas = obtener_dataframe("Metas_Ventas")
    df_unidades = obtener_dataframe("Unidades_Ventas")
    
    if df_metas.empty or df_unidades.empty:
        st.warning("⚠️ Faltan las pestañas 'Metas_Ventas' o 'Unidades_Ventas' en Google Sheets.")
        st.stop()

    df_metas.columns = df_metas.columns.str.strip()
    df_metas['Unidad'] = df_metas['Unidad'].astype(str).str.strip().str.upper()
    df_metas['Sub_Unidad'] = df_metas['Sub_Unidad'].astype(str).str.strip().str.upper()
    
    df_unidades['Unidad'] = df_unidades['Unidad'].astype(str).str.strip().str.upper()
    df_unidades['Sub_Unidad'] = df_unidades['Sub_Unidad'].astype(str).str.strip().str.upper()

    meses_dict = {
        1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
        7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
    }
    
    for num, nombre_mes in meses_dict.items():
        if nombre_mes in df_metas.columns:
            df_metas[nombre_mes] = pd.to_numeric(df_metas[nombre_mes], errors='coerce').fillna(0)

    # --- DEFINICIÓN DE PESTAÑAS ---
    tab_dashboard, tab_carga, tab_limpiador = st.tabs([
        "📈 Dashboard de KPI", 
        "📥 Ingresar Ventas Diarias",
        "🧹 Limpiador de Reportes POS"
    ])

    # --- PESTAÑA DE CARGA ---
    with tab_carga:
        st.subheader("Registrar ventas desde Nexus")
        st.info("El sistema ahora detecta el subsidio automáticamente leyendo la columna 'TipoCliente' de Nexus.")
        
        lista_macro_unidades = df_unidades['Unidad'].dropna().unique().tolist()
        
        col_c1, col_c2 = st.columns(2)
        with col_c1: unidad_seleccionada = st.selectbox("1. Macro-Unidad:", lista_macro_unidades, key="carga_macro")
        sub_unidades_filtradas = df_unidades[df_unidades['Unidad'] == unidad_seleccionada]['Sub_Unidad'].dropna().tolist()
        with col_c2: sub_unidad_seleccionada = st.selectbox("2. Anexo / Sub-Unidad:", sub_unidades_filtradas, key="carga_sub")

        archivo_ventas = st.file_uploader("3. Sube el Excel de Nexus (VentasFormaPago)", type=["xlsx", "xls", "csv"])

        if archivo_ventas:
            try:
                df_nexus = pd.read_csv(archivo_ventas) if archivo_ventas.name.endswith('.csv') else pd.read_excel(archivo_ventas)
                df_nexus.columns = df_nexus.columns.str.strip()
                ventas_por_fecha = {} 
                
                if 'Nombre' in df_nexus.columns and 'TotalAfecto' in df_nexus.columns and 'Fecha' in df_nexus.columns and 'TipoCliente' in df_nexus.columns:
                    df_nexus['Fecha_Limpia'] = pd.to_datetime(df_nexus['Fecha'], errors='coerce', dayfirst=True).dt.strftime('%d/%m/%Y')
                    
                    for index, fila in df_nexus.iterrows():
                        cliente = str(fila['Nombre']).strip().upper()
                        tipo_cliente = str(fila['TipoCliente']).strip().upper()
                        monto = pd.to_numeric(fila['TotalAfecto'], errors='coerce')
                        fecha_fila = str(fila['Fecha_Limpia'])
                        
                        if pd.isna(monto) or fecha_fila == 'nan': continue
                        if "UNIVERSIDAD CENTROAMERICANA" in cliente and monto >= 2000: continue 
                        
                        if tipo_cliente == 'UCA':
                            valor_real = (monto / 0.60)
                            subsidio_dia = valor_real - monto
                        else:
                            valor_real = monto
                            subsidio_dia = 0.0
                            
                        if fecha_fila not in ventas_por_fecha:
                            ventas_por_fecha[fecha_fila] = {"venta_real": 0.0, "subsidio": 0.0}
                            
                        ventas_por_fecha[fecha_fila]["venta_real"] += valor_real
                        ventas_por_fecha[fecha_fila]["subsidio"] += subsidio_dia
                    
                    if ventas_por_fecha:
                        st.write("### 📋 Resumen Calculado a Registrar")
                        df_hist_check = obtener_dataframe("Historico_Ventas")
                        fechas_duplicadas = []

                        if not df_hist_check.empty:
                            df_hist_check['Unidad'] = df_hist_check['Unidad'].astype(str).str.strip().str.upper()
                            df_hist_check['Sub_Unidad'] = df_hist_check['Sub_Unidad'].astype(str).str.strip().str.upper()
                            hist_filtrado = df_hist_check[(df_hist_check['Unidad'] == unidad_seleccionada) & (df_hist_check['Sub_Unidad'] == sub_unidad_seleccionada)]
                            fechas_ya_registradas = hist_filtrado['Fecha'].astype(str).tolist()
                            for f in ventas_por_fecha.keys():
                                if f in fechas_ya_registradas: fechas_duplicadas.append(f)

                        if fechas_duplicadas:
                            st.error("🛑 **¡ALTO! Posible duplicación de datos detectada.**")
                            st.warning(f"Las ventas para el/los día(s) **{', '.join(fechas_duplicadas)}** ya existen en la base de datos para el anexo {sub_unidad_seleccionada}.")
                            st.info("Para proteger el KPI, el botón de guardado ha sido bloqueado.")
                        else:
                            df_resumen = pd.DataFrame([{"Fecha": f, "Venta Neta Real": f"${datos['venta_real']:,.2f}", "Subsidio UCA": f"${datos['subsidio']:,.2f}"} for f, datos in ventas_por_fecha.items()])
                            st.dataframe(df_resumen, hide_index=True)
                            
                            if st.button("Guardar Registros en Base de Datos", use_container_width=True):
                                with st.spinner("Guardando en el Histórico de Ventas..."):
                                    ws_hist_ventas = conectar_hoja("Historico_Ventas")
                                    if ws_hist_ventas:
                                        datos_hist = ws_hist_ventas.get_all_values()
                                        if len(datos_hist) == 0: ws_hist_ventas.append_row(["Fecha", "Unidad", "Sub_Unidad", "Venta_Real", "Subsidio_UCA", "Usuario_Registro"])
                                        
                                        filas_a_insertar = [[f, unidad_seleccionada, sub_unidad_seleccionada, datos['venta_real'], datos['subsidio'], st.session_state.usuario_actual] for f, datos in ventas_por_fecha.items()]
                                        ws_hist_ventas.append_rows(filas_a_insertar)
                                        obtener_dataframe.clear()
                                        st.success(f"✅ ¡{len(filas_a_insertar)} registro(s) guardado(s) exitosamente!")
                    else:
                        st.warning("No se encontraron ventas válidas para procesar.")
                else:
                    st.error("⚠️ El archivo no contiene alguna de las columnas obligatorias: 'Fecha', 'Nombre', 'TotalAfecto', 'TipoCliente'.")
            except Exception as e: st.error(f"Error al leer el archivo: {e}")

    # --- PESTAÑA DE DASHBOARD ---
    with tab_dashboard:
        st.subheader("Filtros de Auditoría")
        
        col_f1, col_f2 = st.columns(2)
        lista_macro_dash = df_unidades['Unidad'].dropna().unique().tolist()
        with col_f1: u_audit_ventas = st.selectbox("Selecciona Unidad a auditar:", lista_macro_dash, key="dash_macro")
        
        sub_unidades_dash = ["CONSOLIDADO GLOBAL"] + df_unidades[df_unidades['Unidad'] == u_audit_ventas]['Sub_Unidad'].dropna().tolist()
        with col_f2: sub_audit_ventas = st.selectbox("Selecciona Anexo (Opcional):", sub_unidades_dash, key="dash_sub")

        st.write("---")
        tipo_filtro = st.radio("Temporalidad de la Meta:", ["Anual (YTD)", "Mensual", "Semanal", "Diario"], horizontal=True)

        if tipo_filtro == "Mensual":
            col_t1, col_t2 = st.columns(2)
            with col_t1: mes_sel = st.selectbox("Mes:", range(1, 13), index=datetime.now().month - 1)
            with col_t2: año_sel = st.selectbox("Año:", range(2024, 2030), index=datetime.now().year - 2024)
            fecha_inicio = date(año_sel, mes_sel, 1)
            fecha_fin = date(año_sel + 1, 1, 1) - timedelta(days=1) if mes_sel == 12 else date(año_sel, mes_sel + 1, 1) - timedelta(days=1)
            titulo_periodo = f"Acumulado Mensual: {fecha_inicio.strftime('%B %Y').upper()}"
            def meta_calc_func(row): return float(row.get(meses_dict.get(mes_sel, 'Enero'), 0))

        elif tipo_filtro == "Semanal":
            fecha_referencia = st.date_input("Selecciona un día de la semana a auditar:")
            fecha_inicio = fecha_referencia - timedelta(days=fecha_referencia.weekday())
            fecha_fin = fecha_inicio + timedelta(days=6)
            dias_en_mes_referencia = (date(fecha_inicio.year, fecha_inicio.month % 12 + 1, 1) - timedelta(days=1)).day if fecha_inicio.month < 12 else 31
            divisor_meta = dias_en_mes_referencia / 7 
            titulo_periodo = f"Semana: {fecha_inicio.strftime('%d/%m')} al {fecha_fin.strftime('%d/%m/%Y')}"
            def meta_calc_func(row): return float(row.get(meses_dict.get(fecha_inicio.month, 'Enero'), 0)) / divisor_meta

        elif tipo_filtro == "Diario":
            fecha_dia = st.date_input("Selecciona el día específico:")
            fecha_inicio = fecha_dia
            fecha_fin = fecha_dia
            
            año_actual = fecha_inicio.year
            mes_actual = fecha_inicio.month
            dias_en_mes = (date(año_actual, mes_actual % 12 + 1, 1) - timedelta(days=1)).day if mes_actual < 12 else 31
            
            peso_total_mes = 0.0
            for d in range(1, dias_en_mes + 1):
                dia_semana = date(año_actual, mes_actual, d).weekday()
                if dia_semana < 5:      peso_total_mes += 1.0
                elif dia_semana == 5:   peso_total_mes += 0.5
                    
            dia_seleccionado_semana = fecha_dia.weekday()
            if dia_seleccionado_semana < 5: peso_dia_seleccionado = 1.0
            elif dia_seleccionado_semana == 5: peso_dia_seleccionado = 0.5
            else: peso_dia_seleccionado = 0.0
            
            titulo_periodo = f"Día: {fecha_dia.strftime('%d/%m/%Y')}"
            
            def meta_calc_func(row): 
                meta_mensual = float(row.get(meses_dict.get(fecha_inicio.month, 'Enero'), 0))
                if peso_total_mes == 0: return 0.0
                return meta_mensual * (peso_dia_seleccionado / peso_total_mes)

        elif tipo_filtro == "Anual (YTD)":
            fecha_corte = st.date_input("Calcular acumulado desde el 1 de Enero hasta la fecha:", value=datetime.now().date())
            año_sel = fecha_corte.year
            fecha_inicio = date(año_sel, 1, 1)
            fecha_fin = fecha_corte
            titulo_periodo = f"Acumulado Anual (YTD) al: {fecha_fin.strftime('%d/%m/%Y')}"
            
            def meta_calc_func(row):
                total_ytd = 0.0
                for m in range(1, fecha_corte.month):
                    total_ytd += float(row.get(meses_dict[m], 0))
                dias_mes = (date(fecha_corte.year, fecha_corte.month % 12 + 1, 1) - timedelta(days=1)).day if fecha_corte.month < 12 else 31
                total_ytd += float(row.get(meses_dict[fecha_corte.month], 0)) * (fecha_corte.day / dias_mes)
                return total_ytd

        st.divider()
        vista_titulo = u_audit_ventas if sub_audit_ventas == "CONSOLIDADO GLOBAL" else sub_audit_ventas
        st.markdown(f"### {titulo_periodo} | {vista_titulo}")

        df_historico_vtas = obtener_dataframe("Historico_Ventas")

        if df_historico_vtas.empty:
            st.info("Aún no hay ventas registradas en el sistema para calcular el KPI.")
        else:
            df_historico_vtas['Unidad'] = df_historico_vtas['Unidad'].astype(str).str.strip().str.upper()
            df_historico_vtas['Sub_Unidad'] = df_historico_vtas['Sub_Unidad'].astype(str).str.strip().str.upper()
            df_historico_vtas['Fecha_DT'] = pd.to_datetime(df_historico_vtas['Fecha'], format='%d/%m/%Y', errors='coerce')
            df_historico_vtas['Venta_Real'] = pd.to_numeric(df_historico_vtas['Venta_Real'], errors='coerce').fillna(0)
            df_historico_vtas['Subsidio_UCA'] = pd.to_numeric(df_historico_vtas.get('Subsidio_UCA', 0), errors='coerce').fillna(0)
            
            filtro_base = (
                (df_historico_vtas['Unidad'] == u_audit_ventas) &
                (df_historico_vtas['Fecha_DT'].dt.date >= fecha_inicio) &
                (df_historico_vtas['Fecha_DT'].dt.date <= fecha_fin)
            )
            
            if sub_audit_ventas != "CONSOLIDADO GLOBAL":
                filtro_base = filtro_base & (df_historico_vtas['Sub_Unidad'] == sub_audit_ventas)
                metas_aplicables = df_metas[(df_metas['Unidad'] == u_audit_ventas) & (df_metas['Sub_Unidad'] == sub_audit_ventas)]
            else:
                metas_aplicables = df_metas[df_metas['Unidad'] == u_audit_ventas]

            df_filtrado = df_historico_vtas[filtro_base]

            meta_calculada = sum(meta_calc_func(row) for _, row in metas_aplicables.iterrows())
            venta_real_total = df_filtrado['Venta_Real'].sum()
            subsidio_total_aplicado = df_filtrado['Subsidio_UCA'].sum()
            kpi_calculado = (venta_real_total / meta_calculada * 100) if meta_calculada > 0 else 0.0

            col_k1, col_k2, col_k3, col_k4 = st.columns(4)
            col_k1.metric(f"Meta Neta {tipo_filtro.split()[0]}", f"${meta_calculada:,.2f}")
            col_k2.metric(f"Venta Neta Real", f"${venta_real_total:,.2f}")
            col_k3.metric("Subsidio UCA", f"${subsidio_total_aplicado:,.2f}")
            col_k4.metric("KPI Obtenido", f"{kpi_calculado:.1f}% {'✅' if kpi_calculado >= 100 else '⚠️' if kpi_calculado >= 85 else '❌'}")

            if sub_audit_ventas == "CONSOLIDADO GLOBAL":
                st.markdown("#### Desglose por Anexo")
                resultados_desglose = []
                
                for index, row in metas_aplicables.iterrows():
                    sub_u = row['Sub_Unidad']
                    meta_sub_calculada = meta_calc_func(row)
                    
                    df_sub_filtrado = df_filtrado[df_filtrado['Sub_Unidad'] == sub_u]
                    venta_sub = df_sub_filtrado['Venta_Real'].sum()
                    subsidio_sub = df_sub_filtrado['Subsidio_UCA'].sum()
                    
                    kpi_sub = (venta_sub / meta_sub_calculada * 100) if meta_sub_calculada > 0 else 0.0
                    
                    estado = "🟢 Excelente" if kpi_sub >= 100 else ("🟡 Regular" if kpi_sub >= 85 else "🔴 Crítico")
                    resultados_desglose.append({
                        "Anexo": sub_u, "Meta Asignada": f"${meta_sub_calculada:,.2f}",
                        "Venta Real": f"${venta_sub:,.2f}", "Subsidio": f"${subsidio_sub:,.2f}",
                        "KPI (%)": f"{kpi_sub:.1f}%", "Estado": estado
                    })

                if resultados_desglose:
                    st.dataframe(pd.DataFrame(resultados_desglose), hide_index=True, use_container_width=True)  

    # --- NUEVA PESTAÑA: LIMPIADOR DE REPORTES POS ---
    with tab_limpiador:
        st.subheader("Estructuración de Reporte por Categorías")
        st.markdown("Sube el archivo bruto del POS. El sistema consolidará los productos, agrupará las cantidades y limpiará los prefijos automáticos.")
        
        archivo_bruto = st.file_uploader("Sube el reporte de ventas (Excel o CSV)", type=["xlsx", "xls", "csv"], key="upload_limpiador")

        if archivo_bruto:
            with st.spinner("Procesando y estructurando datos..."):
                try:
                    # Lectura del archivo crudo
                    if archivo_bruto.name.endswith('.csv'):
                        df_bruto = pd.read_csv(archivo_bruto, header=None)
                    else:
                        df_bruto = pd.read_excel(archivo_bruto, header=None)

                    datos_procesados = []
                    sucursal_actual = "GENERAL"
                    categoria_actual = "SIN CATEGORIA"

                    # Iterar fila por fila simulando la lectura en cascada
                    for index, row in df_bruto.iterrows():
                        # La columna 0 es 'Descripción', la columna 1 es 'Tot. Vendido' (o similar)
                        desc = str(row[0]).strip()
                        
                        # Omitir filas vacías o encabezados inútiles
                        if not desc or desc.lower() in ['nan', 'none', 'descripción', 'pág. no.', 'informe de ventas']:
                            continue
                        if "del 1 de" in desc.lower() or "tot. vendido" in desc.lower():
                            continue

                        # Determinar si es un título (Sucursal/Categoría) o un Producto
                        val_tot = str(row[1]).strip()
                        if val_tot.lower() in ['nan', 'none', '']:
                            # Heurística para diferenciar Sucursal de Categoría
                            if "CAFETERÍA" in desc.upper() or "SUCURSAL" in desc.upper() or "CENTRAL" in desc.upper():
                                sucursal_actual = desc
                            else:
                                categoria_actual = desc
                        else:
                            # Es un producto. Limpiar prefijos estilo "-X3 ", "-X12 "
                            producto_limpio = re.sub(r'^-X\d+\s+', '', desc).strip()
                            
                            # Extraer valores numéricos
                            try:
                                total_vendido = float(str(row[1]).replace('$', '').replace(',', '').strip())
                                cnt_vendida = float(str(row[2]).replace(',', '').strip())
                            except ValueError:
                                continue 
                                
                            # Extraer el código del producto
                            cod_producto = ""
                            for col_idx in range(len(row)-1, 2, -1):
                                val_cod = str(row[col_idx]).strip()
                                if val_cod.startswith('PT'):
                                    cod_producto = val_cod
                                    break
                            
                            if pd.notna(total_vendido) and pd.notna(cnt_vendida):
                                datos_procesados.append({
                                    "Sucursal": sucursal_actual,
                                    "Categoría": categoria_actual,
                                    "Producto": producto_limpio,
                                    "Cantidad": cnt_vendida,
                                    "Total Ventas ($)": total_vendido,
                                    "Código": cod_producto
                                })

                    # Crear DataFrame limpio y agrupar para consolidar las repeticiones
                    df_limpio = pd.DataFrame(datos_procesados)
                    
                    if not df_limpio.empty:
                        df_agrupado = df_limpio.groupby(
                            ['Sucursal', 'Categoría', 'Código', 'Producto'], 
                            as_index=False
                        ).agg({
                            'Cantidad': 'sum',
                            'Total Ventas ($)': 'sum'
                        })

                        # Interfaz de Filtros
                        st.write("---")
                        st.markdown("### 🎛️ Filtros de Exportación")
                        
                        col_fil1, col_fil2 = st.columns(2)
                        with col_fil1:
                            sucursales_disp = df_agrupado['Sucursal'].unique().tolist()
                            sucursal_filtro = st.multiselect("Filtrar por Sucursal:", sucursales_disp, default=sucursales_disp)
                        with col_fil2:
                            categorias_disp = df_agrupado['Categoría'].unique().tolist()
                            categoria_filtro = st.multiselect("Filtrar por Categoría:", categorias_disp, default=categorias_disp)

                        # Aplicar filtros
                        df_final = df_agrupado[
                            (df_agrupado['Sucursal'].isin(sucursal_filtro)) & 
                            (df_agrupado['Categoría'].isin(categoria_filtro))
                        ]

                        st.write(f"**Total de registros filtrados:** {len(df_final)}")
                        st.dataframe(df_final, hide_index=True, use_container_width=True)

                        # Botón de exportación
                        st.write("---")
                        excel_data = generar_excel_bytes(df_final)
                        st.download_button(
                            label="📥 Descargar Datos Limpios (Excel)",
                            data=excel_data,
                            file_name=f"Ventas_Estructuradas_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.warning("No se pudieron extraer datos válidos del archivo. Verifica el formato original.")

                except Exception as e:
                    st.error(f"Error procesando la estructuración: {e}")