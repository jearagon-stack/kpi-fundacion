import streamlit as st
import pandas as pd
from utils import obtener_dataframe

def mostrar_modulo_validacion():
    st.title("🛡️ Protocolo de Verificación de Integridad")
    st.markdown("---")

    # 1. Verificación de datos en memoria
    if 'datos_auditoria' not in st.session_state:
        st.info("💡 Módulo en espera. Por favor, procese los cierres en el módulo 'Contabilidad de Costos' primero.")
        return

    data = st.session_state['datos_auditoria']
    consumo_dict = data.get('consumo', {})
    ventas_mes = data.get('ventas', 0.0)
    costo_real = data.get('costo_real', 0.0)
    
    # 2. Carga de parámetros desde Google Sheets
    try: 
        df_params = obtener_dataframe("Parametros_Auditoria")
    except: 
        df_params = pd.DataFrame()

    limite_var = 0.01
    limites_cat = {
        "LIMPIEZA": 3000.0, 
        "EMPAQUE": 5000.0, 
        "MATERIA_PRIMA": 10000.0, 
        "PRODUCTO_TERMINADO": 8000.0
    }

    if not df_params.empty:
        try:
            val_var = df_params[df_params['Criterio'] == 'VARIACION_MAX_PERMITIDA']['Valor_Tope']
            if not val_var.empty: limite_var = float(val_var.iloc[0])
            for cat in limites_cat.keys():
                val_cat = df_params[df_params['Criterio'].isin([cat, cat.replace("_", " ")])]['Valor_Tope']
                if not val_cat.empty: limites_cat[cat] = float(val_cat.iloc[0])
        except: 
            pass

    # 3. Recolector de anomalías para el freno de mano
    lista_errores = []
    df_c = pd.DataFrame(list(consumo_dict.items()), columns=['Cuenta', 'Consumo'])

    st.subheader("🚥 Semáforos de Auditoría")
    col_v1, col_v2 = st.columns(2)

    with col_v1:
        st.markdown("**1. Anomalías por Unidad/Caja (Saldos Negativos)**")
        negativos = df_c[df_c['Consumo'] < 0]
        if not negativos.empty:
            st.error(f"🚩 **DISCREPANCIA:** {len(negativos)} cuentas en rojo.")
            st.dataframe(negativos, use_container_width=True)
            lista_errores.append(f"Saldos negativos detectados en {len(negativos)} cuentas.")
        else:
            st.success("✅ Cuentas saneadas (Sin saldos negativos).")

        st.divider()
        
        st.markdown("**2. Rentabilidad Operativa**")
        if ventas_mes > 0:
            margen_actual = (costo_real / ventas_mes)
            if 0.61 <= margen_actual <= 0.67:
                st.success(f"✅ Ratio Costo/Ventas: {margen_actual:.2%} (Normal).")
            else:
                st.warning(f"🟡 Desviación en el ratio: {margen_actual:.2%}.")
                lista_errores.append(f"Ratio de Rentabilidad Atípico ({margen_actual:.2%})")
        else:
            st.info("Sin registro de ventas para calcular margen.")

    with col_v2:
        st.markdown(f"**3. Costo Promedio (Límite Variación: {limite_var:.2%})**")
        if 'variaciones_costo' in data:
            df_var = data['variaciones_costo']
            if not df_var.empty:
                try:
                    anomalias = df_var[df_var['Variacion_Porcentual'].abs() > limite_var]
                    if not anomalias.empty:
                        st.error(f"🚩 {len(anomalias)} productos superan el límite de variación.")
                        lista_errores.append(f"Variación de costo atípica en {len(anomalias)} productos.")
                        st.dataframe(anomalias, use_container_width=True)
                    else:
                        st.success("✅ Variación de costos estables.")
                except: 
                    st.info("Estructura de variaciones no compatible.")
            else:
                st.warning("⚠️ La matriz llegó del otro módulo, pero está VACÍA (Los códigos de producto del Inv. Inicial no coinciden con los del Final).")
        else:
            st.info("⏳ Aún no se detecta la matriz. ¿Procesaste los archivos de nuevo en el módulo de Costos?")

        st.divider()

        st.markdown("**4. Topes de Inversión por Categoría**")
        if 'inventario_final' in data:
            df_inv = data['inventario_final']
            # Búsqueda inteligente de columnas
            col_monto = next((c for c in df_inv.columns if str(c).upper() in ['MONTO', 'VALOR', 'TOTAL', 'COSTO_TOTAL']), None)
            col_cat = next((c for c in df_inv.columns if 'CATEGOR' in str(c).upper() or 'CUENTA' in str(c).upper()), None)
            
            if col_monto and col_cat:
                for cat, tope in limites_cat.items():
                    monto_cat = df_inv[df_inv[col_cat].astype(str).str.contains(cat.replace("_", " "), case=False, na=False)][col_monto].sum()
                    if monto_cat > tope:
                        st.error(f"🚩 **{cat.replace('_', ' ')}:** ${monto_cat:,.2f} (Tope: ${tope:,.2f})")
                        lista_errores.append(f"Sobreexistencia en {cat.replace('_', ' ')} (${monto_cat:,.2f})")
                    else:
                        st.success(f"✅ **{cat.replace('_', ' ')}:** ${monto_cat:,.2f} (Tope: ${tope:,.2f})")
            else:
                st.error(f"Columnas no detectadas. Nombres en tu archivo: {list(df_inv.columns)}")
        else:
            st.info("⏳ Esperando desglose de inventario final desde el módulo de Costos...")

# NUEVO: Tabla de Detalles para Auditoría
    st.markdown("---")
    st.subheader("📋 Detalles de Inventario Final (Materia Prima y Producto Terminado)")
    if 'inventario_final' in data:
        df_fin_view = data['inventario_final'].copy()
        
        # Omitimos productos con existencias en 0 para no hacer ruido
        try:
            col_cant_view = next((c for c in df_fin_view.columns if 'EXIST' in str(c).upper() or 'SALDO' in str(c).upper()), None)
            if col_cant_view:
                df_fin_view[col_cant_view] = pd.to_numeric(df_fin_view[col_cant_view], errors='coerce').fillna(0.0)
                df_fin_view = df_fin_view[df_fin_view[col_cant_view] > 0]
        except:
            pass
            
        st.dataframe(df_fin_view, use_container_width=True)
        
    st.markdown("---")

    # 4. Lógica del Freno de Mano
    if len(lista_errores) == 0:
        if st.button("✅ DAR VISTO BUENO (APROBAR PERIODO)", type="primary", use_container_width=True):
            st.session_state['auditoria_aprobada'] = True
            st.balloons()
            st.success("¡Periodo Aprobado! Regresa al módulo de Costos para guardar el cierre.")
    else:
        st.warning("⚠️ **SISTEMA BLOQUEADO: Se detectaron las siguientes anomalías:**")
        for err in lista_errores:
            st.markdown(f"- 🔸 {err}")
        
        st.write("")
        check_autorizacion = st.checkbox("Declaro que he revisado las anomalías y autorizo este cierre bajo mi responsabilidad.")
        
        if check_autorizacion:
            if st.button("🚨 CONFIRMAR Y GUARDAR EXCEPCIONES", type="primary", use_container_width=True):
                st.session_state['auditoria_aprobada'] = True
                st.success("¡Excepciones Aprobadas! Regresa al módulo de Costos para finalizar.")
        else:
            st.error("🔒 Debes marcar la casilla de autorización para habilitar el botón.")