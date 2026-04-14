import streamlit as st
import pandas as pd
from utils import obtener_dataframe

def mostrar_modulo_validacion():
    st.title("🛡️ Protocolo de Verificación de Integridad")
    st.markdown("---")

    # Verifica si hay datos enviados desde la pestaña de Costos
    if 'datos_auditoria' not in st.session_state:
        st.info("💡 Módulo en espera. Por favor, procese los cierres en el módulo 'Contabilidad de Costos' primero.")
        return

    # Extraemos los datos básicos actuales
    data = st.session_state['datos_auditoria']
    consumo_dict = data.get('consumo', {})
    ventas_mes = data.get('ventas', 0.0)
    costo_real = data.get('costo_real', 0.0)
    
    # -------------------------------------------------------------------------
    # 1. LECTURA DE PARÁMETROS DESDE GOOGLE SHEETS
    # -------------------------------------------------------------------------
    try:
        df_params = obtener_dataframe("Parametros_Auditoria")
    except:
        df_params = pd.DataFrame()

    # Valores por defecto por si la hoja está vacía o hay error de conexión
    limite_var = 0.01
    limites_cat = {
        "LIMPIEZA": 3000.0, 
        "EMPAQUE": 5000.0, 
        "MATERIA_PRIMA": 10000.0, 
        "PRODUCTO_TERMINADO": 7000.0
    }

    # Actualizamos con los valores que configuraste en la otra pantalla
    if not df_params.empty:
        try:
            val_var = df_params[df_params['Criterio'] == 'VARIACION_MAX_PERMITIDA']['Valor_Tope']
            if not val_var.empty: limite_var = float(val_var.iloc[0])
            
            for cat in limites_cat.keys():
                val_cat = df_params[df_params['Criterio'] == cat]['Valor_Tope']
                if not val_cat.empty:
                    limites_cat[cat] = float(val_cat.iloc[0])
        except:
            pass

    # -------------------------------------------------------------------------
    # 2. PANEL DE SEMÁFOROS Y ALERTAS
    # -------------------------------------------------------------------------
    alertas_criticas = 0  # Contador para activar el freno de mano
    df_c = pd.DataFrame(list(consumo_dict.items()), columns=['Cuenta', 'Consumo'])

    st.subheader("🚥 Semáforos de Auditoría")
    col_v1, col_v2 = st.columns(2)

    with col_v1:
        st.markdown("**1. Anomalías por Unidad/Caja (Saldos Negativos)**")
        negativos = df_c[df_c['Consumo'] < 0]
        if not negativos.empty:
            st.error(f"🚩 **DISCREPANCIA:** {len(negativos)} cuentas en rojo. Revisar conversiones de unidad.")
            st.dataframe(negativos, use_container_width=True)
            alertas_criticas += 1
        else:
            st.success("✅ Cuentas saneadas (Sin saldos negativos).")

        st.divider()
        
        st.markdown("**2. Rentabilidad Operativa**")
        if ventas_mes > 0:
            margen_actual = (costo_real / ventas_mes)
            if 0.61 <= margen_actual <= 0.67:
                st.success(f"✅ Ratio Costo/Ventas: {margen_actual:.2%} (Dentro de lo normal).")
            else:
                st.warning(f"🟡 Desviación en el ratio: {margen_actual:.2%}.")
                alertas_criticas += 1
        else:
            st.info("Sin registro de ventas para calcular margen.")

    with col_v2:
        st.markdown(f"**3. Costo Promedio (Límite Variación: {limite_var:.2%})**")
        # El módulo de costos debe enviar 'variaciones_costo'
        if 'variaciones_costo' in data and not data['variaciones_costo'].empty:
            df_var = data['variaciones_costo']
            try:
                anomalias = df_var[df_var['Variacion_Porcentual'].abs() > limite_var]
                if not anomalias.empty:
                    st.error(f"🚩 {len(anomalias)} productos superan el límite de variación.")
                    st.dataframe(anomalias, use_container_width=True)
                    alertas_criticas += 1
                else:
                    st.success("✅ Variación de costos estables.")
            except:
                st.info("Formato de variaciones no compatible.")
        else:
            st.info("⏳ Esperando matriz de variaciones desde el módulo de Costos...")

        st.divider()

        st.markdown("**4. Topes de Inversión por Categoría**")
        # El módulo de costos debe enviar 'inventario_final'
        if 'inventario_final' in data and not data['inventario_final'].empty:
            df_inv = data['inventario_final']
            try:
                for cat, tope in limites_cat.items():
                    # Sumamos el monto de la categoría actual
                    monto_cat = df_inv[df_inv['Categoria'].str.contains(cat.replace("_", " "), case=False, na=False)]['Monto'].sum()
                    if monto_cat > tope:
                        st.error(f"🚩 **{cat}:** ${monto_cat:,.2f} (Supera tope de ${tope:,.2f})")
                        alertas_criticas += 1
                    else:
                        st.success(f"✅ **{cat}:** ${monto_cat:,.2f} (Límite: ${tope:,.2f})")
            except:
                st.info("Estructura de inventario no compatible.")
        else:
            st.info("⏳ Esperando desglose de inventario final desde el módulo de Costos...")

    st.markdown("---")

    # -------------------------------------------------------------------------
    # 3. EL FRENO DE MANO (DOBLE CHECK DE SEGURIDAD)
    # -------------------------------------------------------------------------
    if alertas_criticas == 0:
        # Ruta Feliz: Todo está en verde
        if st.button("✅ DAR VISTO BUENO (APROBAR PERIODO)", type="primary", use_container_width=True):
            st.session_state['auditoria_aprobada'] = True
            st.balloons()
            st.success("¡Periodo Aprobado! Regresa al módulo de Costos para guardar en la base histórica.")
    else:
        # Ruta de Excepción: Hay alertas activas
        st.warning(f"⚠️ **SISTEMA BLOQUEADO:** Se detectaron {alertas_criticas} anomalías o topes excedidos en la auditoría.")
        
        # El Checkbox de responsabilidad
        check_autorizacion = st.checkbox("Declaro que he revisado las anomalías de costo/topes y autorizo este cierre bajo mi responsabilidad.")
        
        if check_autorizacion:
            # Solo si marca la casilla aparece el botón rojo
            if st.button("🚨 CONFIRMAR Y GUARDAR EXCEPCIONES", type="primary", use_container_width=True):
                st.session_state['auditoria_aprobada'] = True
                st.success("¡Excepciones Aprobadas! Regresa al módulo de Costos para guardar en la base histórica.")
        else:
            st.error("🔒 Debes marcar la casilla de autorización arriba para poder habilitar el botón de guardado.")