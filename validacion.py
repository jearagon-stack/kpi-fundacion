import streamlit as st
import pandas as pd

def mostrar_modulo_validacion():
    st.title("🛡️ Protocolo de Verificación de Integridad")
    st.markdown("---")

    # Verifica si hay datos enviados desde la pestaña de Costos
    if 'datos_auditoria' not in st.session_state:
        st.info("💡 Módulo en espera. Por favor, cargue y procese los archivos en el módulo 'Contabilidad de Costos' primero.")
        return

    # Extraemos los datos de la memoria
    data = st.session_state['datos_auditoria']
    consumo_dict = data['consumo']
    ventas_mes = data['ventas']
    costo_real = data['costo_real']

    apto_para_cierre = True
    df_c = pd.DataFrame(list(consumo_dict.items()), columns=['Cuenta', 'Consumo'])

    col_v1, col_v2 = st.columns(2)

    with col_v1:
        st.subheader("📋 Consistencia de Inventarios")
        negativos = df_c[df_c['Consumo'] < 0]
        if not negativos.empty:
            st.error(f"🚩 **DISCREPANCIA:** Se detectaron {len(negativos)} cuentas con saldos negativos.")
            st.dataframe(negativos)
            apto_para_cierre = False
        else:
            st.success("✅ Validación de saldos exitosa (Sin negativos).")

        st.markdown("<br>", unsafe_allow_html=True)
        if ventas_mes > 0:
            margen_actual = (costo_real / ventas_mes)
            st.metric("Ratio Costo / Ventas (Real)", f"{margen_actual:.2%}")
            if 0.61 <= margen_actual <= 0.67:
                st.success("🟢 Margen de operación dentro de parámetros históricos.")
            else:
                st.warning(f"🟡 Desviación detectada en el ratio de operación ({margen_actual:.2%}).")

    with col_v2:
        st.subheader("📈 Análisis de Concentración (Top 5)")
        # Limpiar basuras para el gráfico
        df_clean = df_c[~df_c['Cuenta'].astype(str).str.lower().isin(['nan', 'nat', 'no aplica'])]
        top_5 = df_clean.nlargest(5, 'Consumo')
        st.bar_chart(top_5.set_index('Cuenta'))

    st.divider()

    # Botón de Aprobación
    if apto_para_cierre:
        if st.button("✅ DAR VISTO BUENO (APROBAR PERIODO)", type="primary", use_container_width=True):
            st.session_state['auditoria_aprobada'] = True
            st.balloons()
            st.success("¡Periodo Aprobado! Regresa al módulo de 'Contabilidad de Costos' para descargar las partidas y cerrar la base.")
    else:
        st.error("❌ No es posible aprobar el periodo con errores de integridad. Revise los saldos negativos en caja/unidad.")