import streamlit as st
import pandas as pd

def ejecutar_auditoria_costos(df_ini, df_com, df_fin, consumo_dict, ventas_mes, costo_real, mes, anio, unidad):
    st.markdown("---")
    st.header("🔍 Protocolo de Verificación de Integridad") # Antes: Aduana
    
    apto_para_cierre = True
    df_c = pd.DataFrame(list(consumo_dict.items()), columns=['Cuenta', 'Consumo'])

    col_v1, col_v2 = st.columns(2)
    with col_v1:
        # Sección 1: Análisis de Existencias
        st.subheader("📋 Consistencia de Inventarios")
        negativos = df_c[df_c['Consumo'] < 0]
        if not negativos.empty:
            st.error(f"🚩 **DISCREPANCIA:** Se detectaron {len(negativos)} cuentas con saldos inconsistentes (negativos).")
            st.dataframe(negativos)
            apto_para_cierre = False
        else:
            st.success("✅ Validación de saldos exitosa (Sin negativos).")

        # Sección 2: Margen de Contribución
        if ventas_mes > 0:
            margen_actual = (costo_real / ventas_mes)
            st.metric("Ratio Costo / Ventas (Real)", f"{margen_actual:.2%}")
            if 0.61 <= margen_actual <= 0.67:
                st.success("🟢 Margen de operación dentro de parámetros históricos.")
            else:
                st.warning(f"🟡 Desviación detectada en el ratio de operación ({margen_actual:.2%}).")

    with col_v2:
        # Sección 3: Pareto de Costos
        st.subheader("📈 Análisis de Concentración (Top 5)")
        top_5 = df_c.nlargest(5, 'Consumo')
        st.bar_chart(top_5.set_index('Cuenta'))
        
    return apto_para_cierre

def mostrar_modulo_validacion():
    st.title("📊 Sistema de Auditoría y Verificación Operativa") # Antes: Lector Independiente
    st.markdown("---")
    st.info("Este monitor de control centraliza la auditoría preventiva de los datos gestionados en el módulo de costos.")
    
    st.write("Cargue los estados de inventario en 'Contabilidad de Costos' para visualizar el diagnóstico detallado.")