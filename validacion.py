import streamlit as st
import pandas as pd

def ejecutar_auditoria_completa(df_ini, df_com, df_fin, consumo_dict, ventas_mes, costo_real, mes, anio, unidad):
    st.markdown("---")
    st.header("🛡️ Aduana de Validación de Información")
    
    # 1. VARIABLE DE CONTROL PARA EL VISTO BUENO
    apto_para_cierre = True
    df_c = pd.DataFrame(list(consumo_dict.items()), columns=['Cuenta', 'Consumo'])

    # --- FILTRO 1: SEGURO DE DUPLICIDAD ---
    # (Aquí deberías jalar el df_resumen para comparar)
    
    # --- FILTRO 2: CONSUMOS NEGATIVOS ---
    negativos = df_c[df_c['Consumo'] < 0]
    if not negativos.empty:
        st.error(f"🚩 **ERROR CRÍTICO:** Se detectaron {len(negativos)} cuentas con consumo negativo. Revisa el conteo final.")
        st.dataframe(negativos)
        apto_para_cierre = False
    else:
        st.success("✅ Consumos consistentes (No hay negativos).")

    # --- FILTRO 3: ANOMALÍA CAJA VS UNIDAD (PARETO DE IMPACTO) ---
    st.subheader("📊 Análisis de Impacto y Unidades")
    top_5 = df_c.nlargest(5, 'Consumo')
    st.bar_chart(top_10.set_index('Cuenta')['Consumo'])
    
    for _, r in top_5.iterrows():
        # Si una cuenta representa más del 45% del costo total, lanzamos advertencia
        if r['Consumo'] > (costo_real * 0.45):
            st.warning(f"⚠️ **Alerta de Concentración:** La cuenta {r['Cuenta']} tiene un peso excesivo. Verificar si es un error de Unidad de Medida (Caja/Unidad).")

    # --- FILTRO 4: VALIDACIÓN DE MARGEN (62% - 65%) ---
    if ventas_mes > 0:
        margen_calculado = (costo_real / ventas_mes)
        st.metric("Margen Costo/Venta Operativo", f"{margen_calculado:.2%}")
        
        if 0.61 <= margen_calculado <= 0.66:
            st.success("🟢 Margen dentro del rango histórico (62-65%).")
        else:
            st.warning(f"🟡 Margen fuera de rango ({margen_calculado:.2%}). Revisa si faltan facturas o si hay duplicidad.")

    # --- VERDICTO FINAL ---
    if apto_para_cierre:
        st.balloons()
        st.success("### ✅ VERDICTO: INFORMACIÓN APTA PARA CIERRE")
        return True
    else:
        st.error("### ❌ VERDICTO: INFORMACIÓN NO APTA. CORRIJA LOS ERRORES.")
        return False