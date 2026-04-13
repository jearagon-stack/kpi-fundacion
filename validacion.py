import streamlit as st
import pandas as pd

def ejecutar_auditoria_costos(df_ini, df_com, df_fin, consumo_dict, ventas_mes, costo_real, mes, anio, unidad):
    st.markdown("---")
    st.header("🛡️ Aduana de Validación de Información")
    
    apto_para_cierre = True
    df_c = pd.DataFrame(list(consumo_dict.items()), columns=['Cuenta', 'Consumo'])

    col_v1, col_v2 = st.columns(2)

    with col_v1:
        # --- FILTRO 1: CONSUMOS NEGATIVOS ---
        negativos = df_c[df_c['Consumo'] < 0]
        if not negativos.empty:
            st.error(f"🚩 **ERROR CRÍTICO:** Se detectaron {len(negativos)} cuentas con consumo negativo.")
            st.dataframe(negativos)
            apto_para_cierre = False
        else:
            st.success("✅ Consumos consistentes (No hay negativos).")

        # --- FILTRO 2: RANGO DE MARGEN (62% - 65%) ---
        if ventas_mes > 0:
            margen_actual = (costo_real / ventas_mes)
            st.metric("Margen Costo/Venta", f"{margen_actual:.2%}")
            if 0.61 <= margen_actual <= 0.66:
                st.success("🟢 Margen dentro del rango esperado (62-65%).")
            else:
                st.warning(f"🟡 Margen fuera de rango ({margen_actual:.2%}). Revisar facturas o traslados.")

    with col_v2:
        # --- FILTRO 3: TOP IMPACTO (DETECTOR CAJA VS UNIDAD) ---
        st.write("**Top 5 Cuentas con Mayor Impacto:**")
        top_5 = df_c.nlargest(5, 'Consumo')
        st.bar_chart(top_5.set_index('Cuenta'))
        
        for _, r in top_5.iterrows():
            if r['Consumo'] > (costo_real * 0.40):
                st.warning(f"🚨 **Alerta:** La cuenta {r['Cuenta']} representa el {(r['Consumo']/costo_real)*100:.1f}% del costo. ¿Error de Caja vs Unidad?")

    return apto_para_cierre