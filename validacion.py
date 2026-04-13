import streamlit as st
import pandas as pd

def ejecutar_auditoria_costos(df_ini, df_com, df_fin, consumo_dict, ventas_totales, costo_real_total):
    st.markdown("---")
    st.markdown("### 🛡️ Panel de Validación de Información")
    
    # Convertimos el diccionario de consumo a DataFrame para analizarlo
    df_c = pd.DataFrame(list(consumo_dict.items()), columns=['Cuenta', 'Consumo'])
    
    col_v1, col_v2 = st.columns(2)
    
    with col_v1:
        # --- FILTRO 1: CONSUMOS NEGATIVOS ---
        negativos = df_c[df_c['Consumo'] < 0]
        if not negativos.empty:
            st.error(f"🚩 **Error de Conteo:** Hay {len(negativos)} cuentas con consumo negativo.")
            st.dataframe(negativos)
        else:
            st.success("✅ No hay consumos negativos.")

        # --- FILTRO 2: RELACIÓN DE MARGEN ---
        if ventas_totales > 0:
            margen = (ventas_totales - costo_real_total) / ventas_totales
            st.metric("Margen de Contribución Real", f"{margen:.2%}")
            if margen < 0.15:
                st.warning("⚠️ El margen es muy bajo (<15%). Revisa si faltan ventas por reportar o si el costo está inflado.")
            elif margen > 0.60:
                st.info("💡 Margen muy alto (>60%). Verifica si falta cargar alguna factura de compra.")

    with col_v2:
        # --- FILTRO 3: PARETO DE IMPACTO (TOP 5) ---
        st.write("**Top 5 Cuentas con Mayor Gasto:**")
        top_5 = df_c.nlargest(5, 'Consumo')
        st.bar_chart(top_5.set_index('Cuenta'))

    # --- FILTRO 4: ALERTAS DE "CAJA VS UNIDAD" (Basado en variaciones extremas) ---
    # Aquí comparamos si una sola cuenta se lleva más del 40% del costo total
    for index, row in top_5.iterrows():
        participacion = row['Consumo'] / costo_real_total if costo_real_total > 0 else 0
        if participacion > 0.40:
            st.warning(f"🚨 **Alerta de Concentración:** La cuenta {row['Cuenta']} representa el {participacion:.1%} del costo total. ¿Es normal o hay un error de unidad de medida?")