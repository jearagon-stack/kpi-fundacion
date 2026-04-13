import streamlit as st
import pandas as pd

# Esta es la función que usa el módulo de COSTOS
def ejecutar_auditoria_costos(df_ini, df_com, df_fin, consumo_dict, ventas_mes, costo_real, mes, anio, unidad):
    st.markdown("---")
    st.header("🛡️ Aduana de Validación de Información")
    
    apto_para_cierre = True
    df_c = pd.DataFrame(list(consumo_dict.items()), columns=['Cuenta', 'Consumo'])

    col_v1, col_v2 = st.columns(2)
    with col_v1:
        negativos = df_c[df_c['Consumo'] < 0]
        if not negativos.empty:
            st.error(f"🚩 **ERROR CRÍTICO:** Se detectaron {len(negativos)} cuentas con consumo negativo.")
            st.dataframe(negativos)
            apto_para_cierre = False
        else:
            st.success("✅ Consumos consistentes (No hay negativos).")

        if ventas_mes > 0:
            margen_actual = (costo_real / ventas_mes)
            st.metric("Margen Costo/Venta Operativo", f"{margen_actual:.2%}")
            if 0.61 <= margen_actual <= 0.67:
                st.success("🟢 Margen dentro del rango histórico (62-65%).")
            else:
                st.warning(f"🟡 Margen fuera de rango ({margen_actual:.2%}).")

    with col_v2:
        st.write("**Top 5 Cuentas con Mayor Impacto:**")
        top_5 = df_c.nlargest(5, 'Consumo')
        st.bar_chart(top_5.set_index('Cuenta'))
        
    return apto_para_cierre

# ESTA ES LA FUNCIÓN QUE LE FALTA A TU ARCHIVO Y POR ESO DA ERROR
def mostrar_modulo_validacion():
    st.title("🛡️ Lector Independiente de Validación")
    st.markdown("---")
    st.info("Esta sección funcionará como un monitor independiente de los datos cargados en el módulo de costos.")
    
    # Aquí es donde pondremos el visor de auditoría que querés ver por separado
    st.write("Cargue los archivos en 'Contabilidad de Costos' para ver el análisis aquí.")