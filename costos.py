import streamlit as st

# Importamos las funciones principales de nuestros nuevos archivos
from costos_cafeteria import mostrar_modulo_costos as modulo_cafeteria
from costos_talleres import mostrar_modulo_costos as modulo_talleres

def mostrar_modulo_costos():
    # 1. El Menú de Decisión
    unidad = st.radio(
        "🎯 Selecciona la Unidad de Negocio a evaluar:",
        ["☕ Cafetería y Despensa", "🖨️ Talleres Gráficos"],
        horizontal=True
    )
    
    st.markdown("---")

    # 2. El Policía de Tráfico enviando a la sección correcta
    if unidad == "☕ Cafetería y Despensa":
        modulo_cafeteria()
    elif unidad == "🖨️ Talleres Gráficos":
        modulo_talleres()
        # Truco para forzar actualizacion de cache