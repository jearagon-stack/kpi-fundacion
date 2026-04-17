import streamlit as st

# Importaciones de los módulos (AQUÍ ESTÁ LA CORRECCIÓN DE costs_library)
from costos_cafeteria import mostrar_modulo_costos as modulo_cafeteria
from costos_talleres import mostrar_modulo_costos as modulo_talleres
from costs_library import mostrar_modulo_libreria as modulo_libreria
from costs_soho import mostrar_modulo_soho as modulo_soho
from costs_terraza import mostrar_modulo_terraza as modulo_terraza
from costs_campus import mostrar_modulo_campus as modulo_campus
from costs_despensa import mostrar_modulo_despensa as modulo_despensa
from costs_gerencia import mostrar_modulo_gerencia as modulo_gerencia

def mostrar_modulo_costos():
    # El Menú con todas las unidades nuevas
    unidad = st.radio(
        "🎯 Selecciona la Unidad de Negocio a evaluar:",
        [
            "☕ Cafetería", 
            "🖨️ Talleres", 
            "📚 Librería", 
            "🏢 Centro Soho", 
            "🌅 Terraza", 
            "🏫 CID Campus", 
            "🛒 Despensa", 
            "💼 Gerencia Comercial"
        ],
        horizontal=True
    )
    
    st.markdown("---")

    # Lógica de ruteo
    if unidad == "☕ Cafetería":
        modulo_cafeteria()
    elif unidad == "🖨️ Talleres":
        modulo_talleres()
    elif unidad == "📚 Librería":
        modulo_libreria()
    elif unidad == "🏢 Centro Soho":
        modulo_soho()
    elif unidad == "🌅 Terraza":
        modulo_terraza()
    elif unidad == "🏫 CID Campus":
        modulo_campus()
    elif unidad == "🛒 Despensa":
        modulo_despensa()
    elif unidad == "💼 Gerencia Comercial":
        modulo_gerencia()