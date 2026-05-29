import streamlit as st

from costos_cafeteria import mostrar_modulo_costos as modulo_cafeteria
from costos_talleres import mostrar_modulo_costos as modulo_talleres
from costs_library import mostrar_modulo_libreria as modulo_libreria

def mostrar_modulo_costos():
    # El Menú completo
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
            "💼 Gerencia Comercial",
                    ],
        horizontal=True
    )
    
    st.markdown("---")

    # Lógica de ruteo con protección
    if unidad == "☕ Cafetería":
        modulo_cafeteria()
    elif unidad == "🖨️ Talleres":
        modulo_talleres()
    elif unidad == "📚 Librería":
        modulo_libreria()
    elif unidad == "🏢 Centro Soho":
        try:
            from costs_soho import mostrar_modulo_soho
            mostrar_modulo_soho()
        except ImportError:
            st.warning("⚠️ El archivo 'costs_soho.py' aún no ha sido creado...")
    elif unidad == "🌅 Terraza":
        try:
            from costs_terraza import mostrar_modulo_terraza
            mostrar_modulo_terraza()
        except ImportError:
            st.warning("⚠️ El archivo 'costs_terraza.py' aún no ha sido creado...")
    elif unidad == "🏫 CID Campus":
        try:
            from costs_campus import mostrar_modulo_cid_campus
            mostrar_modulo_cid_campus()
        except ImportError:
            st.warning("⚠️ El archivo 'costs_campus.py' aún no ha sido creado...")
    elif unidad == "🛒 Despensa":
        try:
            from costs_despensa import mostrar_modulo_despensa
            mostrar_modulo_despensa()
        except ImportError:
            st.warning("⚠️ El archivo 'costs_despensa.py' aún no ha sido creado...")
    elif unidad == "💼 Gerencia Comercial":
        try:
            from costs_gerencia import mostrar_modulo_gerencia
            mostrar_modulo_gerencia()
        except ImportError:
            st.warning("⚠️ El archivo 'costs_gerencia.py' aún no ha sido creado...")
            