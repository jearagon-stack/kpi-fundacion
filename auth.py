import streamlit as st
import pandas as pd
from utils import conectar_hoja

@st.cache_data(ttl=60, show_spinner=False)
def obtener_usuarios_db():
    ws = conectar_hoja("Usuarios")
    if ws:
        try:
            datos = ws.get_all_values()
            if len(datos) > 1:
                df = pd.DataFrame(datos[1:], columns=datos[0])
                if all(col in df.columns for col in ['Usuario', 'Contrasena', 'Rol']):
                    df['Usuario'] = df['Usuario'].astype(str).str.strip().str.upper()
                    df['Contrasena'] = df['Contrasena'].astype(str).str.strip()
                    df['Rol'] = df['Rol'].astype(str).str.strip().str.upper()
                    return df.set_index('Usuario').to_dict('index')
        except: pass
    
    return {
        "ADMINISTRADOR": {"Contrasena": "199_6", "Rol": "ADMIN"},
        "JAZMIN ZEPEDA": {"Contrasena": "FUNDACION01.", "Rol": "USUARIO"}
    }

