import streamlit as st
import pandas as pd
import json
import base64
import gspread
import time
from google.oauth2.service_account import Credentials as SheetCredentials
from google.oauth2.credentials import Credentials as GmailCredentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

URL_DOCUMENTO = "https://docs.google.com/spreadsheets/d/1WbUyws8uqRG7K3tPJebMa27dBjd5ChKHCe2kfGiOkwo/edit"

@st.cache_resource(ttl=600)
def obtener_cliente_sheets():
    try:
        llave_b64 = st.secrets["google_sheets"]["key_codificada"]
        llave_texto = base64.b64decode(llave_b64).decode("utf-8")
        creds = SheetCredentials.from_service_account_info(json.loads(llave_texto), 
                scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Error de conexión con Google: {e}")
        return None

def conectar_hoja(nombre_pestaña, reintentos=3):
    cliente = obtener_cliente_sheets()
    if cliente:
        for intento in range(reintentos):
            try:
                doc = cliente.open_by_url(URL_DOCUMENTO)
                return doc.worksheet(nombre_pestaña)
            except:
                if intento < reintentos - 1: time.sleep(1)
    return None

def obtener_gmail():
    if "google_token" in st.secrets:
        try:
            t_info = json.loads(st.secrets["google_token"]["contenido"])
            creds = GmailCredentials.from_authorized_user_info(t_info, ['https://www.googleapis.com/auth/gmail.readonly'])
            if creds and creds.expired and creds.refresh_token: creds.refresh(Request())
            return build('gmail', 'v1', credentials=creds)
        except: pass
    return None

@st.cache_data(ttl=60, show_spinner=False)
def obtener_dataframe(pestaña):
    ws = conectar_hoja(pestaña)
    if ws:
        try:
            datos = ws.get_all_values()
            if len(datos) > 1: return pd.DataFrame(datos[1:], columns=datos[0])
            elif len(datos) == 1: return pd.DataFrame(columns=datos[0])
        except: pass
    return pd.DataFrame()