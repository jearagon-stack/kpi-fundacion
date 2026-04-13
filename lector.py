import imaplib
import email
import json
import pandas as pd
    
def escaneo_profundo_monto(data):
    """Busca cualquier rastro de dinero en el JSON si los campos normales fallan"""
    monto_encontrado = 0
    if isinstance(data, dict):
        for k, v in data.items():
            if any(p in k.lower() for p in ['totalpagar', 'montototal', 'total', 'montoitem']):
                try:
                    valor = float(v)
                    if valor > monto_encontrado: monto_encontrado = valor
                except: pass
            res = escaneo_profundo_monto(v)
            if res > monto_encontrado: monto_encontrado = res
    elif isinstance(data, list):
        for item in data:
            res = escaneo_profundo_monto(item)
            if res > monto_encontrado: monto_encontrado = res
    return monto_encontrado

def conectar_y_leer():
    usuario = "jearagon@fundagondra.sv" 
    password = "rtso fjqo vywf knbk" 
    TIPOS_DTE = {"01": "FACTURA", "03": "CREDITO FISCAL", "05": "NOTA CREDITO", "14": "SUJETO EXCLUIDO"}

    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(usuario, password)
        mail.select("DTE_AUDITORIA.") 
        _, data = mail.search(None, "ALL")
        ids = data[0].split()
        registros = []
        
        for i in ids:
            _, msg_data = mail.fetch(i, "(RFC822)")
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    for part in msg.walk():
                        if part.get_filename() and part.get_filename().endswith('.json'):
                            factura = json.loads(part.get_payload(decode=True))
                            
                            ident = factura.get('identificacion', {})
                            # 1. Recuperar Monto con Rayos X
                            monto = escaneo_profundo_monto(factura)
                            
                            # 2. Recuperar Tipo de Documento
                            tipo_txt = TIPOS_DTE.get(ident.get('tipoDte'), "OTROS")
                            
                            registros.append({
                                "Fecha": ident.get('fecEmi'),
                                "Tipo": tipo_txt,
                                "Proveedor": factura.get('emisor', {}).get('nombre'),
                                "Monto": float(monto),
                                "DTE_ID": ident.get('codigoGeneracion', 'N/A')
                            })
        mail.logout()
        return pd.DataFrame(registros)
    except Exception as e: return f"Error: {e}"

# --- REPORTE FINAL ---
df = conectar_y_leer()
print("\n" + "="*120)
print("                    SISTEMA DE AUDITORÍA AUTOMATIZADA - FUNDACIÓN GONDRA")
print("="*120)

if isinstance(df, pd.DataFrame) and not df.empty:
    pd.set_option('display.max_columns', None)
    pd.set_option('display.expand_frame_repr', False)
    pd.set_option('display.max_colwidth', None) # IMPORTANTE: Para que no corte nombres ni IDs
    
    # Ordenar por fecha para que sea más fácil de leer
    df = df.sort_values(by="Fecha")
    
    print(df[['Fecha', 'Tipo', 'Proveedor', 'Monto', 'DTE_ID']].to_string(index=False))
    
    print("-" * 120)
    print(f"DOCUMENTOS: {len(df)} | MONTO TOTAL ACUMULADO: ${df['Monto'].sum():,.2f}")
    print("="*120)
else:
    print("No se encontraron datos.")