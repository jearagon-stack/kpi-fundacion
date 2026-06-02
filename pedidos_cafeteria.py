import streamlit as st
import pandas as pd
from datetime import datetime
import io
from utils import conectar_hoja, obtener_dataframe

# --- FUNCIONES DE LÓGICA ---

def consolidar_carrito(df_carrito):
    """Agrupa los productos en el carrito para evitar duplicados en la visualización."""
    if df_carrito.empty:
        return df_carrito
    return df_carrito.groupby(["ID_Productos", "Descripcion", "Medida"], as_index=False)["Cantidad"].sum()

def guardar_en_sheets(df_envio):
    """Escribe los datos en la hoja Pedidos_Pendientes de Google Sheets."""
    try:
        ws = conectar_hoja("Pedidos_Pendientes")
        datos = df_envio.values.tolist()
        ws.append_rows(datos)
        return True
    except Exception as e:
        st.error(f"Error al escribir en Google Sheets: {e}")
        return False

def eliminar_pedido_procesado(id_pedido_a_borrar):
    """Elimina únicamente las filas del pedido procesado y conserva el resto."""
    try:
        ws = conectar_hoja("Pedidos_Pendientes")
        df_pendientes = obtener_dataframe("Pedidos_Pendientes")
        
        if not df_pendientes.empty:
            # Filtramos para conservar todo lo que NO sea el pedido procesado
            df_restantes = df_pendientes[df_pendientes['ID_Pedido'] != id_pedido_a_borrar]
            
            ws.clear()
            # Escribir nuevamente los encabezados
            encabezados = df_pendientes.columns.tolist()
            ws.append_row(encabezados)
            
            # Escribir los pedidos restantes si los hay
            if not df_restantes.empty:
                ws.append_rows(df_restantes.values.tolist())
        return True
    except Exception as e:
        st.error(f"Error al actualizar la base de datos: {e}")
        return False

# --- MÓDULO PRINCIPAL ---

def mostrar_modulo_pedidos():
    st.title("🛒 Gestión de Pedidos - Cafetería")
    
    anexo_usuario = st.session_state.get('anexo_actual', 'Desconocido')
    rol_usuario = st.session_state.get('rol_actual', 'CAJERA')

    if 'carrito_pedidos' not in st.session_state:
        st.session_state['carrito_pedidos'] = pd.DataFrame(columns=["ID_Productos", "Descripcion", "Medida", "Cantidad"])

    if rol_usuario in ["ADMIN", "BODEGUERO"]:
        tabs = st.tabs(["🛍️ 1. Crear Pedido (Cajas)", "📦 2. Gestión de Bodega"])
        tab_cajas, tab_bodega = tabs[0], tabs[1]
    else:
        tabs = st.tabs(["🛍️ 1. Crear Pedido (Cajas)"])
        tab_cajas = tabs[0]
        tab_bodega = None

    # --------------------------------------------------------
    # PESTAÑA 1: VISTA DE CAJERAS (CREACIÓN DE PEDIDO)
    # --------------------------------------------------------
    with tab_cajas:
        st.info(f"📍 Estás ingresando pedido para el anexo: **{anexo_usuario}**")
        
        try:
            df_cat = obtener_dataframe("Catalogo_Materiales")
            df_cat.columns = ["Categoria", "SubCategoria", "Nombre_Amigable", "ID_Productos", "Descripcion", "Medida"] + list(df_cat.columns[6:])
            
            col1, col2 = st.columns(2)
            with col1:
                categorias = df_cat['Categoria'].dropna().unique()
                cat_seleccionada = st.selectbox("Selecciona Categoría:", options=categorias)
            
            df_filtrado_1 = df_cat[df_cat['Categoria'] == cat_seleccionada]
            
            with col2:
                subcategorias = df_filtrado_1['SubCategoria'].dropna().unique()
                subcat_seleccionada = st.selectbox("Selecciona Sub Categoría:", options=subcategorias)
                
            df_filtrado_2 = df_filtrado_1[df_filtrado_1['SubCategoria'] == subcat_seleccionada]
            
            st.markdown("---")
            st.subheader(f"📦 Catálogo: {subcat_seleccionada}")
            
            cantidades_ingresadas = {}
            
            encabezado_prod, encabezado_cant = st.columns([3, 1])
            with encabezado_prod: st.markdown("**Producto**")
            with encabezado_cant: st.markdown("**Cantidad**")
            st.divider()

            for idx, row in df_filtrado_2.iterrows():
                c_prod, c_cant = st.columns([3, 1])
                with c_prod:
                    st.markdown(f"**{row['Nombre_Amigable']}**")
                    st.caption(f"Medida: {row['Medida']}")
                with c_cant:
                    cantidades_ingresadas[idx] = st.number_input(
                        "Cant", min_value=0, value=0, step=1,
                        key=f"prod_{row['ID_Productos']}_{idx}", label_visibility="collapsed"
                    )

            st.markdown("###")
            if st.button("➕ Agregar Seleccionados a la Lista", type="primary", use_container_width=True):
                nuevas_filas = []
                for idx, cant in cantidades_ingresadas.items():
                    if cant > 0:
                        fila_prod = df_filtrado_2.loc[idx]
                        nuevas_filas.append({
                            "ID_Productos": fila_prod['ID_Productos'],
                            "Descripcion": fila_prod['Descripcion'],
                            "Medida": fila_prod['Medida'],
                            "Cantidad": cant
                        })
                
                if nuevas_filas:
                    df_nuevas = pd.DataFrame(nuevas_filas)
                    carrito_actual = st.session_state['carrito_pedidos']
                    st.session_state['carrito_pedidos'] = consolidar_carrito(pd.concat([carrito_actual, df_nuevas], ignore_index=True))
                    st.success(f"Se agregaron {len(nuevas_filas)} ítems a la lista.")
                    st.rerun()
                else:
                    st.warning("Debes digitar al menos una cantidad mayor a 0.")

            if not st.session_state['carrito_pedidos'].empty:
                st.markdown("---")
                st.subheader("🛒 Lista Preliminar de Pedido")
                st.dataframe(st.session_state['carrito_pedidos'], use_container_width=True, hide_index=True)
                
                confirmar = st.checkbox("¿Estás segura de realizar este pedido?")
                
                col_env, col_can = st.columns([2, 1])
                with col_env:
                    if confirmar:
                        if st.button("🚀 Enviar Pedido a Bodega", type="primary", use_container_width=True):
                            fecha_hoy = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            id_pedido = f"PED-{datetime.now().strftime('%H%M%S')}"
                            
                            df_envio = st.session_state['carrito_pedidos'].copy()
                            df_envio.insert(0, 'Fecha', fecha_hoy)
                            df_envio.insert(1, 'ID_Pedido', id_pedido)
                            df_envio.insert(2, 'Anexo', anexo_usuario)
                            
                            if guardar_en_sheets(df_envio):
                                st.session_state['carrito_pedidos'] = pd.DataFrame(columns=["ID_Productos", "Descripcion", "Medida", "Cantidad"])
                                st.session_state['msg_exito'] = f"Pedido {id_pedido} enviado correctamente a bodega."
                                st.rerun()

                with col_can:
                    if st.button("🗑️ Limpiar Lista", use_container_width=True):
                        st.session_state['carrito_pedidos'] = pd.DataFrame(columns=["ID_Productos", "Descripcion", "Medida", "Cantidad"])
                        st.rerun()

            if 'msg_exito' in st.session_state:
                st.success(st.session_state['msg_exito'])
                del st.session_state['msg_exito']

        except Exception as e:
            st.error(f"Error en la carga: {e}")

    # --------------------------------------------------------
    # PESTAÑA 2: VISTA DE BODEGUERO (PROCESAMIENTO)
    # --------------------------------------------------------
    if tab_bodega is not None:
        with tab_bodega:
            col_titulo, col_recargar = st.columns([3, 1])
            with col_titulo:
                st.subheader("📋 Pedidos Pendientes en Bodega")
            with col_recargar:
                if st.button("🔄 Recargar", use_container_width=True):
                    st.cache_data.clear()
            
            try:
                df_pendientes = obtener_dataframe("Pedidos_Pendientes")
                
                if df_pendientes.empty:
                    st.info("No hay pedidos pendientes de procesar en este momento.")
                else:
                    # Agrupar pedidos únicos
                    pedidos_unicos = df_pendientes['ID_Pedido'].unique()
                    
                    # Crear diccionario para mostrar "ID - Anexo" en el selector
                    opciones_selector = {}
                    for pid in pedidos_unicos:
                        anexo_val = df_pendientes[df_pendientes['ID_Pedido'] == pid]['Anexo'].iloc[0]
                        opciones_selector[pid] = f"{pid} - {anexo_val}"
                    
                    pedido_seleccionado = st.selectbox(
                        "Selecciona un pedido para revisar:",
                        options=pedidos_unicos,
                        format_func=lambda x: opciones_selector[x]
                    )
                    
                    # Filtrar datos del pedido seleccionado
                    df_pedido_actual = df_pendientes[df_pendientes['ID_Pedido'] == pedido_seleccionado].copy()
                    anexo_actual = df_pedido_actual['Anexo'].iloc[0]
                    fecha_actual = df_pedido_actual['Fecha'].iloc[0]
                    
                    st.markdown(f"**Detalle del {opciones_selector[pedido_seleccionado]}**")
                    
                    # Ocultar columnas operativas para el editor
                    df_editor = df_pedido_actual.drop(columns=['Fecha', 'ID_Pedido', 'Anexo'])
                    
                    pedido_editado = st.data_editor(
                        df_editor,
                        num_rows="dynamic",
                        use_container_width=True,
                        key=f"editor_{pedido_seleccionado}"
                    )
                    
                    # Módulo para agregar productos por código Nexus
                    with st.expander("➕ Agregar producto adicional al pedido"):
                        df_cat_bod = obtener_dataframe("Catalogo_Materiales")
                        df_cat_bod.columns = ["Categoria", "SubCategoria", "Nombre_Amigable", "ID_Productos", "Descripcion", "Medida"] + list(df_cat_bod.columns[6:])
                        
                        # Crear columna de visualización combinada (Código - Descripción)
                        df_cat_bod['Display'] = df_cat_bod['ID_Productos'].astype(str) + " - " + df_cat_bod['Descripcion'].astype(str)
                        
                        prod_add = st.selectbox("Buscar por Código o Descripción Nexus:", df_cat_bod['Display'].dropna().unique())
                        cant_add = st.number_input("Cantidad a agregar:", min_value=1, value=1, step=1)
                        
                        if st.button("Guardar en el pedido actual"):
                            fila_prod = df_cat_bod[df_cat_bod['Display'] == prod_add].iloc[0]
                            df_nuevo_item = pd.DataFrame([{
                                'Fecha': fecha_actual,
                                'ID_Pedido': pedido_seleccionado,
                                'Anexo': anexo_actual,
                                'ID_Productos': fila_prod['ID_Productos'],
                                'Descripcion': fila_prod['Descripcion'],
                                'Medida': fila_prod['Medida'],
                                'Cantidad': cant_add
                            }])
                            if guardar_en_sheets(df_nuevo_item):
                                st.cache_data.clear()
                                st.success("Producto agregado correctamente.")
                                st.rerun()

                    st.markdown("---")
                    
                    # Procesamiento y Descarga
                    col_dl, col_del = st.columns([2, 1])
                    
                    with col_dl:
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            pedido_editado.to_excel(writer, index=False, sheet_name='Traslado_Nexus')
                        datos_excel = output.getvalue()
                        
                        st.download_button(
                            label=f"⬇️ Aprobar y Descargar ({pedido_seleccionado})",
                            data=datos_excel,
                            file_name=f"{pedido_seleccionado}_{anexo_actual}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        
                    with col_del:
                        if st.button("✔️ Marcar como Procesado", type="secondary", use_container_width=True):
                            if eliminar_pedido_procesado(pedido_seleccionado):
                                st.cache_data.clear()
                                st.success(f"Pedido procesado y removido de la bandeja.")
                                st.rerun()
                            
            except Exception as e:
                st.error(f"No se pudo leer la hoja de pedidos: {e}")