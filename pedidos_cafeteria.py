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
    # Agrupamos por los campos clave y sumamos cantidades
    return df_carrito.groupby(["IdProducto", "Descripcion", "MEDIDA"], as_index=False)["Cantidad"].sum()

def guardar_en_sheets(df_envio):
    """Escribe los datos en la hoja Pedidos_Pendientes de Google Sheets."""
    try:
        ws = conectar_hoja("Pedidos_Pendientes")
        # Convertimos el DataFrame a lista de listas para subirlo a la hoja
        datos = df_envio.values.tolist()
        ws.append_rows(datos)
        return True
    except Exception as e:
        st.error(f"Error al escribir en Google Sheets: {e}")
        return False

# --- MÓDULO PRINCIPAL ---

def mostrar_modulo_pedidos():
    st.title("🛒 Gestión de Pedidos - Cafetería")
    
    # Obtener variables de sesión
    anexo_usuario = st.session_state.get('anexo_actual', 'Desconocido')
    rol_usuario = st.session_state.get('rol_actual', 'CAJERA')

    # Inicialización de estados de sesión para mantener la persistencia
    if 'carrito_pedidos' not in st.session_state:
        st.session_state['carrito_pedidos'] = pd.DataFrame(columns=["IdProducto", "Descripcion", "MEDIDA", "Cantidad"])

    # Permisos de Pestañas
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
            # Lectura del catálogo
            df_cat = obtener_dataframe("Catalogo_Materiales")
            # Forzamos nombres de columnas para asegurar compatibilidad
            df_cat.columns = ["Categoria", "SubCategoria", "Nombre_Amigable", "IdProducto", "Descripcion", "MEDIDA"] + list(df_cat.columns[6:])
            
            # Filtros de búsqueda
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
            
            # Formulario rápido para digitar cantidades
            cantidades_ingresadas = {}
            
            # Encabezados
            encabezado_prod, encabezado_cant = st.columns([3, 1])
            with encabezado_prod: st.markdown("**Producto**")
            with encabezado_cant: st.markdown("**Cantidad**")
            st.divider()

            # Lista de productos iterables
            for idx, row in df_filtrado_2.iterrows():
                c_prod, c_cant = st.columns([3, 1])
                with c_prod:
                    st.markdown(f"**{row['Nombre_Amigable']}**")
                    st.caption(f"Medida: {row['MEDIDA']}")
                with c_cant:
                    # El key es único por ID de producto para evitar conflictos
                    cantidades_ingresadas[idx] = st.number_input(
                        "Cant", min_value=0, value=0, step=1,
                        key=f"prod_{row['IdProducto']}_{idx}", label_visibility="collapsed"
                    )

            st.markdown("###")
            if st.button("➕ Agregar Seleccionados a la Lista", type="primary", use_container_width=True):
                nuevas_filas = []
                for idx, cant in cantidades_ingresadas.items():
                    if cant > 0:
                        fila_prod = df_filtrado_2.loc[idx]
                        nuevas_filas.append({
                            "IdProducto": fila_prod['IdProducto'],
                            "Descripcion": fila_prod['Descripcion'],
                            "MEDIDA": fila_prod['MEDIDA'],
                            "Cantidad": cant
                        })
                
                if nuevas_filas:
                    df_nuevas = pd.DataFrame(nuevas_filas)
                    carrito_actual = st.session_state['carrito_pedidos']
                    # Unimos y consolidamos
                    st.session_state['carrito_pedidos'] = consolidar_carrito(pd.concat([carrito_actual, df_nuevas], ignore_index=True))
                    st.success(f"✅ Se agregaron {len(nuevas_filas)} ítems a la lista.")
                    st.rerun()
                else:
                    st.warning("⚠️ Debes digitar al menos una cantidad mayor a 0.")

            # Resumen y envío
            if not st.session_state['carrito_pedidos'].empty:
                st.markdown("---")
                st.subheader("🛒 Lista Preliminar de Pedido")
                st.dataframe(st.session_state['carrito_pedidos'], use_container_width=True, hide_index=True)
                
                st.warning("⚠️ ¿Estás segura de enviar este pedido? Una vez enviado, no podrás editarlo desde esta pantalla.")
                
                if st.button("🚀 Enviar Pedido a Bodega", type="primary", use_container_width=True):
                    # 1. Preparar datos
                    fecha_hoy = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    id_pedido = f"PED-{datetime.now().strftime('%H%M%S')}"
                    
                    df_envio = st.session_state['carrito_pedidos'].copy()
                    # Insertar columnas en orden exacto según lo que definiste en tu Sheets
                    df_envio.insert(0, 'Fecha', fecha_hoy)
                    df_envio.insert(1, 'ID_Pedido', id_pedido)
                    df_envio.insert(2, 'Anexo', anexo_usuario)
                    
                    # 2. Guardar en Sheets
                    if guardar_en_sheets(df_envio):
                        st.session_state['carrito_pedidos'] = pd.DataFrame(columns=["IdProducto", "Descripcion", "MEDIDA", "Cantidad"])
                        st.success(f"✅ Pedido {id_pedido} enviado correctamente a bodega.")
                        st.rerun()

                if st.button("🗑️ Limpiar Lista"):
                    st.session_state['carrito_pedidos'] = pd.DataFrame(columns=["IdProducto", "Descripcion", "MEDIDA", "Cantidad"])
                    st.rerun()

        except Exception as e:
            st.error(f"Error en la carga: {e}")

    # --------------------------------------------------------
    # PESTAÑA 2: VISTA DE BODEGUERO (PROCESAMIENTO)
    # --------------------------------------------------------
    if tab_bodega is not None:
        with tab_bodega:
            st.subheader("📋 Pedidos Pendientes en Bodega")
            
            try:
                # Leer desde Sheets en tiempo real
                df_pendientes = obtener_dataframe("Pedidos_Pendientes")
                
                if df_pendientes.empty:
                    st.info("No hay pedidos pendientes de procesar en este momento.")
                else:
                    st.markdown("Puedes editar cantidades directamente en la tabla o agregar nuevas filas.")
                    
                    # Editor interactivo para el bodeguero
                    pedido_editado = st.data_editor(
                        df_pendientes,
                        num_rows="dynamic",
                        use_container_width=True,
                        key="editor_bodega"
                    )
                    
                    st.markdown("---")
                    
                    # Botón de exportación (Aprobar)
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        pedido_editado.to_excel(writer, index=False, sheet_name='Traslado_Nexus')
                    datos_excel = output.getvalue()
                    
                    st.download_button(
                        label="⬇️ Aprobar y Descargar Archivo para Nexus",
                        data=datos_excel,
                        file_name=f"Pedido_Aprobado_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                    if st.button("✔️ Marcar todos como Procesados (Limpiar lista)", type="secondary"):
                        ws = conectar_hoja("Pedidos_Pendientes")
                        ws.clear()
                        # Volvemos a poner los encabezados para que la estructura no se rompa
                        ws.append_row(["Fecha", "ID_Pedido", "Anexo", "IdProducto", "Descripcion", "MEDIDA", "Cantidad"])
                        st.success("Bandeja limpiada correctamente.")
                        st.rerun()
                        
            except Exception as e:
                st.error(f"No se pudo leer la hoja de pedidos: {e}")