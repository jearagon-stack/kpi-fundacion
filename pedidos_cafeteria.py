import streamlit as st
import pandas as pd
import io

# IMPORTAMOS TU FUNCIÓN DE CONEXIÓN YA EXISTENTE
try:
    from utils import obtener_dataframe
except ImportError:
    st.error("⚠️ Error: No se encontró el módulo 'utils' para conectar con la base de datos.")

def consolidar_carrito(df_carrito):
    if df_carrito.empty:
        return df_carrito
    df_agrupado = df_carrito.groupby(["IdProducto", "Descripcion", "MEDIDA"], as_index=False)["Cantidad"].sum()
    return df_agrupado

def mostrar_modulo_pedidos():
    st.title("🛒 Gestión de Pedidos - Cafetería")
    
    # ========================================================
    # INICIALIZACIÓN DE MEMORIA (ESTADOS)
    # ========================================================
    # Carrito temporal de la cajera
    if 'carrito_pedidos' not in st.session_state:
        st.session_state['carrito_pedidos'] = pd.DataFrame(columns=["IdProducto", "Descripcion", "MEDIDA", "Cantidad"])
    
    # Bandeja de entrada del bodeguero
    if 'pedidos_enviados' not in st.session_state:
        st.session_state['pedidos_enviados'] = pd.DataFrame(columns=["IdProducto", "Descripcion", "MEDIDA", "Cantidad"])

    # ========================================================
    # CREACIÓN DE PESTAÑAS (TABS) POR ROL
    # ========================================================
    tab_cajas, tab_bodega = st.tabs(["🛍️ 1. Crear Pedido (Cajas)", "📦 2. Gestión de Bodega"])

    # --------------------------------------------------------
    # PESTAÑA 1: VISTA DE CAJERAS
    # --------------------------------------------------------
    with tab_cajas:
        st.info("Selecciona la categoría e ingresa las cantidades de los productos que necesites pedir.")
        try:
            df_cat = obtener_dataframe("Catalogo_Materiales")
            nombres_correctos = ["Categoria", "SubCategoria", "Nombre_Amigable", "IdProducto", "Descripcion", "MEDIDA"]
            df_cat.columns = nombres_correctos + list(df_cat.columns[6:])
            
            col1, col2 = st.columns(2)
            with col1:
                categorias = df_cat['Categoria'].dropna().unique()
                cat_seleccionada = st.selectbox("Categoría:", options=categorias)
            
            df_filtrado_1 = df_cat[df_cat['Categoria'] == cat_seleccionada]
            
            with col2:
                subcategorias = df_filtrado_1['SubCategoria'].dropna().unique()
                subcat_seleccionada = st.selectbox("Sub Categoría:", options=subcategorias)
                
            df_filtrado_2 = df_filtrado_1[df_filtrado_1['SubCategoria'] == subcat_seleccionada]
            
            st.markdown("---")
            st.subheader(f"📦 Catálogo: {subcat_seleccionada}")
            
            cantidades_ingresadas = {}
            
            encabezado_prod, encabezado_cant = st.columns([3, 1])
            with encabezado_prod: st.markdown("**Nombre del Producto**")
            with encabezado_cant: st.markdown("**Cantidad**")
                
            st.divider()

            for idx, row in df_filtrado_2.iterrows():
                c_prod, c_cant = st.columns([3, 1])
                with c_prod:
                    st.markdown(f"{row['Nombre_Amigable']} *(Medida: {row['MEDIDA']})*")
                with c_cant:
                    cantidades_ingresadas[idx] = st.number_input(
                        "Cantidad", min_value=0, value=0, step=1,
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
                    carrito_actual = pd.concat([carrito_actual, df_nuevas], ignore_index=True)
                    st.session_state['carrito_pedidos'] = consolidar_carrito(carrito_actual)
                    st.success(f"✅ Se agregaron {len(nuevas_filas)} productos a la lista preliminar.")
                else:
                    st.warning("⚠️ No ingresaste ninguna cantidad mayor a 0.")

            # RESUMEN Y CONFIRMACIÓN DE ENVÍO (Cajas)
            if not st.session_state['carrito_pedidos'].empty:
                st.markdown("---")
                st.subheader("🛒 Lista Preliminar del Pedido")
                st.dataframe(st.session_state['carrito_pedidos'], use_container_width=True, hide_index=True)
                
                st.warning("⚠️ ¿Estás segura de hacer este pedido? Una vez enviado a bodega, no podrás hacer cambios.")
                confirmar_envio = st.checkbox("Sí, estoy segura de enviar el pedido.")
                
                col_env, col_can = st.columns([2, 1])
                with col_env:
                    if confirmar_envio:
                        if st.button("🚀 Enviar Pedido a Bodega", type="primary", use_container_width=True):
                            # Pasar los datos a la bandeja del bodeguero
                            df_historico = st.session_state['pedidos_enviados']
                            df_consolidado = pd.concat([df_historico, st.session_state['carrito_pedidos']], ignore_index=True)
                            st.session_state['pedidos_enviados'] = consolidar_carrito(df_consolidado)
                            
                            # Limpiar la pantalla de la cajera
                            st.session_state['carrito_pedidos'] = pd.DataFrame(columns=["IdProducto", "Descripcion", "MEDIDA", "Cantidad"])
                            st.success("✅ El pedido fue enviado exitosamente a la bodega.")
                            st.rerun()
                with col_can:
                    if st.button("🗑️ Vaciar Lista", use_container_width=True):
                        st.session_state['carrito_pedidos'] = pd.DataFrame(columns=["IdProducto", "Descripcion", "MEDIDA", "Cantidad"])
                        st.rerun()

        except Exception as e:
            st.error(f"Error técnico. Detalle: {e}")

    # --------------------------------------------------------
    # PESTAÑA 2: VISTA DE BODEGUERO
    # --------------------------------------------------------
    with tab_bodega:
        st.subheader("📋 Pedidos Pendientes de Procesar")
        
        if st.session_state['pedidos_enviados'].empty:
            st.info("No hay pedidos entrantes en este momento.")
        else:
            st.markdown("Revisa el pedido. Puedes hacer doble clic en la tabla para **editar cantidades** o agregar filas nuevas si es necesario.")
            
            # Tabla editable para el bodeguero
            pedido_editado = st.data_editor(
                st.session_state['pedidos_enviados'],
                num_rows="dynamic", # Permite agregar filas nuevas manualmente
                use_container_width=True,
                key="editor_bodega"
            )
            
            st.markdown("---")
            col_dl, col_del = st.columns([2, 1])
            
            with col_dl:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    pedido_editado.to_excel(writer, index=False, sheet_name='Traslado_Nexus')
                datos_excel = output.getvalue()
                
                st.download_button(
                    label="⬇️ Aprobar y Descargar Archivo para Nexus",
                    data=datos_excel,
                    file_name="Pedido_Bodega_Aprobado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col_del:
                if st.button("✔️ Marcar como Procesado (Limpiar Bandeja)", use_container_width=True):
                    st.session_state['pedidos_enviados'] = pd.DataFrame(columns=["IdProducto", "Descripcion", "MEDIDA", "Cantidad"])
                    st.rerun()