import streamlit as st
import pandas as pd
import io

# Función para consolidar el pedido (agrupar duplicados)
def consolidar_carrito(df_carrito):
    if df_carrito.empty:
        return df_carrito
    # Agrupa por código y suma las cantidades, manteniendo las demás descripciones
    df_agrupado = df_carrito.groupby(["Codigo_Nexus", "Descripcion_Nexus", "Unidad_Medida"], as_index=False)["Cantidad"].sum()
    return df_agrupado

def mostrar_modulo_pedidos():
    st.title("🛒 Pedidos de Materiales - Cafetería")
    st.info("Selecciona los productos para generar la solicitud de traslado de bodega.")

    # 1. MEMORIA DEL CARRITO (Session State)
    if 'carrito_pedidos' not in st.session_state:
        st.session_state['carrito_pedidos'] = pd.DataFrame(columns=["Codigo_Nexus", "Descripcion_Nexus", "Unidad_Medida", "Cantidad"])

    # 2. CARGA DEL CATÁLOGO
    # Para la fase de prueba, usamos un uploader. 
    # (Luego podemos automatizarlo para que lea el Google Sheet directamente)
    st.subheader("1. Base de Datos")
    archivo_catalogo = st.file_uploader("Sube el archivo maestro (Auditoria_Historico.xlsx) para leer el catálogo:", type=["xls", "xlsx"])

    if archivo_catalogo:
        try:
            # Lee estrictamente la pestaña que creamos
            df_cat = pd.read_excel(archivo_catalogo, sheet_name="Catalogo_Materiales", dtype=str)
            
            st.markdown("---")
            st.subheader("2. Selección de Productos")
            
            # Diseño en 3 columnas para los filtros
            col1, col2, col3 = st.columns(3)
            
            with col1:
                categorias = df_cat['Categoria'].dropna().unique()
                cat_seleccionada = st.selectbox("Categoría:", options=categorias)
            
            # Filtramos el DataFrame por la categoría seleccionada
            df_filtrado_1 = df_cat[df_cat['Categoria'] == cat_seleccionada]
            
            with col2:
                subcategorias = df_filtrado_1['Sub_Categoria'].dropna().unique()
                subcat_seleccionada = st.selectbox("Sub Categoría:", options=subcategorias)
                
            # Filtramos el DataFrame por la subcategoría seleccionada
            df_filtrado_2 = df_filtrado_1[df_filtrado_1['Sub_Categoria'] == subcat_seleccionada]
            
            with col3:
                productos = df_filtrado_2['Nombre_Amigable'].dropna().unique()
                prod_seleccionado = st.selectbox("Producto:", options=productos)

            # Obtenemos la fila exacta del producto final seleccionado
            producto_final = df_filtrado_2[df_filtrado_2['Nombre_Amigable'] == prod_seleccionado].iloc[0]
            
            # 3. SELECCIÓN DE CANTIDAD Y BOTÓN DE AGREGAR
            st.markdown("###")
            col_cant, col_btn = st.columns([1, 2])
            
            with col_cant:
                cantidad = st.number_input(f"Cantidad ({producto_final['Unidad_Medida']}):", min_value=1, value=1, step=1)
                
            with col_btn:
                st.markdown("<br>", unsafe_allow_html=True) # Espacio para alinear con el input
                if st.button("➕ Agregar al Pedido", type="primary", use_container_width=True):
                    # Crear nueva fila con el formato de Nexus
                    nueva_fila = pd.DataFrame({
                        "Codigo_Nexus": [producto_final['Codigo_Nexus']],
                        "Descripcion_Nexus": [producto_final['Descripcion_Nexus']],
                        "Unidad_Medida": [producto_final['Unidad_Medida']],
                        "Cantidad": [cantidad]
                    })
                    
                    # Añadir al carrito y consolidar automáticamente
                    carrito_actual = st.session_state['carrito_pedidos']
                    carrito_actual = pd.concat([carrito_actual, nueva_fila], ignore_index=True)
                    st.session_state['carrito_pedidos'] = consolidar_carrito(carrito_actual)
                    
                    st.success(f"✅ Se agregaron {cantidad}x {prod_seleccionado}.")

            # 4. VISUALIZACIÓN DEL CARRITO Y DESCARGA
            if not st.session_state['carrito_pedidos'].empty:
                st.markdown("---")
                st.subheader("🛒 Pedido Actual")
                
                # Mostrar tabla con estilo
                st.dataframe(st.session_state['carrito_pedidos'], use_container_width=True, hide_index=True)
                
                # Botones de acción final
                col_dl, col_del = st.columns([2, 1])
                
                with col_dl:
                    # Generar Excel en memoria
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        st.session_state['carrito_pedidos'].to_excel(writer, index=False, sheet_name='Traslado_Nexus')
                    datos_excel = output.getvalue()
                    
                    st.download_button(
                        label="⬇️ Descargar Archivo para Nexus",
                        data=datos_excel,
                        file_name="Pedido_Bodega_Cafeteria.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                with col_del:
                    if st.button("🗑️ Vaciar Pedido", use_container_width=True):
                        st.session_state['carrito_pedidos'] = pd.DataFrame(columns=["Codigo_Nexus", "Descripcion_Nexus", "Unidad_Medida", "Cantidad"])
                        st.rerun()

        except Exception as e:
            st.error(f"Error al leer el catálogo. Verifica que la pestaña se llame 'Catalogo_Materiales' y tenga las columnas correctas. Detalle técnico: {e}")