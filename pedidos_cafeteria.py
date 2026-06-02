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
    df_agrupado = df_carrito.groupby(["Codigo_Nexus", "Descripcion_Nexus", "Unidad_Medida"], as_index=False)["Cantidad"].sum()
    return df_agrupado

def mostrar_modulo_pedidos():
    st.title("🛒 Pedidos de Materiales - Cafetería")
    st.info("Selecciona los productos para generar la solicitud de traslado de bodega.")

    if 'carrito_pedidos' not in st.session_state:
        st.session_state['carrito_pedidos'] = pd.DataFrame(columns=["Codigo_Nexus", "Descripcion_Nexus", "Unidad_Medida", "Cantidad"])

    # ========================================================
    # 1. CARGA AUTOMÁTICA USANDO TU FUNCIÓN INTERNA
    # ========================================================
    try:
        # Usamos tu función para leer directamente la pestaña
        df_cat = obtener_dataframe("Catalogo_Materiales")
        
        # Limpiamos los nombres de las columnas por si tienen espacios invisibles
        df_cat.columns = df_cat.columns.str.strip()
        
        st.subheader("1. Selección de Productos")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            categorias = df_cat['Categoria'].dropna().unique()
            cat_seleccionada = st.selectbox("Categoría:", options=categorias)
        
        df_filtrado_1 = df_cat[df_cat['Categoria'] == cat_seleccionada]
        
        with col2:
            # CORREGIDO: Buscamos "Subcategoria" sin guion bajo
            subcategorias = df_filtrado_1['Subcategoria'].dropna().unique()
            subcat_seleccionada = st.selectbox("Sub Categoría:", options=subcategorias)
            
        df_filtrado_2 = df_filtrado_1[df_filtrado_1['Subcategoria'] == subcat_seleccionada]
        
        with col3:
            productos = df_filtrado_2['Nombre_Amigable'].dropna().unique()
            prod_seleccionado = st.selectbox("Producto:", options=productos)

        producto_final = df_filtrado_2[df_filtrado_2['Nombre_Amigable'] == prod_seleccionado].iloc[0]
        
        # ========================================================
        # 2. SELECCIÓN DE CANTIDAD Y BOTÓN DE AGREGAR
        # ========================================================
        st.markdown("###")
        col_cant, col_btn = st.columns([1, 2])
        
        with col_cant:
            cantidad = st.number_input(f"Cantidad ({producto_final['Unidad_Medida']}):", min_value=1, value=1, step=1)
            
        with col_btn:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("➕ Agregar al Pedido", type="primary", use_container_width=True):
                nueva_fila = pd.DataFrame({
                    "Codigo_Nexus": [producto_final['Codigo_Nexus']],
                    "Descripcion_Nexus": [producto_final['Descripcion_Nexus']],
                    "Unidad_Medida": [producto_final['Unidad_Medida']],
                    "Cantidad": [cantidad]
                })
                
                carrito_actual = st.session_state['carrito_pedidos']
                carrito_actual = pd.concat([carrito_actual, nueva_fila], ignore_index=True)
                st.session_state['carrito_pedidos'] = consolidar_carrito(carrito_actual)
                
                st.success(f"✅ Se agregaron {cantidad}x {prod_seleccionado}.")

        # ========================================================
        # 3. VISUALIZACIÓN DEL CARRITO Y DESCARGA
        # ========================================================
        if not st.session_state['carrito_pedidos'].empty:
            st.markdown("---")
            st.subheader("🛒 Pedido Actual")
            
            st.dataframe(st.session_state['carrito_pedidos'], use_container_width=True, hide_index=True)
            
            col_dl, col_del = st.columns([2, 1])
            
            with col_dl:
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
        st.error(f"Error de conexión con la base de datos. Verifica que la pestaña se llame exactamente 'Catalogo_Materiales' en tu Google Sheets. Detalle: {e}")