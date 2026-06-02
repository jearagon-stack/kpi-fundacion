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
    st.title("🛒 Pedidos de Materiales - Cafetería")
    st.info("Selecciona la categoría e ingresa las cantidades de los productos que necesites pedir.")

    if 'carrito_pedidos' not in st.session_state:
        st.session_state['carrito_pedidos'] = pd.DataFrame(columns=["IdProducto", "Descripcion", "MEDIDA", "Cantidad"])

    # ========================================================
    # 1. CARGA AUTOMÁTICA USANDO TU FUNCIÓN INTERNA
    # ========================================================
    try:
        df_cat = obtener_dataframe("Catalogo_Materiales")
        
        # Forzamos los nombres de las primeras 6 columnas
        nombres_correctos = ["Categoria", "SubCategoria", "Nombre_Amigable", "IdProducto", "Descripcion", "MEDIDA"]
        df_cat.columns = nombres_correctos + list(df_cat.columns[6:])
        
        st.subheader("1. Buscar Productos")
        
        col1, col2 = st.columns(2)
        
        with col1:
            categorias = df_cat['Categoria'].dropna().unique()
            cat_seleccionada = st.selectbox("Categoría:", options=categorias)
        
        df_filtrado_1 = df_cat[df_cat['Categoria'] == cat_seleccionada]
        
        with col2:
            subcategorias = df_filtrado_1['SubCategoria'].dropna().unique()
            subcat_seleccionada = st.selectbox("Sub Categoría:", options=subcategorias)
            
        # Filtramos los productos que pertenecen a la selección
        df_filtrado_2 = df_filtrado_1[df_filtrado_1['SubCategoria'] == subcat_seleccionada]
        
        # ========================================================
        # 2. LISTA DE PRODUCTOS (TIPO POS) Y BOTÓN DE AGREGAR
        # ========================================================
        st.markdown("---")
        st.subheader(f"📦 Catálogo: {subcat_seleccionada}")
        
        # Diccionario para guardar lo que digita la cajera
        cantidades_ingresadas = {}
        
        # Creamos encabezados visuales
        encabezado_prod, encabezado_cant = st.columns([3, 1])
        with encabezado_prod:
            st.markdown("**Nombre del Producto**")
        with encabezado_cant:
            st.markdown("**Cantidad**")
            
        st.divider()

        # Dibujamos una fila por cada producto de la subcategoría
        for idx, row in df_filtrado_2.iterrows():
            c_prod, c_cant = st.columns([3, 1])
            
            with c_prod:
                st.markdown(f"{row['Nombre_Amigable']} *(Medida: {row['MEDIDA']})*")
                
            with c_cant:
                cantidades_ingresadas[idx] = st.number_input(
                    "Cantidad",
                    min_value=0,
                    value=0,
                    step=1,
                    key=f"prod_{row['IdProducto']}_{idx}",
                    label_visibility="collapsed"
                )

        st.markdown("###")
        if st.button("➕ Agregar Seleccionados al Pedido", type="primary", use_container_width=True):
            nuevas_filas = []
            
            # Revisamos cuáles productos tuvieron una cantidad mayor a cero
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
                
                st.success(f"✅ Se agregaron {len(nuevas_filas)} productos al carrito de pedidos.")
            else:
                st.warning("⚠️ No ingresaste ninguna cantidad mayor a 0.")

        # ========================================================
        # 3. VISUALIZACIÓN DEL CARRITO Y DESCARGA
        # ========================================================
        if not st.session_state['carrito_pedidos'].empty:
            st.markdown("---")
            st.subheader("🛒 Resumen del Pedido Actual")
            
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
                    st.session_state['carrito_pedidos'] = pd.DataFrame(columns=["IdProducto", "Descripcion", "MEDIDA", "Cantidad"])
                    st.rerun()

    except Exception as e:
        st.error(f"Error técnico inesperado. Detalle: {e}")