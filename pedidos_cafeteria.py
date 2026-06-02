import streamlit as st
import pandas as pd
import io

try:
    from utils import obtener_dataframe
except ImportError:
    st.error("⚠️ Error: No se encontró el módulo 'utils' para conectar con la base de datos.")

def consolidar_carrito(df_carrito):
    if df_carrito.empty:
        return df_carrito
    df_agrupado = df_carrito.groupby(["IdProducto", "Descripcion", "MEDIDA"], as_index=False)["Cantidad"].sum()
    return df_agrupado

# ========================================================
# FUNCIONES INFALIBLES (CALLBACKS) PARA BOTONES
# ========================================================
def enviar_pedido_bodega(anexo_usuario):
    # Generamos el ID
    num_correlativo = st.session_state['correlativo_pedido']
    id_pedido = f"Pedido {num_correlativo:04d}"
    
    # Preparamos los datos
    df_nuevo = st.session_state['carrito_pedidos'].copy()
    df_nuevo['ID_Pedido'] = id_pedido
    df_nuevo['Anexo'] = anexo_usuario
    
    # Guardamos en la bandeja del bodeguero
    if st.session_state['pedidos_enviados'].empty:
        st.session_state['pedidos_enviados'] = df_nuevo
    else:
        st.session_state['pedidos_enviados'] = pd.concat([st.session_state['pedidos_enviados'], df_nuevo], ignore_index=True)
    
    # Limpiamos el carrito y subimos el número de pedido
    st.session_state['correlativo_pedido'] += 1
    st.session_state['carrito_pedidos'] = pd.DataFrame(columns=["IdProducto", "Descripcion", "MEDIDA", "Cantidad"])
    st.session_state['msg_exito_pedido'] = f"✅ El {id_pedido} fue enviado exitosamente a la bodega."

def vaciar_carrito():
    st.session_state['carrito_pedidos'] = pd.DataFrame(columns=["IdProducto", "Descripcion", "MEDIDA", "Cantidad"])

def marcar_procesado(id_pedido):
    # Borramos solo el pedido seleccionado de la bandeja
    st.session_state['pedidos_enviados'] = st.session_state['pedidos_enviados'][st.session_state['pedidos_enviados']['ID_Pedido'] != id_pedido]
# ========================================================

def mostrar_modulo_pedidos():
    st.title("🛒 Gestión de Pedidos - Cafetería")
    
    anexo_usuario = st.session_state.get('anexo_actual', 'Desconocido')
    rol_usuario = st.session_state.get('rol_actual', 'CAJERA')

    # Inicialización de memorias
    if 'carrito_pedidos' not in st.session_state:
        st.session_state['carrito_pedidos'] = pd.DataFrame(columns=["IdProducto", "Descripcion", "MEDIDA", "Cantidad"])
    
    if 'pedidos_enviados' not in st.session_state:
        st.session_state['pedidos_enviados'] = pd.DataFrame(columns=["ID_Pedido", "Anexo", "IdProducto", "Descripcion", "MEDIDA", "Cantidad"])
        
    if 'correlativo_pedido' not in st.session_state:
        st.session_state['correlativo_pedido'] = 1

    # Permisos de Pestañas
    if rol_usuario in ["ADMIN", "BODEGUERO"]:
        tabs = st.tabs(["🛍️ 1. Crear Pedido (Cajas)", "📦 2. Gestión de Bodega"])
        tab_cajas = tabs[0]
        tab_bodega = tabs[1]
    else:
        tabs = st.tabs(["🛍️ 1. Crear Pedido (Cajas)"])
        tab_cajas = tabs[0]
        tab_bodega = None

    # --------------------------------------------------------
    # PESTAÑA 1: VISTA DE CAJERAS
    # --------------------------------------------------------
    with tab_cajas:
        # Mostrar mensaje de éxito si existe
        if 'msg_exito_pedido' in st.session_state:
            st.success(st.session_state['msg_exito_pedido'])
            del st.session_state['msg_exito_pedido'] # Lo borramos para que no salga siempre
            
        st.info(f"📍 Creando pedido para: **{anexo_usuario}**")
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

            # Mostramos los productos
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

            # Si hay algo en la lista, mostramos opciones de envío
            if not st.session_state['carrito_pedidos'].empty:
                st.markdown("---")
                st.subheader("🛒 Lista Preliminar del Pedido")
                st.dataframe(st.session_state['carrito_pedidos'], use_container_width=True, hide_index=True)
                
                st.warning("⚠️ Revisa tu pedido antes de enviarlo. No podrás hacer cambios desde esta pantalla después.")
                
                col_env, col_can = st.columns([2, 1])
                with col_env:
                    # Usamos 'on_click' para garantizar que la función de guardado se ejecute
                    st.button("🚀 Enviar Pedido a Bodega", type="primary", use_container_width=True, on_click=enviar_pedido_bodega, args=(anexo_usuario,))
                with col_can:
                    st.button("🗑️ Vaciar Lista", use_container_width=True, on_click=vaciar_carrito)

        except Exception as e:
            st.error(f"Error técnico. Detalle: {e}")

    # --------------------------------------------------------
    # PESTAÑA 2: VISTA DE BODEGUERO
    # --------------------------------------------------------
    if tab_bodega is not None:
        with tab_bodega:
            st.subheader("📋 Pedidos Pendientes de Procesar")
            
            if st.session_state['pedidos_enviados'].empty:
                st.info("No hay pedidos entrantes en este momento.")
            else:
                lista_pedidos = st.session_state['pedidos_enviados']['ID_Pedido'].unique()
                pedido_seleccionado = st.selectbox("Seleccionar Pedido a Revisar:", lista_pedidos)
                
                df_pedido_actual = st.session_state['pedidos_enviados'][st.session_state['pedidos_enviados']['ID_Pedido'] == pedido_seleccionado].copy()
                anexo_del_pedido = df_pedido_actual['Anexo'].iloc[0]
                
                st.markdown(f"**Origen:** {anexo_del_pedido}")
                st.markdown("Revisa el pedido. Puedes hacer doble clic en la tabla para **editar cantidades** o agregar filas nuevas si es necesario.")
                
                df_editor = df_pedido_actual.drop(columns=['ID_Pedido', 'Anexo'])
                
                pedido_editado = st.data_editor(
                    df_editor,
                    num_rows="dynamic",
                    use_container_width=True,
                    key=f"editor_{pedido_seleccionado}"
                )
                
                st.markdown("---")
                col_dl, col_del = st.columns([2, 1])
                
                nombre_archivo = f"{pedido_seleccionado} {anexo_del_pedido}.xlsx"
                
                with col_dl:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        pedido_editado.to_excel(writer, index=False, sheet_name='Traslado_Nexus')
                    datos_excel = output.getvalue()
                    
                    st.download_button(
                        label=f"⬇️ Aprobar y Descargar ({nombre_archivo})",
                        data=datos_excel,
                        file_name=nombre_archivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                with col_del:
                    st.button("✔️ Marcar como Procesado", use_container_width=True, on_click=marcar_procesado, args=(pedido_seleccionado,))