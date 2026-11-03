import io

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import validacion_cxp as cxp  # Reutiliza lectores y funciones auxiliares

# ==========================================
# CONFIGURACIÓN
# ==========================================
UMBRAL_ALERTA = 10.0        # % de variación de precio que dispara una alerta
FACTOR_CAMBIO_UNIDAD = 2.5  # Si el precio se multiplica o divide por esto, probablemente cambió la unidad de medida
AGRUPACIONES = ["Día", "Semana", "Mes", "Año"]

CATEGORIAS_GASTO = [
    "Combustible", "Mantenimiento", "Arrendamiento", "Cuotas de activos", "Control de plagas",
    "Liquidación de personal", "Prestaciones al personal", "Transporte de personal",
    "Utensilios y menaje", "Insumos menores", "Máquinas vending", "Otros",
]
SIN_CLASIFICAR = "Sin clasificar"

# Palabras clave (sin tildes ni espacios). Se revisan en este orden: la primera que coincide gana.
REGLAS_GASTO = [
    ("Liquidación de personal", ["LIQUIDAC", "INDEMNIZ", "FINIQUITO"]),
    ("Prestaciones al personal", ["PRESTACION", "REINTEGRO"]),
    ("Transporte de personal", ["UBER", "INDRIVE", "TAXI", "TRANSPORTE"]),
    ("Control de plagas", ["PLAGA", "FUMIGA", "DESINFECT", "DESIFECT"]),
    ("Combustible", ["DIESEL", "DISEL", "GASOLINA", "FOVIAL", "COTRANS", "CONTRANS", "COMBUSTIBLE"]),
    ("Cuotas de activos", ["CUOTA"]),
    ("Arrendamiento", ["ARRENDAMIENTO", "ALQUILER"]),
    ("Máquinas vending", ["VENDING"]),
    ("Mantenimiento", ["MANTENIMIENTO", "REPARACION", "TRAMPADEGRASA"]),
    ("Utensilios y menaje", ["ESPATULA", "GUANTE", "FRASCO", "TAZA", "EXTENSION", "FOCO", "SCOTCH",
                             "CINTA", "UTENSILIO", "SARTEN", "OLLA", "CUCHILLO"]),
]


def clasificar_gasto(descripcion):
    d = cxp._norm(descripcion)
    for categoria, palabras in REGLAS_GASTO:
        if any(p in d for p in palabras):
            return categoria
    return None


def _pct(nuevo, anterior):
    if anterior is None or pd.isna(anterior) or anterior == 0 or pd.isna(nuevo):
        return None
    return (nuevo - anterior) / abs(anterior) * 100


# ==========================================
# LECTURA DEL HISTORIAL
# ==========================================
@st.cache_data(show_spinner=False)
def _leer_uno(contenido, nombre):
    archivo = io.BytesIO(contenido)
    archivo.name = nombre
    df = cxp._leer_archivo(archivo, ["NUMERO", "TOTAL"])
    if df is None:
        raise ValueError(f"No encontré los encabezados 'Numero' y 'Total' en '{nombre}'.")
    col = lambda ops: cxp._buscar_col(df, ops)
    req = {"Numero": col(["NUMERO"]), "Fecha": col(["FECHA"]), "Nombre": col(["NOMBRE", "PROVEEDOR"]),
           "IdProducto": col(["IDPRODUCTO", "CODIGOPRODUCTO"]), "Cantidad": col(["CANTIDAD"]),
           "Descripcion": col(["DESCRIPCION"]), "Total": col(["TOTAL"])}
    faltan = [k for k, v in req.items() if v is None]
    if faltan:
        raise ValueError(f"Al archivo '{nombre}' le faltan las columnas: {', '.join(faltan)}.")
    c_cat, c_tipo = col(["CATEGORIA"]), col(["TIPODOCUMENTO"])

    tipo = df[c_tipo].astype(str).str.strip() if c_tipo else pd.Series("", index=df.index)
    es_nc = tipo.apply(cxp._norm).str.contains("NOTADECREDITO")
    signo = es_nc.map({True: -1.0, False: 1.0})  # Notas de crédito restan
    return pd.DataFrame({
        "ARCHIVO": nombre,
        "FECHA": cxp._a_fecha(df[req["Fecha"]]),
        "NUMERO": df[req["Numero"]].astype(str).str.strip(),
        "PROVEEDOR": df[req["Nombre"]].astype(str).str.strip(),
        "CODIGO": df[req["IdProducto"]].astype(str).str.strip().str.upper(),
        "DESCRIPCION": df[req["Descripcion"]].astype(str).str.strip(),
        "CATEGORIA": df[c_cat].astype(str).str.strip().str.upper() if c_cat else "SIN CATEGORÍA",
        "TIPO DOCUMENTO": tipo,
        "ES NOTA CREDITO": es_nc,
        "CANTIDAD": cxp._a_numero(df[req["Cantidad"]]).abs() * signo,
        "MONTO": cxp._a_numero(df[req["Total"]]).abs() * signo,
    })


def leer_historial(archivos):
    partes = []
    for i, a in enumerate(archivos):
        parte = _leer_uno(a.getvalue(), a.name).copy()
        parte["_ARCH"] = i  # identifica cada archivo subido (aunque dos tengan el mismo nombre)
        partes.append(parte)
    df = pd.concat(partes, ignore_index=True)
    # Si dos archivos cubren las mismas fechas, se quitan las líneas repetidas entre archivos
    cols = ["FECHA", "NUMERO", "CODIGO", "DESCRIPCION", "CANTIDAD", "MONTO"]
    df["_n"] = df.groupby(["_ARCH"] + cols, dropna=False).cumcount()
    antes = len(df)
    df = df.drop_duplicates(subset=cols + ["_n"]).drop(columns=["_n", "_ARCH"])
    repetidas = antes - len(df)

    df = df.dropna(subset=["FECHA"]).copy()
    df["ES GASTO"] = df["CODIGO"].str.startswith("G")
    ultima_desc = df.sort_values("FECHA").groupby("CODIGO")["DESCRIPCION"].last()
    df["PRODUCTO"] = df["CODIGO"] + " - " + df["CODIGO"].map(ultima_desc)
    return df.reset_index(drop=True), repetidas


def agregar_periodo(df, agrupacion):
    df = df.copy()
    f = df["FECHA"].dt.normalize()
    if agrupacion == "Día":
        inicio = f
        etiqueta = inicio.dt.strftime("%d/%m/%Y")
    elif agrupacion == "Semana":
        inicio = f - pd.to_timedelta(f.dt.weekday, unit="D")
        etiqueta = "Sem. " + inicio.dt.strftime("%d/%m/%Y") + " al " + (inicio + pd.Timedelta(days=6)).dt.strftime("%d/%m/%Y")
    elif agrupacion == "Mes":
        inicio = f.dt.to_period("M").dt.to_timestamp()
        etiqueta = inicio.apply(cxp._etiqueta_mes)
    else:
        inicio = f.dt.to_period("Y").dt.to_timestamp()
        etiqueta = inicio.dt.year.astype(str)
    df["PER INICIO"] = inicio
    df["PERIODO"] = etiqueta
    return df


def lista_periodos(df):
    return df.drop_duplicates("PERIODO").sort_values("PER INICIO")["PERIODO"].tolist()


# ==========================================
# CÁLCULOS
# ==========================================
def comparar_periodos(df, clave, per_a, per_b, con_cantidad=True):
    """Compara dos períodos: cantidad, precio promedio y monto, con efecto precio y efecto cantidad."""
    def _agr(per):
        sub = df[df["PERIODO"] == per]
        return sub.groupby(clave).agg(CANT=("CANTIDAD", "sum"), MONTO=("MONTO", "sum"), DOCS=("NUMERO", "nunique"))

    a, b = _agr(per_a), _agr(per_b)
    t = a.join(b, how="outer", lsuffix="_A", rsuffix="_B").fillna(0).reset_index()
    out = pd.DataFrame({clave: t[clave]})

    if con_cantidad:
        out["CANTIDAD A"], out["CANTIDAD B"] = t["CANT_A"], t["CANT_B"]
        out["VAR CANTIDAD"] = t["CANT_B"] - t["CANT_A"]
        out["VAR CANTIDAD %"] = [_pct(n, v) for n, v in zip(t["CANT_B"], t["CANT_A"])]
        pa = (t["MONTO_A"] / t["CANT_A"]).where(t["CANT_A"] > 0)
        pb = (t["MONTO_B"] / t["CANT_B"]).where(t["CANT_B"] > 0)
        out["PRECIO PROM A"], out["PRECIO PROM B"] = pa, pb
        out["VAR PRECIO %"] = [_pct(n, v) for n, v in zip(pb, pa)]

    out["MONTO A"], out["MONTO B"] = t["MONTO_A"], t["MONTO_B"]
    out["VAR MONTO $"] = t["MONTO_B"] - t["MONTO_A"]
    out["VAR MONTO %"] = [_pct(n, v) for n, v in zip(t["MONTO_B"], t["MONTO_A"])]

    if con_cantidad:
        # Efecto precio = cambio de precio × cantidad nueva; el resto del cambio en monto es efecto cantidad
        ambos = (t["CANT_A"] > 0) & (t["CANT_B"] > 0)
        out["EFECTO PRECIO $"] = ((pb - pa) * t["CANT_B"]).where(ambos, 0.0)
        out["EFECTO CANTIDAD $"] = out["VAR MONTO $"] - out["EFECTO PRECIO $"]

    out["DOCS A"], out["DOCS B"] = t["DOCS_A"].astype(int), t["DOCS_B"].astype(int)
    out["_ABS"] = out["VAR MONTO $"].abs()
    return out.sort_values("_ABS", ascending=False).drop(columns="_ABS").reset_index(drop=True)


def compras_por_documento(df):
    """Una fila por compra de cada producto, con su variación contra la compra anterior y contra el promedio."""
    base = df[~df["ES NOTA CREDITO"] & ~df["ES GASTO"] & (df["CANTIDAD"] > 0)]
    c = base.groupby(["PRODUCTO", "FECHA", "NUMERO", "PROVEEDOR", "CATEGORIA"], as_index=False).agg(
        CANTIDAD=("CANTIDAD", "sum"), MONTO=("MONTO", "sum"))
    c = c.sort_values(["PRODUCTO", "FECHA", "NUMERO"]).reset_index(drop=True)
    c["PRECIO UNITARIO"] = c["MONTO"] / c["CANTIDAD"]

    g = c.groupby("PRODUCTO")
    c["FECHA ANTERIOR"] = g["FECHA"].shift()
    c["PRECIO ANTERIOR"] = g["PRECIO UNITARIO"].shift()
    c["CANTIDAD ANTERIOR"] = g["CANTIDAD"].shift()
    c["MONTO ANTERIOR"] = g["MONTO"].shift()
    c["DIAS DESDE ANTERIOR"] = (c["FECHA"] - c["FECHA ANTERIOR"]).dt.days
    c["VAR PRECIO vs ANTERIOR %"] = [_pct(n, v) for n, v in zip(c["PRECIO UNITARIO"], c["PRECIO ANTERIOR"])]
    c["VAR CANTIDAD vs ANTERIOR"] = c["CANTIDAD"] - c["CANTIDAD ANTERIOR"]
    c["VAR MONTO vs ANTERIOR $"] = c["MONTO"] - c["MONTO ANTERIOR"]
    promedio = g["MONTO"].transform("sum") / g["CANTIDAD"].transform("sum")
    c["PRECIO PROMEDIO"] = promedio
    c["VAR PRECIO vs PROMEDIO %"] = [_pct(n, v) for n, v in zip(c["PRECIO UNITARIO"], promedio)]
    return c


def marcar_alertas(c, umbral):
    def _tipo(r):
        ratio_ant = r["PRECIO UNITARIO"] / r["PRECIO ANTERIOR"] if pd.notna(r["PRECIO ANTERIOR"]) and r["PRECIO ANTERIOR"] else None
        if ratio_ant and (ratio_ant >= FACTOR_CAMBIO_UNIDAD or ratio_ant <= 1 / FACTOR_CAMBIO_UNIDAD):
            return "⚠️ Posible cambio de unidad"
        va, vp = r["VAR PRECIO vs ANTERIOR %"], r["VAR PRECIO vs PROMEDIO %"]
        fuera = [v for v in (va, vp) if v is not None and pd.notna(v) and abs(v) >= umbral]
        if not fuera:
            return None
        return "⬆️ Subió de precio" if fuera[0] > 0 else "⬇️ Bajó de precio"

    c = c.copy()
    c["ALERTA"] = c.apply(_tipo, axis=1)
    return c


# ==========================================
# PANTALLA
# ==========================================
def _cfg(dinero=(), precio=(), pct=(), num=()):
    cfg = {c: st.column_config.NumberColumn(c, format="$%.2f") for c in dinero}
    cfg.update({c: st.column_config.NumberColumn(c, format="$%.4f") for c in precio})
    cfg.update({c: st.column_config.NumberColumn(c, format="%.1f%%") for c in pct})
    cfg.update({c: st.column_config.NumberColumn(c, format="%.2f") for c in num})
    return cfg


CFG_COMPARA = dict(
    dinero=["MONTO A", "MONTO B", "VAR MONTO $", "EFECTO PRECIO $", "EFECTO CANTIDAD $"],
    precio=["PRECIO PROM A", "PRECIO PROM B"],
    pct=["VAR CANTIDAD %", "VAR PRECIO %", "VAR MONTO %"],
    num=["CANTIDAD A", "CANTIDAD B", "VAR CANTIDAD"],
)


def _elegir_periodos(df, prefijo):
    periodos = lista_periodos(df)
    if len(periodos) < 2:
        st.info("Con los filtros actuales solo hay un período. Amplía el rango de fechas, cambia la agrupación o sube más archivos.")
        return None, None
    c1, c2 = st.columns(2)
    sufijo = abs(hash(tuple(periodos)))  # si cambia la lista de períodos, la selección se reinicia sin error
    per_a = c1.selectbox("Período A (base):", periodos, index=len(periodos) - 2, key=f"{prefijo}_a_{sufijo}")
    per_b = c2.selectbox("Período B (a comparar):", periodos, index=len(periodos) - 1, key=f"{prefijo}_b_{sufijo}")
    return per_a, per_b


def _vista_periodos(prod):
    st.markdown("**Compara cantidad, precio promedio y monto de cada producto entre dos períodos.**")
    per_a, per_b = _elegir_periodos(prod, "mc_pp")
    if not per_a:
        return
    comp = comparar_periodos(prod, "PRODUCTO", per_a, per_b)

    tot_a, tot_b = comp["MONTO A"].sum(), comp["MONTO B"].sum()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric(f"Monto {per_a}", f"${tot_a:,.2f}")
    m2.metric(f"Monto {per_b}", f"${tot_b:,.2f}", f"{(tot_b - tot_a):+,.2f}", delta_color="inverse")
    m3.metric("Explicado por precio", f"${comp['EFECTO PRECIO $'].sum():+,.2f}")
    m4.metric("Explicado por cantidad", f"${comp['EFECTO CANTIDAD $'].sum():+,.2f}")

    top = comp.head(15).melt(id_vars="PRODUCTO", value_vars=["EFECTO PRECIO $", "EFECTO CANTIDAD $"],
                             var_name="EFECTO", value_name="MONTO")
    if not top.empty:
        fig = px.bar(top, y="PRODUCTO", x="MONTO", color="EFECTO", orientation="h", barmode="relative",
                     title="Los 15 productos que más cambiaron en monto: ¿por precio o por cantidad?",
                     labels={"MONTO": "Cambio en monto ($)", "PRODUCTO": ""})
        fig.update_layout(height=520, yaxis={"autorange": "reversed"})
        st.plotly_chart(fig, use_container_width=True)

    st.dataframe(comp, use_container_width=True, hide_index=True, column_config=_cfg(**CFG_COMPARA))
    st.download_button("📥 Descargar comparativo (Excel)", data=cxp._generar_excel({"Periodo vs periodo": comp}),
                       file_name="Comparativo_periodos.xlsx", key="mc_dl_pp")


def _vista_ficha(prod, umbral):
    compras = marcar_alertas(compras_por_documento(prod), umbral)
    if compras.empty:
        st.info("No hay compras de productos con los filtros actuales.")
        return
    orden = compras.groupby("PRODUCTO")["MONTO"].sum().sort_values(ascending=False).index.tolist()
    producto = st.selectbox("Producto (ordenados de mayor a menor gasto):", orden,
                            key=f"mc_producto_{abs(hash(tuple(orden)))}")
    c = compras[compras["PRODUCTO"] == producto].reset_index(drop=True)

    prom = c["MONTO"].sum() / c["CANTIDAD"].sum()
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Compras", len(c))
    m2.metric("Cantidad total", f"{c['CANTIDAD'].sum():,.2f}")
    m3.metric("Monto total", f"${c['MONTO'].sum():,.2f}")
    m4.metric("Precio promedio", f"${prom:,.4f}")
    m5.metric("Precio mín / máx", f"${c['PRECIO UNITARIO'].min():,.2f} / ${c['PRECIO UNITARIO'].max():,.2f}")

    fig = go.Figure()
    fig.add_bar(x=c["FECHA"], y=c["CANTIDAD"], name="Cantidad", yaxis="y2", opacity=0.45)
    fig.add_scatter(x=c["FECHA"], y=c["PRECIO UNITARIO"], name="Precio unitario", mode="lines+markers")
    fig.add_hline(y=prom, line_dash="dot", annotation_text="Promedio")
    fig.update_layout(title=f"Precio y cantidad de cada compra — {producto}", height=420,
                      yaxis=dict(title="Precio unitario ($)"),
                      yaxis2=dict(title="Cantidad", overlaying="y", side="right", showgrid=False),
                      legend=dict(orientation="h", y=-0.2))
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("**🔀 Comparador libre: elige dos compras cualesquiera**")
    etiquetas = [f"{r.FECHA.strftime('%d/%m/%Y')} | {r.PROVEEDOR[:30]} | {r.CANTIDAD:,.2f} a ${r['PRECIO UNITARIO']:,.4f}"
                 for _, r in c.iterrows()]
    k1, k2 = st.columns(2)
    ia = k1.selectbox("Compra A:", range(len(c)), index=max(len(c) - 2, 0), format_func=lambda i: etiquetas[i], key=f"mc_ca_{producto}_{len(c)}")
    ib = k2.selectbox("Compra B:", range(len(c)), index=len(c) - 1, format_func=lambda i: etiquetas[i], key=f"mc_cb_{producto}_{len(c)}")
    a, b = c.loc[ia], c.loc[ib]
    lado = pd.DataFrame({
        "CONCEPTO": ["Fecha", "Proveedor", "Cantidad", "Precio unitario", "Monto"],
        "COMPRA A": [a["FECHA"].strftime("%d/%m/%Y"), a["PROVEEDOR"], f"{a['CANTIDAD']:,.2f}", f"${a['PRECIO UNITARIO']:,.4f}", f"${a['MONTO']:,.2f}"],
        "COMPRA B": [b["FECHA"].strftime("%d/%m/%Y"), b["PROVEEDOR"], f"{b['CANTIDAD']:,.2f}", f"${b['PRECIO UNITARIO']:,.4f}", f"${b['MONTO']:,.2f}"],
        "DIFERENCIA": [f"{(b['FECHA'] - a['FECHA']).days} días", "",
                       f"{b['CANTIDAD'] - a['CANTIDAD']:+,.2f} ({(_pct(b['CANTIDAD'], a['CANTIDAD']) or 0):+.1f}%)",
                       f"{b['PRECIO UNITARIO'] - a['PRECIO UNITARIO']:+,.4f} ({(_pct(b['PRECIO UNITARIO'], a['PRECIO UNITARIO']) or 0):+.1f}%)",
                       f"{b['MONTO'] - a['MONTO']:+,.2f} ({(_pct(b['MONTO'], a['MONTO']) or 0):+.1f}%)"],
    })
    st.dataframe(lado, use_container_width=True, hide_index=True)
    efecto_precio = (b["PRECIO UNITARIO"] - a["PRECIO UNITARIO"]) * b["CANTIDAD"]
    st.caption(f"De la diferencia en monto (${b['MONTO'] - a['MONTO']:+,.2f}): "
               f"${efecto_precio:+,.2f} por precio y ${(b['MONTO'] - a['MONTO']) - efecto_precio:+,.2f} por cantidad.")

    st.markdown("**Historial de compras del producto**")
    det = c[["FECHA", "PROVEEDOR", "NUMERO", "CANTIDAD", "PRECIO UNITARIO", "MONTO", "DIAS DESDE ANTERIOR",
             "VAR PRECIO vs ANTERIOR %", "VAR CANTIDAD vs ANTERIOR", "VAR MONTO vs ANTERIOR $",
             "VAR PRECIO vs PROMEDIO %", "ALERTA"]].sort_values("FECHA", ascending=False).copy()
    det["FECHA"] = cxp._fmt_fecha(det["FECHA"])
    st.dataframe(det, use_container_width=True, hide_index=True,
                 column_config=_cfg(dinero=["MONTO", "VAR MONTO vs ANTERIOR $"], precio=["PRECIO UNITARIO"],
                                    pct=["VAR PRECIO vs ANTERIOR %", "VAR PRECIO vs PROMEDIO %"],
                                    num=["CANTIDAD", "VAR CANTIDAD vs ANTERIOR"]))
    st.download_button("📥 Descargar historial del producto (Excel)", data=cxp._generar_excel({"Historial": det}),
                       file_name="Historial_producto.xlsx", key="mc_dl_ficha")


def _vista_alertas(prod, umbral):
    compras = marcar_alertas(compras_por_documento(prod), umbral)
    alertas = compras[compras["ALERTA"].notna()].sort_values("FECHA", ascending=False)
    conteo = alertas["ALERTA"].value_counts()
    a1, a2, a3 = st.columns(3)
    a1.metric("⬆️ Subieron de precio", int(conteo.get("⬆️ Subió de precio", 0)))
    a2.metric("⬇️ Bajaron de precio", int(conteo.get("⬇️ Bajó de precio", 0)))
    a3.metric("⚠️ Posible cambio de unidad", int(conteo.get("⚠️ Posible cambio de unidad", 0)))
    st.caption(f"Compras cuyo precio varió {umbral:.0f}% o más contra la compra anterior o contra su precio promedio. "
               "'Posible cambio de unidad' = el precio se multiplicó o dividió por 2.5 o más: revisa si cambió la presentación.")

    tipos = st.multiselect("Tipo de alerta:", sorted(alertas["ALERTA"].unique()), key="mc_tipo_alerta")
    if tipos:
        alertas = alertas[alertas["ALERTA"].isin(tipos)]
    vista = alertas[["FECHA", "PRODUCTO", "CATEGORIA", "PROVEEDOR", "CANTIDAD", "PRECIO UNITARIO", "PRECIO ANTERIOR",
                     "VAR PRECIO vs ANTERIOR %", "PRECIO PROMEDIO", "VAR PRECIO vs PROMEDIO %", "MONTO",
                     "DIAS DESDE ANTERIOR", "ALERTA"]].copy()
    vista["FECHA"] = cxp._fmt_fecha(vista["FECHA"])
    st.dataframe(vista, use_container_width=True, hide_index=True,
                 column_config=_cfg(dinero=["MONTO"], precio=["PRECIO UNITARIO", "PRECIO ANTERIOR", "PRECIO PROMEDIO"],
                                    pct=["VAR PRECIO vs ANTERIOR %", "VAR PRECIO vs PROMEDIO %"], num=["CANTIDAD"]))
    st.download_button("📥 Descargar alertas (Excel)", data=cxp._generar_excel({"Alertas": vista}),
                       file_name="Alertas_precios.xlsx", key="mc_dl_alertas")


def _vista_gastos(gastos_todos, gastos):
    if gastos_todos.empty:
        st.info("No hay gastos (códigos que empiezan con G) en los archivos.")
        return

    # Clasificación: reglas por palabras clave; lo que no encaja lo eliges tú (no se guarda en la base)
    descripciones = sorted(gastos_todos["DESCRIPCION"].unique())
    auto = {d: clasificar_gasto(d) for d in descripciones}
    pendientes = [d for d in descripciones if auto[d] is None]
    elegidas = {}
    if pendientes:
        st.warning(f"⚠️ {len(pendientes)} descripciones de gasto no encajan en ninguna categoría. Elige la categoría en la tabla:")
        editor = st.data_editor(
            pd.DataFrame({"DESCRIPCION": pendientes, "CATEGORIA": [None] * len(pendientes)}),
            column_config={"CATEGORIA": st.column_config.SelectboxColumn("CATEGORIA", options=CATEGORIAS_GASTO)},
            disabled=["DESCRIPCION"], hide_index=True, use_container_width=True, key="mc_editor_gastos",
        )
        elegidas = {r["DESCRIPCION"]: r["CATEGORIA"] for _, r in editor.iterrows() if r["CATEGORIA"]}
    with st.expander("Ver cómo se clasificó cada gasto automáticamente"):
        st.dataframe(pd.DataFrame([(d, c) for d, c in auto.items() if c], columns=["DESCRIPCION", "CATEGORIA"]),
                     use_container_width=True, hide_index=True)

    g = gastos.copy()
    g["CATEGORIA GASTO"] = g["DESCRIPCION"].map(lambda d: auto.get(d) or elegidas.get(d) or SIN_CLASIFICAR)

    st.markdown("**Comparativo por categoría de gasto**")
    per_a, per_b = _elegir_periodos(g, "mc_gp")
    if per_a:
        comp = comparar_periodos(g, "CATEGORIA GASTO", per_a, per_b, con_cantidad=False)
        st.dataframe(comp, use_container_width=True, hide_index=True,
                     column_config=_cfg(dinero=["MONTO A", "MONTO B", "VAR MONTO $"], pct=["VAR MONTO %"]))

    serie = g.groupby(["PER INICIO", "PERIODO", "CATEGORIA GASTO"], as_index=False)["MONTO"].sum().sort_values("PER INICIO")
    if not serie.empty:
        st.plotly_chart(px.bar(serie, x="PERIODO", y="MONTO", color="CATEGORIA GASTO", barmode="stack",
                               title="Gastos por categoría en cada período", labels={"MONTO": "Monto ($)", "PERIODO": ""}),
                        use_container_width=True)

    det = g[["FECHA", "PROVEEDOR", "NUMERO", "CODIGO", "DESCRIPCION", "CATEGORIA GASTO", "MONTO"]].sort_values("FECHA", ascending=False).copy()
    det["FECHA"] = cxp._fmt_fecha(det["FECHA"])
    st.markdown("**Detalle de gastos**")
    st.dataframe(det, use_container_width=True, hide_index=True, column_config=_cfg(dinero=["MONTO"]))
    st.download_button("📥 Descargar gastos (Excel)", data=cxp._generar_excel({"Gastos": det}),
                       file_name="Gastos_clasificados.xlsx", key="mc_dl_gastos")


def mostrar_tab_monitoreo(unidad):
    st.markdown(f"#### 🔍 Monitoreo de compras — {unidad}")
    archivos = st.file_uploader(
        "Compras de la unidad (Rpt Compras Producto Proveedor). Puedes subir un período largo o varios archivos a la vez:",
        type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="mc_up_compras")
    if not archivos:
        st.info("Sube el historial de compras de la unidad para empezar.")
        return

    try:
        df, repetidas = leer_historial(archivos)
    except ValueError as e:
        st.error(f"❌ {e}")
        return
    except Exception as e:
        st.error(f"🚨 Error técnico al leer los archivos: {e}")
        return
    if df.empty:
        st.warning("Los archivos no tienen compras con fecha válida.")
        return

    fmin, fmax = df["FECHA"].min().date(), df["FECHA"].max().date()
    n_nc = int(df["ES NOTA CREDITO"].sum())
    st.caption(f"{len(df):,} líneas · {df['NUMERO'].nunique():,} documentos · del {fmin:%d/%m/%Y} al {fmax:%d/%m/%Y}"
               + (f" · {n_nc} líneas de notas de crédito (restan)" if n_nc else "")
               + (f" · {repetidas} líneas repetidas entre archivos eliminadas" if repetidas else ""))

    with st.expander("🎛️ Filtros", expanded=True):
        f1, f2, f3 = st.columns([2, 2, 1])
        rango = f1.date_input("Rango de fechas:", value=(fmin, fmax), min_value=fmin, max_value=fmax, key="mc_rango")
        agrupacion = f2.radio("Agrupar por:", AGRUPACIONES, index=2, horizontal=True, key="mc_agrup")
        umbral = f3.number_input("Alerta de precio (%):", value=UMBRAL_ALERTA, min_value=1.0, step=1.0, key="mc_umbral")
        f4, f5, f6 = st.columns(3)
        cats = f4.multiselect("Categoría:", sorted(df.loc[~df["ES GASTO"], "CATEGORIA"].unique()), key="mc_cat")
        provs = f5.multiselect("Proveedor:", sorted(df["PROVEEDOR"].unique()), key="mc_prov")
        texto = f6.text_input("Buscar producto (código o nombre):", key="mc_texto")

    filtrado = df
    if isinstance(rango, (list, tuple)) and len(rango) == 2:
        filtrado = filtrado[(filtrado["FECHA"].dt.date >= rango[0]) & (filtrado["FECHA"].dt.date <= rango[1])]
    if provs:
        filtrado = filtrado[filtrado["PROVEEDOR"].isin(provs)]
    if texto:
        filtrado = filtrado[filtrado["PRODUCTO"].str.contains(texto.strip(), case=False, regex=False)]
    filtrado = agregar_periodo(filtrado, agrupacion)

    prod = filtrado[~filtrado["ES GASTO"]]
    if cats:
        prod = prod[prod["CATEGORIA"].isin(cats)]
    gastos = filtrado[filtrado["ES GASTO"]]

    t1, t2, t3, t4 = st.tabs(["📊 Período vs período", "🧾 Compra vs compra", "🚨 Alertas", "💼 Gastos (códigos G)"])
    with t1:
        _vista_periodos(prod)
    with t2:
        _vista_ficha(prod, umbral)
    with t3:
        _vista_alertas(prod, umbral)
    with t4:
        _vista_gastos(df[df["ES GASTO"]], gastos)