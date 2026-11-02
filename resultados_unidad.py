import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import validacion_cxp as cxp  # Reutiliza lectores y reglas de IVA de Validación CXP

# ==========================================
# CONFIGURACIÓN
# ==========================================
UNIDADES = cxp.UNIDADES_CXP

TASA_ISSS_PATRONAL = 7.5    # %
TOPE_ISSS = 1000.0          # Salario máximo cotizable ISSS ($)
TASA_AFP_PATRONAL = 8.75    # %

ETIQUETA_FORMA = [
    ("CONTANDO", "Contado"), ("CONTADO", "Contado"), ("EFECTIVO", "Contado"),
    ("TARJETA", "Tarjeta"), ("CREDITO", "Crédito"), ("CHEQUE", "Cheque"),
    ("UBER", "Apps de delivery"), ("PEDIDOSYA", "Apps de delivery"), ("TRANSFER", "Transferencia"),
]


def _forma_pago(texto):
    t = cxp._norm(texto)
    # "TARJETA DE CRÉDITO" contiene CREDITO: por eso se revisa TARJETA antes que CREDITO
    for clave, etiqueta in ETIQUETA_FORMA:
        if clave in t:
            return etiqueta
    return str(texto).strip().title() or "Sin forma de pago"


def _pct(parte, total):
    return (parte / total * 100) if total else 0.0


# ==========================================
# LECTURA DE ARCHIVOS
# ==========================================
def leer_ventas(archivo):
    df = cxp._leer_archivo(archivo, ["DESCFORMAPAGO", "TOTALCOMPROBANTE"])
    if df is None:
        raise ValueError("No encontré los encabezados 'DescFormaPago' y 'TotalComprobante' en el archivo de ventas. "
                         "Verifica que sea el reporte de Ventas por Forma de Pago.")
    col = lambda ops: cxp._buscar_col(df, ops)
    c_forma, c_num, c_fecha, c_cli = col(["DESCFORMAPAGO", "FORMAPAGO"]), col(["NUMERO"]), col(["FECHA"]), col(["NOMBRE"])
    c_tipo_cli, c_suc = col(["TIPOCLIENTE"]), col(["SUCURSAL"])
    c_exento, c_afecto, c_descto = col(["TOTALEXENTO"]), col(["TOTALAFECTO"]), col(["TOTALDESCTO", "DESCUENTO"])
    c_iva, c_ret, c_total = col(["TOTALIVA"]), col(["TOTALIMPUESTO1"]), col(["TOTALCOMPROBANTE"])
    if c_afecto is None or c_iva is None:
        raise ValueError("Al archivo de ventas le faltan las columnas 'TotalAfecto' o 'TotalIva'.")

    num = lambda c: cxp._a_numero(df[c]) if c else 0.0
    txt = lambda c, defecto: df[c].astype(str).str.strip() if c else defecto
    v = pd.DataFrame({
        "FORMA DE PAGO": txt(c_forma, "").apply(_forma_pago),
        "FORMA ORIGINAL": txt(c_forma, ""),
        "NUMERO": txt(c_num, ""),
        "FECHA": cxp._a_fecha(df[c_fecha]) if c_fecha else pd.NaT,
        "CLIENTE": txt(c_cli, ""),
        "TIPO CLIENTE": txt(c_tipo_cli, "SIN TIPO"),
        "SUCURSAL": txt(c_suc, "SIN SUCURSAL"),
        "NETO": num(c_afecto) + num(c_exento) - num(c_descto),
        "IVA": num(c_iva),
        "RETENCION 1%": num(c_ret),
        "TOTAL": num(c_total),
    })
    anulado = v["CLIENTE"].str.upper().str.contains("ANULAD", na=False)
    v = v[~anulado].copy()
    v["BRUTO"] = v["NETO"] + v["IVA"]
    return v, int(anulado.sum())


def leer_compras(archivo):
    """Una fila por documento con su costo, IVA crédito fiscal y categoría principal."""
    df = cxp._leer_archivo(archivo, ["NUMERO", "TOTAL"])
    if df is None:
        raise ValueError("No encontré los encabezados 'Numero' y 'Total' en el archivo de compras.")
    lineas = cxp._leer_compras(archivo)
    c_cat = cxp._buscar_col(df, ["CATEGORIA"])
    lineas["CATEGORIA"] = df.loc[lineas.index, c_cat].astype(str).str.strip().str.upper() if c_cat else "SIN CATEGORÍA"

    # Mismo identificador que Validación CXP: números cortos se separan por proveedor
    es_largo = lineas["LLAVE"].apply(cxp._es_codigo_largo)
    lineas["ID DOC"] = lineas["LLAVE"].where(es_largo, lineas["LLAVE"] + "|" + lineas["PROVEEDOR"].apply(cxp._norm))
    return lineas


def leer_nomina(archivos):
    partes = []
    for arch in archivos:
        df = cxp._leer_archivo(arch, ["DEPARTAMENTO", "TOTALINGRESOS"])
        if df is None:
            raise ValueError(f"No encontré los encabezados 'Departamento' y 'TotalIngresos' en la nómina '{arch.name}'.")
        col = lambda ops: cxp._buscar_col(df, ops)
        num = lambda ops: cxp._a_numero(df[col(ops)]) if col(ops) else 0.0
        txt = lambda ops: df[col(ops)].astype(str).str.strip() if col(ops) else ""
        partes.append(pd.DataFrame({
            "ARCHIVO": arch.name,
            "DEPARTAMENTO": txt(["DEPARTAMENTO"]),
            "NOMBRE": txt(["NOMBRE"]),
            "PLAZA": txt(["PLAZA"]),
            "DIAS LABORADOS": num(["DIASLABORADOS"]),
            "SALARIO": num(["SALARIO"]) if col(["SALARIO"]) else num(["SALARIOORDINARIO"]),
            "HORAS EXTRAS": num(["HORASEXTRAS"]),
            "BONIFICACIONES": num(["BONIFICACIONES"]),
            "AGUINALDO": num(["AGUINALDO"]),
            "VACACION": num(["VACACION"]),
            "VIATICOS": num(["VIATICOS"]),
            "OTROS INGRESOS": num(["OTROSINGRESOS"]),
            "TOTAL INGRESOS": num(["TOTALINGRESOS"]),
        }))
    n = pd.concat(partes, ignore_index=True)
    return n[n["DEPARTAMENTO"].str.len() > 0].copy()


def departamentos_sugeridos(departamentos, unidad):
    clave = cxp._palabras(unidad)
    if not clave:
        return []
    return [d for d in departamentos if clave[0] in cxp._palabras(d)]


# ==========================================
# CÁLCULO DEL RESULTADO
# ==========================================
def calcular(ventas, compras, nomina, deptos, incluir_patronal, tasa_isss, tope_isss, tasa_afp, percepcion=None):
    # ---------- Ingresos ----------
    ing = {
        "neto": ventas["NETO"].sum(), "iva": ventas["IVA"].sum(), "bruto": ventas["BRUTO"].sum(),
        "retencion": ventas["RETENCION 1%"].sum(), "total": ventas["TOTAL"].sum(), "docs": ventas["NUMERO"].nunique(),
    }
    por_forma = ventas.groupby("FORMA DE PAGO").agg(**{
        "DOCUMENTOS": ("NUMERO", "nunique"), "NETO": ("NETO", "sum"), "IVA": ("IVA", "sum"),
        "BRUTO": ("BRUTO", "sum"), "RETENCION 1%": ("RETENCION 1%", "sum"), "TOTAL DOCUMENTO": ("TOTAL", "sum"),
    }).reset_index().sort_values("NETO", ascending=False)
    por_forma["% VENTAS"] = por_forma["NETO"].apply(lambda x: _pct(x, ing["neto"]))

    por_tipo_cli = ventas.groupby("TIPO CLIENTE").agg(**{
        "DOCUMENTOS": ("NUMERO", "nunique"), "NETO": ("NETO", "sum"), "BRUTO": ("BRUTO", "sum"),
    }).reset_index().sort_values("NETO", ascending=False)
    por_tipo_cli["% VENTAS"] = por_tipo_cli["NETO"].apply(lambda x: _pct(x, ing["neto"]))

    por_dia = ventas.dropna(subset=["FECHA"]).groupby("FECHA", as_index=False)["NETO"].sum()

    # ---------- Compras (costo) ----------
    docs = compras.groupby("ID DOC").agg(**{
        "TIPO DOCUMENTO": ("TIPO DOCUMENTO", "first"), "PROVEEDOR": ("PROVEEDOR", "first"),
        "NETO": ("NETO", "sum"), "NO GRAVADO": ("NO GRAVADO", "sum"),
    })
    es_ccf = docs["TIPO DOCUMENTO"].apply(cxp._norm).str.contains("CREDITO")
    docs["IVA"] = ((docs["NETO"] - docs["NO GRAVADO"]) * cxp.TASA_IVA).where(es_ccf, 0.0).round(2)
    com = {
        "costo": compras["NETO"].sum(),
        "iva": docs["IVA"].sum(),
        "docs": len(docs),
        "proveedores": compras["PROVEEDOR"].nunique(),
        "percepcion": percepcion,
    }
    por_categoria = compras.groupby("CATEGORIA").agg(**{
        "LINEAS": ("NETO", "size"), "COSTO": ("NETO", "sum"),
    }).reset_index().sort_values("COSTO", ascending=False)
    por_categoria["% COMPRAS"] = por_categoria["COSTO"].apply(lambda x: _pct(x, com["costo"]))
    por_categoria["% VENTAS"] = por_categoria["COSTO"].apply(lambda x: _pct(x, ing["neto"]))

    por_tipo_doc = docs.groupby("TIPO DOCUMENTO").agg(**{
        "DOCUMENTOS": ("NETO", "size"), "COSTO": ("NETO", "sum"), "IVA CREDITO FISCAL": ("IVA", "sum"),
    }).reset_index().sort_values("COSTO", ascending=False)

    top_prov = docs.groupby("PROVEEDOR").agg(**{"DOCUMENTOS": ("NETO", "size"), "COSTO": ("NETO", "sum")}) \
        .reset_index().sort_values("COSTO", ascending=False)
    top_prov["% COMPRAS"] = top_prov["COSTO"].apply(lambda x: _pct(x, com["costo"]))

    # ---------- Nómina ----------
    nom = nomina[nomina["DEPARTAMENTO"].isin(deptos)].copy()
    base = nom["TOTAL INGRESOS"] - nom["AGUINALDO"] - nom["VIATICOS"]  # aguinaldo y viáticos no cotizan
    nom["ISSS PATRONAL"] = (base.clip(upper=tope_isss) * tasa_isss / 100).round(2)
    nom["AFP PATRONAL"] = (base * tasa_afp / 100).round(2)
    nom["COSTO TOTAL"] = nom["TOTAL INGRESOS"] + nom["ISSS PATRONAL"] + nom["AFP PATRONAL"]
    nomi = {
        "empleados": len(nom),
        "total_ingresos": nom["TOTAL INGRESOS"].sum(),
        "patronal": (nom["ISSS PATRONAL"] + nom["AFP PATRONAL"]).sum(),
    }
    nomi["costo"] = nomi["total_ingresos"] + (nomi["patronal"] if incluir_patronal else 0.0)

    # ---------- Estado de resultados ----------
    margen_bruto = ing["neto"] - com["costo"]
    resultado = margen_bruto - nomi["costo"]
    filas = [
        ("Ventas netas (sin IVA)", ing["neto"]),
        ("(−) Compras (costo)", -com["costo"]),
        ("= Margen bruto", margen_bruto),
        ("(−) Nómina: total ingresos", -nomi["total_ingresos"]),
    ]
    if incluir_patronal:
        filas.append(("(−) Nómina: aporte patronal ISSS/AFP (estimado)", -nomi["patronal"]))
    filas.append(("= Resultado de la unidad", resultado))
    estado = pd.DataFrame(filas, columns=["CONCEPTO", "MONTO"])
    estado["% DE VENTAS"] = estado["MONTO"].apply(lambda x: _pct(x, ing["neto"]))

    # ---------- IVA (indicador de la unidad) ----------
    filas_iva = [
        ("IVA de ventas (débito fiscal)", ing["iva"]),
        ("(−) IVA de compras (crédito fiscal)", -com["iva"]),
    ]
    if percepcion is not None:
        filas_iva.append(("(−) Percepción 1% cobrada por proveedores", -percepcion))
    filas_iva.append(("(−) Retención 1% hecha por clientes", -ing["retencion"]))
    iva_pagar = sum(m for _, m in filas_iva)
    filas_iva.append(("= IVA a pagar (aporte de la unidad)", iva_pagar))
    tabla_iva = pd.DataFrame(filas_iva, columns=["CONCEPTO", "MONTO"])

    # ---------- Por sucursal (compras y nómina repartidas según % de ventas) ----------
    suc = ventas.groupby("SUCURSAL").agg(**{"DOCUMENTOS": ("NUMERO", "nunique"), "VENTAS NETAS": ("NETO", "sum")}).reset_index()
    suc["% PARTICIPACION"] = suc["VENTAS NETAS"].apply(lambda x: _pct(x, ing["neto"]))
    suc["COMPRAS ASIGNADAS"] = com["costo"] * suc["% PARTICIPACION"] / 100
    suc["NOMINA ASIGNADA"] = nomi["costo"] * suc["% PARTICIPACION"] / 100
    suc["RESULTADO"] = suc["VENTAS NETAS"] - suc["COMPRAS ASIGNADAS"] - suc["NOMINA ASIGNADA"]
    suc["% RESULTADO"] = suc.apply(lambda r: _pct(r["RESULTADO"], r["VENTAS NETAS"]), axis=1)
    suc = suc.sort_values("VENTAS NETAS", ascending=False)

    return {
        "ing": ing, "com": com, "nomi": nomi, "resultado": resultado, "margen_bruto": margen_bruto,
        "estado": estado, "tabla_iva": tabla_iva, "iva_pagar": iva_pagar,
        "por_forma": por_forma.round(2), "por_tipo_cli": por_tipo_cli.round(2), "por_dia": por_dia,
        "por_categoria": por_categoria.round(2), "por_tipo_doc": por_tipo_doc.round(2), "top_prov": top_prov.round(2),
        "nomina_det": nom, "por_sucursal": suc.round(2),
        "periodo_ventas": (ventas["FECHA"].min(), ventas["FECHA"].max()),
        "periodo_compras": (compras["FECHA COMPRA"].min(), compras["FECHA COMPRA"].max()),
        "incluir_patronal": incluir_patronal,
    }


# ==========================================
# PANTALLA
# ==========================================
def _rango(par):
    a, b = par
    if pd.isna(a):
        return "sin fecha"
    return f"{a.strftime('%d/%m/%Y')} al {b.strftime('%d/%m/%Y')}"


def _mostrar(r, unidad, sucursales_sel):
    ing, com, nomi = r["ing"], r["com"], r["nomi"]
    st.markdown(f"### 📌 {unidad}" + (f" — {', '.join(sucursales_sel)}" if sucursales_sel else ""))
    st.caption(f"Ventas del {_rango(r['periodo_ventas'])} · Compras del {_rango(r['periodo_compras'])}")
    mv, mc = r["periodo_ventas"][0], r["periodo_compras"][0]
    if pd.notna(mv) and pd.notna(mc) and (mv.year, mv.month) != (mc.year, mc.month):
        st.warning("⚠️ Las ventas y las compras parecen ser de meses distintos. Revisa los archivos.")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Ventas netas", f"${ing['neto']:,.2f}")
    m2.metric("Compras (costo)", f"${com['costo']:,.2f}", f"{_pct(com['costo'], ing['neto']):.1f}% de ventas", delta_color="off")
    m3.metric("Nómina", f"${nomi['costo']:,.2f}", f"{_pct(nomi['costo'], ing['neto']):.1f}% de ventas", delta_color="off")
    m4.metric("Resultado", f"${r['resultado']:,.2f}", f"{_pct(r['resultado'], ing['neto']):.1f}% de ventas",
              delta_color="normal" if r["resultado"] >= 0 else "inverse")

    t_res, t_suc, t_ing, t_com, t_nom, t_iva = st.tabs([
        "📊 Resumen", "🏪 Por sucursal", "💵 Ingresos", "🛒 Compras", "👥 Nómina", "🧾 IVA"
    ])
    dinero = cxp._config_dinero
    pct_cfg = lambda cols: {c: st.column_config.NumberColumn(c, format="%.1f%%") for c in cols}

    with t_res:
        c1, c2 = st.columns([3, 2])
        with c1:
            fig = go.Figure(go.Waterfall(
                x=["Ventas netas", "Compras", "Nómina", "Resultado"],
                measure=["absolute", "relative", "relative", "total"],
                y=[ing["neto"], -com["costo"], -nomi["costo"], 0],
                text=[f"${ing['neto']:,.0f}", f"-${com['costo']:,.0f}", f"-${nomi['costo']:,.0f}", f"${r['resultado']:,.0f}"],
                textposition="outside",
            ))
            fig.update_layout(title="De las ventas al resultado", showlegend=False, height=420)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            st.markdown("**Estado de resultados**")
            st.dataframe(r["estado"], use_container_width=True, hide_index=True,
                         column_config={**dinero(["MONTO"]), **pct_cfg(["% DE VENTAS"])})
            st.caption(f"IVA a pagar (aporte de la unidad): **${r['iva_pagar']:,.2f}** · ver pestaña IVA")

    with t_suc:
        st.info("Las ventas son reales por sucursal. Las compras y la nómina no vienen separadas por sucursal, "
                "así que se reparten según el % de ventas de cada una (estimado).")
        st.dataframe(r["por_sucursal"], use_container_width=True, hide_index=True,
                     column_config={**dinero(["VENTAS NETAS", "COMPRAS ASIGNADAS", "NOMINA ASIGNADA", "RESULTADO"]),
                                    **pct_cfg(["% PARTICIPACION", "% RESULTADO"])})
        st.plotly_chart(px.bar(r["por_sucursal"], x="SUCURSAL", y="VENTAS NETAS", text_auto=",.0f",
                               title="Ventas netas por sucursal"), use_container_width=True)

    with t_ing:
        i1, i2, i3 = st.columns(3)
        i1.metric("Ventas brutas (con IVA)", f"${ing['bruto']:,.2f}")
        i2.metric("IVA de ventas", f"${ing['iva']:,.2f}")
        i3.metric("Documentos", f"{ing['docs']:,}")
        st.markdown("**Por forma de pago**")
        c1, c2 = st.columns([3, 2])
        with c1:
            st.dataframe(r["por_forma"], use_container_width=True, hide_index=True,
                         column_config={**dinero(["NETO", "IVA", "BRUTO", "RETENCION 1%", "TOTAL DOCUMENTO"]), **pct_cfg(["% VENTAS"])})
        with c2:
            st.plotly_chart(px.pie(r["por_forma"], values="NETO", names="FORMA DE PAGO", hole=0.4), use_container_width=True)
        st.markdown("**Por tipo de cliente**")
        st.dataframe(r["por_tipo_cli"], use_container_width=True, hide_index=True,
                     column_config={**dinero(["NETO", "BRUTO"]), **pct_cfg(["% VENTAS"])})
        if not r["por_dia"].empty:
            st.plotly_chart(px.line(r["por_dia"], x="FECHA", y="NETO", markers=True, title="Ventas netas por día",
                                    labels={"NETO": "Ventas netas ($)", "FECHA": "Fecha"}), use_container_width=True)

    with t_com:
        k1, k2, k3 = st.columns(3)
        k1.metric("Costo de compras", f"${com['costo']:,.2f}")
        k2.metric("IVA crédito fiscal", f"${com['iva']:,.2f}")
        k3.metric("Documentos / proveedores", f"{com['docs']} / {com['proveedores']}")
        st.caption("Costo = neto de los créditos fiscales (su IVA se recupera) + total de Sujeto Excluido y Facturas (su IVA va al costo).")
        c1, c2 = st.columns([3, 2])
        with c1:
            st.markdown("**Por categoría**")
            st.dataframe(r["por_categoria"], use_container_width=True, hide_index=True,
                         column_config={**dinero(["COSTO"]), **pct_cfg(["% COMPRAS", "% VENTAS"])})
        with c2:
            st.plotly_chart(px.pie(r["por_categoria"], values="COSTO", names="CATEGORIA", hole=0.4), use_container_width=True)
        st.markdown("**Por tipo de documento**")
        st.dataframe(r["por_tipo_doc"], use_container_width=True, hide_index=True,
                     column_config=dinero(["COSTO", "IVA CREDITO FISCAL"]))
        st.markdown("**Proveedores (de mayor a menor)**")
        st.dataframe(r["top_prov"], use_container_width=True, hide_index=True,
                     column_config={**dinero(["COSTO"]), **pct_cfg(["% COMPRAS"])})

    with t_nom:
        n1, n2, n3 = st.columns(3)
        n1.metric("Empleados", nomi["empleados"])
        n2.metric("Total ingresos", f"${nomi['total_ingresos']:,.2f}")
        n3.metric("Aporte patronal (estimado)", f"${nomi['patronal']:,.2f}",
                  "incluido en el resultado" if r["incluir_patronal"] else "no incluido en el resultado", delta_color="off")
        det = r["nomina_det"][["DEPARTAMENTO", "NOMBRE", "PLAZA", "DIAS LABORADOS", "SALARIO", "HORAS EXTRAS",
                               "BONIFICACIONES", "AGUINALDO", "VACACION", "OTROS INGRESOS", "VIATICOS",
                               "TOTAL INGRESOS", "ISSS PATRONAL", "AFP PATRONAL", "COSTO TOTAL", "ARCHIVO"]]
        st.dataframe(det, use_container_width=True, hide_index=True,
                     column_config=dinero(["SALARIO", "HORAS EXTRAS", "BONIFICACIONES", "AGUINALDO", "VACACION",
                                           "OTROS INGRESOS", "VIATICOS", "TOTAL INGRESOS", "ISSS PATRONAL",
                                           "AFP PATRONAL", "COSTO TOTAL"]))

    with t_iva:
        st.dataframe(r["tabla_iva"], use_container_width=True, hide_index=True, column_config=dinero(["MONTO"]))
        if com["percepcion"] is None:
            st.warning("⚠️ La percepción 1% de compras no está incluida: valida primero esta unidad en "
                       "Contabilidad → 7. Validación CXP y vuelve a calcular.")
        st.caption("El IVA se declara por toda la Fundación. Este cálculo indica cuánto aporta esta unidad al IVA a pagar del mes.")

    hojas = {
        "Estado de resultados": r["estado"], "IVA": r["tabla_iva"], "Por sucursal": r["por_sucursal"],
        "Ventas por forma de pago": r["por_forma"], "Ventas por tipo cliente": r["por_tipo_cli"],
        "Compras por categoria": r["por_categoria"], "Compras por tipo doc": r["por_tipo_doc"],
        "Compras por proveedor": r["top_prov"], "Nomina": r["nomina_det"],
    }
    st.download_button("📥 Descargar Resultados (Excel)", data=cxp._generar_excel(hojas),
                       file_name=f"Resultados_{cxp._norm(unidad).title()}.xlsx", use_container_width=True, key="ru_descarga")


def mostrar_modulo_resultados():
    st.title("📈 Resultados por Unidad")
    st.caption("Ingresos contra compras y nómina de una unidad, con su IVA. Solo consulta: no guarda nada en la base.")

    unidad = st.selectbox("1. Unidad:", UNIDADES, key="ru_unidad")
    c1, c2, c3 = st.columns(3)
    with c1:
        arch_ventas = st.file_uploader("2. Ventas por forma de pago (VentasFormaPago):", type=["xlsx", "xls", "csv"], key="ru_up_ventas")
    with c2:
        arch_compras = st.file_uploader("3. Compras de la unidad (Rpt Compras Producto Proveedor):", type=["xlsx", "xls", "csv"], key="ru_up_compras")
    with c3:
        arch_nomina = st.file_uploader("4. Nómina (una o varias planillas del mes):", type=["xlsx", "xls", "csv"],
                                       accept_multiple_files=True, key="ru_up_nomina")

    if not (arch_ventas and arch_compras and arch_nomina):
        st.info("Sube los tres archivos para calcular el resultado.")
        return

    try:
        ventas, n_anulados = leer_ventas(arch_ventas)
        compras = leer_compras(arch_compras)
        nomina = leer_nomina(arch_nomina)
    except ValueError as e:
        st.error(f"❌ {e}")
        return
    except Exception as e:
        st.error(f"🚨 Error técnico al leer los archivos: {e}")
        return

    with st.expander("⚙️ Opciones de cálculo", expanded=True):
        o1, o2 = st.columns(2)
        with o1:
            sucursales = sorted(ventas["SUCURSAL"].unique())
            suc_sel = st.multiselect("Sucursales a incluir (vacío = todas):", sucursales, key="ru_sucursales")
            deptos = sorted(nomina["DEPARTAMENTO"].unique())
            deptos_sel = st.multiselect("Departamentos de la nómina que pertenecen a la unidad:", deptos,
                                        default=departamentos_sugeridos(deptos, unidad), key=f"ru_deptos_{unidad}")
        with o2:
            incluir_patronal = st.checkbox("Incluir aporte patronal ISSS/AFP en el resultado (estimado)", value=False, key="ru_patronal")
            p1, p2, p3 = st.columns(3)
            tasa_isss = p1.number_input("ISSS patronal %", value=TASA_ISSS_PATRONAL, step=0.25, key="ru_isss")
            tope_isss = p2.number_input("Tope ISSS $", value=TOPE_ISSS, step=50.0, key="ru_tope")
            tasa_afp = p3.number_input("AFP patronal %", value=TASA_AFP_PATRONAL, step=0.25, key="ru_afp")
        if n_anulados:
            st.caption(f"Se excluyeron {n_anulados} documentos de venta anulados.")

    if not deptos_sel:
        st.warning("⚠️ Selecciona al menos un departamento de la nómina para la unidad.")

    ventas_f = ventas[ventas["SUCURSAL"].isin(suc_sel)] if suc_sel else ventas
    if suc_sel:
        st.caption("Con sucursales filtradas, las compras y la nómina se mantienen completas (no vienen por sucursal). "
                   "Para ver el reparto estimado por sucursal usa la pestaña 'Por sucursal' sin filtro.")

    percepcion = None
    if "cxp_resultado" in st.session_state and st.session_state.get("cxp_unidad_proc") == unidad:
        percepcion = float(st.session_state["cxp_resultado"]["kpis"]["percepcion"])

    try:
        r = calcular(ventas_f, compras, nomina, deptos_sel, incluir_patronal, tasa_isss, tope_isss, tasa_afp, percepcion)
    except Exception as e:
        st.error(f"🚨 Error técnico al calcular el resultado: {e}")
        return

    st.divider()
    _mostrar(r, unidad, suc_sel)