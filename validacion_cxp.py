import io
import re
import unicodedata

import pandas as pd
import plotly.express as px
import streamlit as st

# ==========================================
# CONFIGURACIÓN GENERAL DE LA PESTAÑA CXP
# ==========================================
UNIDADES_CXP = [
    "Cafetería", "Librería", "CID Campus", "Talleres Gráfico", "Despensa",
    "Terraza", "Servicios Generales", "Gerencias", "Contabilidad"
]

TASA_IVA = 0.13            # IVA El Salvador
TASA_PERCEPCION = 0.01     # Percepción 1% que cobran los grandes contribuyentes
MINIMO_PERCEPCION = 100.0  # La percepción solo aplica a documentos con neto de $100 o más
TOLERANCIA = 0.10          # Diferencia máxima ($) para considerar que un abono cuadra (redondeos)

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}
DIAS_ES = {0: "Lunes", 1: "Martes", 2: "Miércoles", 3: "Jueves", 4: "Viernes", 5: "Sábado", 6: "Domingo"}

EST_CANCELADO = "CANCELADO"
EST_PERCEPCION = "CANCELADO (PERCEPCIÓN 1%)"
EST_PENDIENTE = "PENDIENTE"
EST_PARCIAL = "ABONO PARCIAL"
EST_DE_MAS = "PAGADO DE MÁS"


# ==========================================
# FUNCIONES AUXILIARES
# ==========================================
def _norm(texto):
    """Quita tildes, espacios y símbolos para comparar nombres de columnas."""
    t = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode().upper()
    return re.sub(r"[^A-Z0-9]", "", t)


def _buscar_col(df, opciones):
    """Busca una columna primero por nombre exacto y luego por coincidencia parcial."""
    columnas = {_norm(c): c for c in df.columns}
    for op in opciones:
        if _norm(op) in columnas:
            return columnas[_norm(op)]
    for op in opciones:
        for k, c in columnas.items():
            if _norm(op) in k:
                return c
    return None


def _leer_archivo(archivo, claves_encabezado):
    """Lee Excel/CSV de Nexus y ubica la fila de encabezados aunque haya filas basura arriba."""
    archivo.seek(0)
    if archivo.name.lower().endswith(".csv"):
        linea = archivo.readline().decode("utf-8", errors="ignore")
        archivo.seek(0)
        separador = ";" if ";" in linea else ","
        crudo = pd.read_csv(archivo, sep=separador, dtype=str, header=None)
    else:
        try:
            crudo = pd.read_excel(archivo, dtype=str, header=None)
        except Exception:
            archivo.seek(0)
            tabla = pd.read_html(io.StringIO(archivo.read().decode("utf-8", errors="ignore")))[0].astype(str)
            crudo = pd.DataFrame([list(tabla.columns)] + tabla.values.tolist())

    claves = [_norm(k) for k in claves_encabezado]
    for i in range(min(25, len(crudo))):
        fila = [_norm(v) for v in crudo.iloc[i].tolist()]
        if all(k in fila for k in claves):
            df = crudo.iloc[i + 1:].copy()
            df.columns = [str(v).strip() for v in crudo.iloc[i].tolist()]
            df = df.loc[:, ~pd.Index(df.columns).duplicated()]
            df = df.dropna(how="all").reset_index(drop=True)
            return df
    return None


def _a_numero(serie):
    return pd.to_numeric(
        serie.astype(str).str.replace(r"[^\d\.\-]", "", regex=True), errors="coerce"
    ).fillna(0.0)


def _a_fecha(serie):
    """Convierte fechas tipo 01/09/26, 17/09/2026 o 2026-09-17 (día primero)."""
    s = serie.astype(str).str.strip().str.split(" ").str[0]
    fecha = pd.to_datetime(s, format="%d/%m/%y", errors="coerce")
    fecha = fecha.fillna(pd.to_datetime(s, format="%d/%m/%Y", errors="coerce"))
    fecha = fecha.fillna(pd.to_datetime(s, format="%Y-%m-%d", errors="coerce"))
    return fecha


def _llave(serie):
    """Número de documento normalizado: quita el prefijo 'CF ' del archivo de abonos."""
    return (
        serie.astype(str).str.upper()
        .str.replace(r"^\s*CF\s+", "", regex=True)
        .str.replace(r"\s+", "", regex=True)
    )


def _etiqueta_mes(fecha):
    if pd.isna(fecha):
        return "SIN FECHA"
    return f"{fecha.year}-{fecha.month:02d} {MESES_ES[fecha.month]}"


def _fmt_fecha(serie):
    return serie.dt.strftime("%d/%m/%Y").fillna("")


def _config_dinero(columnas):
    return {c: st.column_config.NumberColumn(c, format="$%.2f") for c in columnas}


def _generar_excel(dict_dfs):
    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:
        for nombre, df in dict_dfs.items():
            df.to_excel(writer, index=False, sheet_name=nombre[:31])
    return salida.getvalue()


# ==========================================
# LÓGICA PRINCIPAL DEL CRUCE
# ==========================================
def procesar_cxp(arch_compras, arch_abonos):
    # ---------- 1. COMPRAS (archivo de la unidad) ----------
    df_c = _leer_archivo(arch_compras, ["NUMERO", "TOTAL"])
    if df_c is None:
        raise ValueError("No encontré los encabezados 'Numero' y 'Total' en el archivo de compras. "
                         "Verifica que sea el reporte de Compras por Producto/Proveedor.")

    c_num = _buscar_col(df_c, ["NUMERO", "NUMERODOCUMENTO"])
    c_fecha = _buscar_col(df_c, ["FECHA"])
    c_nom = _buscar_col(df_c, ["NOMBRE", "PROVEEDOR"])
    c_total = _buscar_col(df_c, ["TOTAL"])
    c_tipo = _buscar_col(df_c, ["TIPODOCUMENTO", "TIPODOC"])

    faltan = [n for n, c in [("Numero", c_num), ("Fecha", c_fecha), ("Nombre", c_nom), ("Total", c_total)] if c is None]
    if faltan:
        raise ValueError(f"Al archivo de compras le faltan las columnas: {', '.join(faltan)}.")

    compras = pd.DataFrame({
        "LLAVE": _llave(df_c[c_num]),
        "NUMERO DOCUMENTO": df_c[c_num].astype(str).str.strip(),
        "PROVEEDOR": df_c[c_nom].astype(str).str.strip(),
        "TIPO DOCUMENTO": df_c[c_tipo].astype(str).str.strip() if c_tipo else "NO IDENTIFICADO",
        "FECHA COMPRA": _a_fecha(df_c[c_fecha]),
        "NETO": _a_numero(df_c[c_total]),
    })
    compras = compras[~compras["LLAVE"].isin(["", "NAN", "NONE"])]

    # Una fila por documento (el reporte viene una fila por producto)
    docs = compras.groupby("LLAVE", as_index=False).agg(**{
        "NUMERO DOCUMENTO": ("NUMERO DOCUMENTO", "first"),
        "PROVEEDOR": ("PROVEEDOR", "first"),
        "TIPO DOCUMENTO": ("TIPO DOCUMENTO", "first"),
        "FECHA COMPRA": ("FECHA COMPRA", "min"),
        "NETO": ("NETO", "sum"),
        "LINEAS": ("NETO", "size"),
    })

    es_ccf = docs["TIPO DOCUMENTO"].apply(_norm).str.contains("CREDITO")
    docs["IVA"] = (docs["NETO"] * TASA_IVA).where(es_ccf, 0.0).round(2)
    docs["TOTAL CON IVA"] = (docs["NETO"] + docs["IVA"]).round(2)
    percepcion_posible = (docs["NETO"] * TASA_PERCEPCION).where(es_ccf & (docs["NETO"] >= MINIMO_PERCEPCION), 0.0).round(2)

    # ---------- 2. ABONOS (archivo completo, se usan solo los que cruzan) ----------
    df_a = _leer_archivo(arch_abonos, ["NUMCOMPRA", "MONTOABONADO"])
    if df_a is None:
        raise ValueError("No encontré los encabezados 'NumCompra' y 'MontoAbonado' en el archivo de abonos. "
                         "Verifica que sea el reporte de Listado de Abonos.")

    a_num = _buscar_col(df_a, ["NUMCOMPRA", "NUMEROCOMPRA"])
    a_monto = _buscar_col(df_a, ["MONTOABONADO", "ABONO"])
    a_fecha = _buscar_col(df_a, ["FECHA"])
    a_comp = _buscar_col(df_a, ["NUMEROCOMPROBANTE", "COMPROBANTE"])
    a_nom = _buscar_col(df_a, ["NOMBRE"])

    abonos = pd.DataFrame({
        "LLAVE": _llave(df_a[a_num]),
        "NUM COMPRA (ABONO)": df_a[a_num].astype(str).str.strip(),
        "PROVEEDOR (ABONO)": df_a[a_nom].astype(str).str.strip() if a_nom else "",
        "FECHA PAGO": _a_fecha(df_a[a_fecha]) if a_fecha else pd.NaT,
        "MONTO ABONADO": _a_numero(df_a[a_monto]),
        "COMPROBANTE": df_a[a_comp].astype(str).str.strip() if a_comp else "",
    })
    total_filas_abonos = len(abonos)
    abonos = abonos[abonos["LLAVE"].isin(docs["LLAVE"])].copy()

    ab_agr = abonos.groupby("LLAVE", as_index=False).agg(**{
        "ABONADO": ("MONTO ABONADO", "sum"),
        "N ABONOS": ("MONTO ABONADO", "size"),
        "PRIMER PAGO": ("FECHA PAGO", "min"),
        "FECHA PAGO": ("FECHA PAGO", "max"),
        "COMPROBANTES": ("COMPROBANTE", lambda s: ", ".join(sorted(set(s)))),
    })

    det = docs.merge(ab_agr, on="LLAVE", how="left")
    det["ABONADO"] = det["ABONADO"].fillna(0.0).round(2)
    det["N ABONOS"] = det["N ABONOS"].fillna(0).astype(int)
    det["COMPROBANTES"] = det["COMPROBANTES"].fillna("")
    det["PERCEPCION POSIBLE"] = percepcion_posible

    # ---------- 3. ESTADO DE CADA DOCUMENTO ----------
    def _estado(r):
        if r["N ABONOS"] == 0:
            return EST_PENDIENTE
        if abs(r["ABONADO"] - r["TOTAL CON IVA"]) <= TOLERANCIA:
            return EST_CANCELADO
        if r["PERCEPCION POSIBLE"] > 0 and abs(r["ABONADO"] - (r["TOTAL CON IVA"] + r["PERCEPCION POSIBLE"])) <= TOLERANCIA:
            return EST_PERCEPCION
        if r["ABONADO"] < r["TOTAL CON IVA"]:
            return EST_PARCIAL
        return EST_DE_MAS

    det["ESTADO"] = det.apply(_estado, axis=1)

    # Percepción: real en los cancelados con 1%, estimada en pendientes de proveedores que la cobran
    prov_percepcion = set(det.loc[det["ESTADO"] == EST_PERCEPCION, "PROVEEDOR"])
    estimada = (det["ESTADO"] == EST_PENDIENTE) & det["PROVEEDOR"].isin(prov_percepcion) & (det["PERCEPCION POSIBLE"] > 0)
    det["PERCEPCION 1%"] = det["PERCEPCION POSIBLE"].where((det["ESTADO"] == EST_PERCEPCION) | estimada, 0.0)
    det["TOTAL A PAGAR"] = (det["TOTAL CON IVA"] + det["PERCEPCION 1%"]).round(2)
    det["SALDO"] = (det["TOTAL A PAGAR"] - det["ABONADO"]).round(2)
    det.loc[det["ESTADO"].isin([EST_CANCELADO, EST_PERCEPCION]), "SALDO"] = 0.0
    det["OBSERVACION"] = ""
    det.loc[estimada, "OBSERVACION"] = "Incluye percepción 1% estimada (el proveedor la cobra)"
    det.loc[det["N ABONOS"] > 1, "OBSERVACION"] = det["OBSERVACION"].where(det["OBSERVACION"] == "", det["OBSERVACION"] + " | ") + "Pagado en varios abonos"

    det["MES COMPRA"] = det["FECHA COMPRA"].apply(_etiqueta_mes)
    det["MES PAGO"] = det["FECHA PAGO"].apply(lambda f: _etiqueta_mes(f) if pd.notna(f) else "PENDIENTE")
    det["DIAS PARA PAGO"] = (det["FECHA PAGO"] - det["FECHA COMPRA"]).dt.days

    det = det.sort_values(["FECHA COMPRA", "PROVEEDOR"]).reset_index(drop=True)

    # ---------- 4. CORTES: MES DE COMPRA vs MES DE PAGO ----------
    abonos = abonos.merge(det[["LLAVE", "MES COMPRA", "PROVEEDOR", "NUMERO DOCUMENTO"]], on="LLAVE", how="left")
    abonos["MES PAGO"] = abonos["FECHA PAGO"].apply(_etiqueta_mes)

    pagado_mes = abonos.groupby(["MES COMPRA", "MES PAGO"], as_index=False)["MONTO ABONADO"].sum()
    pendiente_mes = det[det["SALDO"] > 0].groupby("MES COMPRA", as_index=False)["SALDO"].sum()
    pendiente_mes = pendiente_mes.rename(columns={"SALDO": "MONTO ABONADO"})
    pendiente_mes["MES PAGO"] = "PENDIENTE"
    cortes_largo = pd.concat([pagado_mes, pendiente_mes], ignore_index=True).rename(columns={"MONTO ABONADO": "MONTO"})

    cortes = cortes_largo.pivot_table(index="MES COMPRA", columns="MES PAGO", values="MONTO", aggfunc="sum", fill_value=0.0)
    columnas_orden = sorted([c for c in cortes.columns if c != "PENDIENTE"]) + (["PENDIENTE"] if "PENDIENTE" in cortes.columns else [])
    cortes = cortes[columnas_orden]
    comprado_mes = det.groupby("MES COMPRA")["TOTAL A PAGAR"].sum()
    cortes.insert(0, "TOTAL COMPRADO", comprado_mes)
    cortes = cortes.round(2).reset_index()
    cortes.columns.name = None

    # ---------- 5. PAGOS POR SEMANA (lunes a domingo) ----------
    ab_fecha = abonos.dropna(subset=["FECHA PAGO"]).copy()
    if not ab_fecha.empty:
        ab_fecha["INICIO"] = ab_fecha["FECHA PAGO"] - pd.to_timedelta(ab_fecha["FECHA PAGO"].dt.weekday, unit="D")
        semanas = ab_fecha.groupby("INICIO").agg(**{
            "DOCUMENTOS": ("LLAVE", "nunique"),
            "PROVEEDORES": ("PROVEEDOR", "nunique"),
            "MONTO PAGADO": ("MONTO ABONADO", "sum"),
            "DIAS DE PAGO": ("FECHA PAGO", lambda s: ", ".join(
                f"{DIAS_ES[d.weekday()]} {d.strftime('%d/%m')}" for d in sorted(set(s)))),
        }).reset_index().sort_values("INICIO")
        semanas.insert(0, "SEMANA", semanas["INICIO"].apply(
            lambda d: f"{d.strftime('%d/%m/%Y')} al {(d + pd.Timedelta(days=6)).strftime('%d/%m/%Y')}"))
        semanas["MES PAGO"] = semanas["INICIO"].apply(_etiqueta_mes)
        semanas["MONTO PAGADO"] = semanas["MONTO PAGADO"].round(2)
        semanas = semanas.drop(columns=["INICIO"])
    else:
        semanas = pd.DataFrame(columns=["SEMANA", "DOCUMENTOS", "PROVEEDORES", "MONTO PAGADO", "DIAS DE PAGO", "MES PAGO"])

    # ---------- 6. RESUMEN POR PROVEEDOR ----------
    prov = det.groupby("PROVEEDOR").agg(**{
        "DOCUMENTOS": ("LLAVE", "size"),
        "NETO": ("NETO", "sum"),
        "IVA": ("IVA", "sum"),
        "PERCEPCION 1%": ("PERCEPCION 1%", "sum"),
        "TOTAL A PAGAR": ("TOTAL A PAGAR", "sum"),
        "ABONADO": ("ABONADO", "sum"),
        "SALDO": ("SALDO", "sum"),
        "DOCS PENDIENTES": ("ESTADO", lambda s: (s == EST_PENDIENTE).sum()),
        "DOCS CON DIFERENCIA": ("ESTADO", lambda s: s.isin([EST_PARCIAL, EST_DE_MAS]).sum()),
        "PROMEDIO DIAS PAGO": ("DIAS PARA PAGO", "mean"),
    }).reset_index()
    prov = prov.round(2).sort_values("TOTAL A PAGAR", ascending=False).reset_index(drop=True)

    # ---------- 7. INDICADORES ----------
    kpis = {
        "fecha_min": det["FECHA COMPRA"].min(),
        "fecha_max": det["FECHA COMPRA"].max(),
        "docs": len(det),
        "proveedores": det["PROVEEDOR"].nunique(),
        "neto": det["NETO"].sum(),
        "iva": det["IVA"].sum(),
        "percepcion": det["PERCEPCION 1%"].sum(),
        "total_pagar": det["TOTAL A PAGAR"].sum(),
        "abonado": det["ABONADO"].sum(),
        "pendiente": det.loc[det["SALDO"] > 0, "SALDO"].sum(),
        "por_estado": det["ESTADO"].value_counts().to_dict(),
        "abonos_archivo": total_filas_abonos,
        "abonos_cruzados": len(abonos),
        "promedio_dias": det["DIAS PARA PAGO"].mean(),
    }

    columnas_detalle = [
        "PROVEEDOR", "NUMERO DOCUMENTO", "TIPO DOCUMENTO", "FECHA COMPRA", "MES COMPRA",
        "NETO", "IVA", "PERCEPCION 1%", "TOTAL A PAGAR", "ABONADO", "SALDO", "ESTADO",
        "FECHA PAGO", "MES PAGO", "DIAS PARA PAGO", "N ABONOS", "COMPROBANTES", "OBSERVACION",
    ]
    abonos_det = abonos[["PROVEEDOR", "NUMERO DOCUMENTO", "MES COMPRA", "FECHA PAGO", "MES PAGO",
                         "MONTO ABONADO", "COMPROBANTE", "NUM COMPRA (ABONO)"]].sort_values("FECHA PAGO")

    return {
        "detalle": det[columnas_detalle],
        "cortes": cortes,
        "cortes_largo": cortes_largo,
        "semanas": semanas,
        "proveedores": prov,
        "abonos": abonos_det.reset_index(drop=True),
        "kpis": kpis,
    }


# ==========================================
# PANTALLA DE LA PESTAÑA
# ==========================================
def _mostrar_resultados(res, unidad):
    k = res["kpis"]
    det = res["detalle"]

    periodo = ""
    if pd.notna(k["fecha_min"]):
        periodo = f" | Compras del {k['fecha_min'].strftime('%d/%m/%Y')} al {k['fecha_max'].strftime('%d/%m/%Y')}"
    st.markdown(f"### 📌 Resultado: {unidad}{periodo}")

    pct_pagado = (k["abonado"] / k["total_pagar"] * 100) if k["total_pagar"] > 0 else 0
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Comprado (con IVA)", f"${k['total_pagar']:,.2f}")
    m2.metric("Pagado", f"${k['abonado']:,.2f}", f"{pct_pagado:.1f}% del total", delta_color="off")
    m3.metric("Saldo pendiente", f"${k['pendiente']:,.2f}")
    m4.metric("Días promedio de pago", f"{k['promedio_dias']:.0f} días" if pd.notna(k["promedio_dias"]) else "-")

    pe = k["por_estado"]
    e1, e2, e3, e4 = st.columns(4)
    e1.metric("Documentos", f"{k['docs']} ({k['proveedores']} proveedores)")
    e2.metric("✅ Cancelados", pe.get(EST_CANCELADO, 0) + pe.get(EST_PERCEPCION, 0))
    e3.metric("⏳ Pendientes", pe.get(EST_PENDIENTE, 0))
    e4.metric("⚠️ Con diferencia", pe.get(EST_PARCIAL, 0) + pe.get(EST_DE_MAS, 0))

    st.caption(
        f"Neto: ${k['neto']:,.2f} · IVA crédito fiscal: ${k['iva']:,.2f} · Percepción 1%: ${k['percepcion']:,.2f} · "
        f"Abonos cruzados: {k['abonos_cruzados']} de {k['abonos_archivo']} filas del archivo de abonos"
    )

    t_cortes, t_semanas, t_prov, t_det, t_abonos = st.tabs([
        "📊 Cortes por mes", "📅 Pagos semanales", "🏢 Por proveedor", "📄 Detalle por documento", "💳 Abonos cruzados"
    ])

    with t_cortes:
        st.markdown("**¿En qué mes se pagó lo que se compró en cada mes?**")
        cl = res["cortes_largo"]
        if not cl.empty:
            fig = px.bar(cl, x="MES COMPRA", y="MONTO", color="MES PAGO", barmode="stack",
                         text_auto=",.0f", labels={"MONTO": "Monto ($)", "MES COMPRA": "Mes de compra", "MES PAGO": "Mes de pago"})
            st.plotly_chart(fig, use_container_width=True)
        cols_dinero = [c for c in res["cortes"].columns if c != "MES COMPRA"]
        st.dataframe(res["cortes"], use_container_width=True, hide_index=True, column_config=_config_dinero(cols_dinero))

    with t_semanas:
        st.markdown("**Pagos realizados por semana (lunes a domingo)**")
        sem = res["semanas"]
        if not sem.empty:
            fig2 = px.bar(sem, x="SEMANA", y="MONTO PAGADO", color="MES PAGO", text_auto=",.0f",
                          labels={"MONTO PAGADO": "Monto pagado ($)", "SEMANA": "Semana"})
            st.plotly_chart(fig2, use_container_width=True)
            st.dataframe(sem, use_container_width=True, hide_index=True, column_config=_config_dinero(["MONTO PAGADO"]))
        else:
            st.info("No hay pagos con fecha para mostrar.")

    with t_prov:
        st.dataframe(
            res["proveedores"], use_container_width=True, hide_index=True,
            column_config={
                **_config_dinero(["NETO", "IVA", "PERCEPCION 1%", "TOTAL A PAGAR", "ABONADO", "SALDO"]),
                "PROMEDIO DIAS PAGO": st.column_config.NumberColumn("PROMEDIO DIAS PAGO", format="%.0f"),
            },
        )

    with t_det:
        f1, f2 = st.columns([2, 2])
        with f1:
            estados = st.multiselect("Filtrar por estado:", sorted(det["ESTADO"].unique()), key="cxp_f_estado")
        with f2:
            proveedores = st.multiselect("Filtrar por proveedor:", sorted(det["PROVEEDOR"].unique()), key="cxp_f_prov")
        vista = det.copy()
        if estados:
            vista = vista[vista["ESTADO"].isin(estados)]
        if proveedores:
            vista = vista[vista["PROVEEDOR"].isin(proveedores)]
        vista["FECHA COMPRA"] = _fmt_fecha(vista["FECHA COMPRA"])
        vista["FECHA PAGO"] = _fmt_fecha(vista["FECHA PAGO"])
        st.caption(f"{len(vista)} documentos · Total a pagar ${vista['TOTAL A PAGAR'].sum():,.2f} · Saldo ${vista['SALDO'].sum():,.2f}")
        st.dataframe(vista, use_container_width=True, hide_index=True,
                     column_config=_config_dinero(["NETO", "IVA", "PERCEPCION 1%", "TOTAL A PAGAR", "ABONADO", "SALDO"]))

    with t_abonos:
        st.markdown("**Cada abono del archivo que cruzó con una compra de la unidad**")
        ab = res["abonos"].copy()
        ab["FECHA PAGO"] = _fmt_fecha(ab["FECHA PAGO"])
        st.dataframe(ab, use_container_width=True, hide_index=True, column_config=_config_dinero(["MONTO ABONADO"]))

    resumen_excel = pd.DataFrame([
        ["Unidad", unidad],
        ["Documentos", k["docs"]],
        ["Proveedores", k["proveedores"]],
        ["Neto", round(k["neto"], 2)],
        ["IVA crédito fiscal", round(k["iva"], 2)],
        ["Percepción 1%", round(k["percepcion"], 2)],
        ["Total a pagar", round(k["total_pagar"], 2)],
        ["Pagado", round(k["abonado"], 2)],
        ["Saldo pendiente", round(k["pendiente"], 2)],
    ] + [[f"Documentos {e}", n] for e, n in pe.items()], columns=["CONCEPTO", "VALOR"])

    nombre_archivo = f"Validacion_CXP_{_norm(unidad).title()}.xlsx"
    st.download_button(
        "📥 Descargar Validación CXP (Excel)",
        data=_generar_excel({
            "Resumen": resumen_excel,
            "Cortes por mes": res["cortes"],
            "Pagos semanales": res["semanas"],
            "Por proveedor": res["proveedores"],
            "Detalle": res["detalle"],
            "Abonos cruzados": res["abonos"],
        }),
        file_name=nombre_archivo,
        use_container_width=True,
        key="cxp_descarga",
    )


def mostrar_tab_cxp():
    """Pestaña 7 del módulo CONTABILIDAD: Validación de Cuentas por Pagar (CXP)."""
    st.subheader("🧾 Validación de Cuentas por Pagar")
    st.caption("Compara las compras ingresadas de una unidad contra los abonos realizados a proveedores. Solo consulta: no guarda nada en la base.")

    unidad = st.selectbox("1. Unidad a validar:", UNIDADES_CXP, key="cxp_unidad")
    c1, c2 = st.columns(2)
    with c1:
        arch_compras = st.file_uploader("2. Reporte de compras de la unidad (Rpt Compras Producto Proveedor):",
                                        type=["xlsx", "xls", "csv"], key="cxp_up_compras")
    with c2:
        arch_abonos = st.file_uploader("3. Reporte de abonos completo, todas las unidades (Rpt List Abonos):",
                                       type=["xlsx", "xls", "csv"], key="cxp_up_abonos")

    if st.button("🔄 Validar Cuentas por Pagar", type="primary", use_container_width=True, key="cxp_btn"):
        if not arch_compras or not arch_abonos:
            st.warning("⚠️ Sube los dos archivos: compras de la unidad y abonos.")
        else:
            with st.spinner("Cruzando compras contra abonos..."):
                try:
                    st.session_state["cxp_resultado"] = procesar_cxp(arch_compras, arch_abonos)
                    st.session_state["cxp_unidad_proc"] = unidad
                    st.success("✅ Validación completada.")
                except ValueError as e:
                    st.error(f"❌ {e}")
                except Exception as e:
                    st.error(f"🚨 Error técnico al procesar los archivos: {e}")

    if "cxp_resultado" in st.session_state:
        st.divider()
        _mostrar_resultados(st.session_state["cxp_resultado"], st.session_state.get("cxp_unidad_proc", unidad))