import io
import re
import unicodedata

import pandas as pd
import streamlit as st


# ---------------------------------------------------------
# Configuración
# ---------------------------------------------------------
st.set_page_config(
    page_title="Conciliación Facturación vs. Bancos",
    page_icon="📊",
    layout="wide",
)

TOLERANCIA = 0.01

FACTURAS_REQUERIDAS = {
    "Folio",
    "Fecha_Factura",
    "RFC_Cliente",
    "Cliente",
    "Total",
    "Referencia_Esperada",
}

MOVIMIENTOS_REQUERIDOS = {
    "ID_Movimiento",
    "Fecha_Movimiento",
    "Cuenta_Bancaria",
    "Referencia",
    "Concepto",
    "Importe",
}


def normalizar_texto(valor):
    """Convierte texto a una forma comparable: mayúsculas, sin acentos y sin espacios extra."""
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"\s+", " ", texto)
    return texto


def leer_excel(archivo):
    return pd.read_excel(archivo, engine="openpyxl")


def validar_columnas(df, requeridas, nombre_base):
    faltantes = sorted(requeridas - set(df.columns))
    if faltantes:
        raise ValueError(
            f"La base '{nombre_base}' no contiene las columnas requeridas: "
            + ", ".join(faltantes)
        )


def preparar_bases(facturas, movimientos):
    facturas = facturas.copy()
    movimientos = movimientos.copy()

    facturas["Total"] = pd.to_numeric(facturas["Total"], errors="coerce")
    movimientos["Importe"] = pd.to_numeric(movimientos["Importe"], errors="coerce")

    facturas["Fecha_Factura"] = pd.to_datetime(facturas["Fecha_Factura"], errors="coerce")
    movimientos["Fecha_Movimiento"] = pd.to_datetime(
        movimientos["Fecha_Movimiento"], errors="coerce"
    )

    facturas["ref_norm"] = facturas["Referencia_Esperada"].map(normalizar_texto)
    facturas["cliente_norm"] = facturas["Cliente"].map(normalizar_texto)

    movimientos["ref_norm"] = movimientos["Referencia"].map(normalizar_texto)
    movimientos["concepto_norm"] = movimientos["Concepto"].map(normalizar_texto)

    return facturas, movimientos


def dias_entre(fecha_factura, fecha_cobro):
    if pd.isna(fecha_factura) or pd.isna(fecha_cobro):
        return None
    return int((fecha_cobro - fecha_factura).days)


def conciliar(facturas, movimientos, tolerancia=TOLERANCIA):
    """
    Conciliación didáctica basada en reglas transparentes.

    Prioridad por factura:
    1) Pago duplicado
    2) Conciliada
    3) Pago parcial
    4) Diferencia de importe
    5) Posible coincidencia por cliente + importe
    6) Factura sin cobro
    """
    facturas, movimientos = preparar_bases(facturas, movimientos)

    resultados = []
    movimientos_asignados = set()

    for _, factura in facturas.iterrows():
        folio = factura["Folio"]
        total = factura["Total"]
        ref = factura["ref_norm"]
        cliente = factura["cliente_norm"]
        fecha_factura = factura["Fecha_Factura"]

        if pd.isna(total):
            resultados.append(
                {
                    "Folio": folio,
                    "Cliente": factura["Cliente"],
                    "Fecha_Factura": fecha_factura,
                    "Fecha_Cobro": pd.NaT,
                    "Dias_Cobro": None,
                    "Total_Factura": total,
                    "Estatus": "Dato inválido",
                    "Movimientos_Relacionados": "",
                    "Importe_Banco": 0.0,
                    "Diferencia": None,
                    "Comentario": "El total de la factura no es numérico.",
                }
            )
            continue

        por_referencia = movimientos[movimientos["ref_norm"] == ref]

        # 1. Duplicado: dos o más abonos exactos para la misma referencia.
        exactos_ref = por_referencia[
            (por_referencia["Importe"] - total).abs() <= tolerancia
        ]
        if len(exactos_ref) >= 2:
            ids = exactos_ref["ID_Movimiento"].astype(str).tolist()
            movimientos_asignados.update(ids)
            importe_banco = float(exactos_ref["Importe"].sum())
            fecha_cobro = exactos_ref["Fecha_Movimiento"].min()
            resultados.append(
                {
                    "Folio": folio,
                    "Cliente": factura["Cliente"],
                    "Fecha_Factura": fecha_factura,
                    "Fecha_Cobro": fecha_cobro,
                    "Dias_Cobro": dias_entre(fecha_factura, fecha_cobro),
                    "Total_Factura": total,
                    "Estatus": "Pago duplicado",
                    "Movimientos_Relacionados": ", ".join(ids),
                    "Importe_Banco": importe_banco,
                    "Diferencia": round(total - importe_banco, 2),
                    "Comentario": "Se localizaron dos o más abonos exactos para la misma referencia.",
                }
            )
            continue

        # 2. Conciliación exacta.
        if len(exactos_ref) == 1:
            mov = exactos_ref.iloc[0]
            movimientos_asignados.add(str(mov["ID_Movimiento"]))
            fecha_cobro = mov["Fecha_Movimiento"]
            resultados.append(
                {
                    "Folio": folio,
                    "Cliente": factura["Cliente"],
                    "Fecha_Factura": fecha_factura,
                    "Fecha_Cobro": fecha_cobro,
                    "Dias_Cobro": dias_entre(fecha_factura, fecha_cobro),
                    "Total_Factura": total,
                    "Estatus": "Conciliada",
                    "Movimientos_Relacionados": str(mov["ID_Movimiento"]),
                    "Importe_Banco": float(mov["Importe"]),
                    "Diferencia": 0.0,
                    "Comentario": "Referencia e importe coinciden.",
                }
            )
            continue

        # 3 y 4. Existe referencia, pero no coincide el importe.
        if not por_referencia.empty:
            idx = (por_referencia["Importe"] - total).abs().idxmin()
            mov = por_referencia.loc[idx]
            importe = float(mov["Importe"])
            movimientos_asignados.add(str(mov["ID_Movimiento"]))
            fecha_cobro = mov["Fecha_Movimiento"]

            estatus = "Pago parcial" if importe < total else "Diferencia de importe"
            comentario = (
                "La referencia coincide, pero el abono es menor al total facturado."
                if importe < total
                else "La referencia coincide, pero el importe del abono es distinto al total facturado."
            )

            resultados.append(
                {
                    "Folio": folio,
                    "Cliente": factura["Cliente"],
                    "Fecha_Factura": fecha_factura,
                    "Fecha_Cobro": fecha_cobro,
                    "Dias_Cobro": dias_entre(fecha_factura, fecha_cobro),
                    "Total_Factura": total,
                    "Estatus": estatus,
                    "Movimientos_Relacionados": str(mov["ID_Movimiento"]),
                    "Importe_Banco": importe,
                    "Diferencia": round(total - importe, 2),
                    "Comentario": comentario,
                }
            )
            continue

        # 5. Posible coincidencia: mismo importe y cliente presente en el concepto.
        candidatos = movimientos[
            ((movimientos["Importe"] - total).abs() <= tolerancia)
            & movimientos["concepto_norm"].str.contains(re.escape(cliente), na=False)
        ]
        candidatos = candidatos[
            ~candidatos["ID_Movimiento"].astype(str).isin(movimientos_asignados)
        ]

        if not candidatos.empty:
            mov = candidatos.iloc[0]
            movimientos_asignados.add(str(mov["ID_Movimiento"]))
            fecha_cobro = mov["Fecha_Movimiento"]
            resultados.append(
                {
                    "Folio": folio,
                    "Cliente": factura["Cliente"],
                    "Fecha_Factura": fecha_factura,
                    "Fecha_Cobro": fecha_cobro,
                    "Dias_Cobro": dias_entre(fecha_factura, fecha_cobro),
                    "Total_Factura": total,
                    "Estatus": "Posible coincidencia",
                    "Movimientos_Relacionados": str(mov["ID_Movimiento"]),
                    "Importe_Banco": float(mov["Importe"]),
                    "Diferencia": round(total - float(mov["Importe"]), 2),
                    "Comentario": "Importe y cliente son consistentes, pero la referencia no coincide.",
                }
            )
            continue

        # 6. Sin cobro identificable.
        resultados.append(
            {
                "Folio": folio,
                "Cliente": factura["Cliente"],
                "Fecha_Factura": fecha_factura,
                "Fecha_Cobro": pd.NaT,
                "Dias_Cobro": None,
                "Total_Factura": total,
                "Estatus": "Factura sin cobro",
                "Movimientos_Relacionados": "",
                "Importe_Banco": 0.0,
                "Diferencia": round(total, 2),
                "Comentario": "No se localizó un movimiento bancario asociado.",
            }
        )

    resultado_facturas = pd.DataFrame(resultados)

    # Movimientos no asignados a ninguna factura.
    no_asignados = movimientos[
        ~movimientos["ID_Movimiento"].astype(str).isin(movimientos_asignados)
    ].copy()

    columnas_salida = [
        "ID_Movimiento",
        "Fecha_Movimiento",
        "Cuenta_Bancaria",
        "Referencia",
        "Concepto",
        "Importe",
    ]
    no_asignados = no_asignados[columnas_salida]
    no_asignados["Estatus"] = "Depósito sin factura"
    no_asignados["Comentario"] = "El movimiento no pudo asociarse a una factura."

    return resultado_facturas, no_asignados


def construir_indicadores(resultado_facturas, no_asignados):
    total_facturas = len(resultado_facturas)
    conciliadas_mask = resultado_facturas["Estatus"] == "Conciliada"
    conciliadas = int(conciliadas_mask.sum())

    # 1. % de facturas conciliadas
    pct_facturas_conciliadas = (
        conciliadas / total_facturas * 100 if total_facturas else 0.0
    )

    # 2. % del monto facturado conciliado
    monto_facturado = pd.to_numeric(
        resultado_facturas["Total_Factura"], errors="coerce"
    ).fillna(0).sum()
    monto_conciliado = pd.to_numeric(
        resultado_facturas.loc[conciliadas_mask, "Total_Factura"], errors="coerce"
    ).fillna(0).sum()
    pct_monto_conciliado = (
        monto_conciliado / monto_facturado * 100 if monto_facturado else 0.0
    )

    # 3. Monto pendiente de cobro: sólo saldos positivos por cobrar.
    diferencias = pd.to_numeric(resultado_facturas["Diferencia"], errors="coerce").fillna(0)
    importe_pendiente = diferencias.clip(lower=0).sum()

    # 4. Días promedio de cobro: facturas con un movimiento identificado.
    dias_validos = pd.to_numeric(resultado_facturas["Dias_Cobro"], errors="coerce").dropna()
    dias_validos = dias_validos[dias_validos >= 0]
    dias_promedio_cobro = float(dias_validos.mean()) if not dias_validos.empty else 0.0

    # 5. Depósitos no identificados: número y monto.
    depositos_no_identificados = len(no_asignados)
    monto_depositos_no_identificados = pd.to_numeric(
        no_asignados["Importe"], errors="coerce"
    ).fillna(0).sum()

    # Indicadores operativos complementarios.
    excepciones = int((resultado_facturas["Estatus"] != "Conciliada").sum()) + len(no_asignados)

    return {
        "Facturas revisadas": total_facturas,
        "Facturas conciliadas": conciliadas,
        "% facturas conciliadas": float(pct_facturas_conciliadas),
        "Monto total facturado": float(monto_facturado),
        "Monto conciliado": float(monto_conciliado),
        "% monto facturado conciliado": float(pct_monto_conciliado),
        "Monto pendiente de cobro": float(importe_pendiente),
        "Días promedio de cobro": float(dias_promedio_cobro),
        "Depósitos no identificados": int(depositos_no_identificados),
        "Monto depósitos no identificados": float(monto_depositos_no_identificados),
        "Excepciones totales": int(excepciones),
    }


def construir_top_excepciones(resultado_facturas, no_asignados, n=10):
    # Excepciones de facturas.
    exc_facturas = resultado_facturas[
        resultado_facturas["Estatus"] != "Conciliada"
    ].copy()

    exc_facturas["Importe_Impacto"] = pd.to_numeric(
        exc_facturas["Diferencia"], errors="coerce"
    ).abs().fillna(0)

    # Cuando la diferencia es 0 pero aún requiere revisión (p. ej. posible coincidencia),
    # se toma el valor de la factura como monto sujeto a revisión.
    mascara_cero = exc_facturas["Importe_Impacto"] <= TOLERANCIA
    exc_facturas.loc[mascara_cero, "Importe_Impacto"] = pd.to_numeric(
        exc_facturas.loc[mascara_cero, "Total_Factura"], errors="coerce"
    ).fillna(0)

    exc_facturas = exc_facturas.assign(
        Tipo="Factura",
        Identificador=exc_facturas["Folio"].astype(str),
        Detalle=exc_facturas["Cliente"].astype(str),
    )[["Tipo", "Identificador", "Estatus", "Detalle", "Importe_Impacto"]]

    # Excepciones de movimientos bancarios.
    exc_movs = no_asignados.copy()
    exc_movs["Importe_Impacto"] = pd.to_numeric(exc_movs["Importe"], errors="coerce").fillna(0)
    exc_movs = exc_movs.assign(
        Tipo="Movimiento bancario",
        Identificador=exc_movs["ID_Movimiento"].astype(str),
        Detalle=exc_movs["Concepto"].astype(str),
    )[["Tipo", "Identificador", "Estatus", "Detalle", "Importe_Impacto"]]

    top = pd.concat([exc_facturas, exc_movs], ignore_index=True)
    top = top.sort_values("Importe_Impacto", ascending=False).head(n).reset_index(drop=True)
    return top


def generar_excel(resultado_facturas, no_asignados, indicadores, top_excepciones):
    buffer = io.BytesIO()

    indicadores_df = pd.DataFrame(
        {
            "Indicador": list(indicadores.keys()),
            "Valor": list(indicadores.values()),
        }
    )

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        indicadores_df.to_excel(writer, index=False, sheet_name="Indicadores")
        resultado_facturas.to_excel(writer, index=False, sheet_name="Conciliacion")
        no_asignados.to_excel(writer, index=False, sheet_name="Mov_No_Conciliados")
        top_excepciones.to_excel(writer, index=False, sheet_name="Top10_Excepciones")

    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------
# Interfaz Streamlit
# ---------------------------------------------------------
st.title("Conciliación automática: Facturación vs. Bancos")
st.caption(
    "Modelo didáctico en Python para comparar facturas y movimientos bancarios, "
    "identificar coincidencias, medir cobranza y generar excepciones para revisión contable."
)

with st.expander("¿Qué hace el modelo?", expanded=False):
    st.markdown(
        """
        - Lee dos archivos Excel: **Facturas** y **Movimientos Bancarios**.
        - Estandariza referencias, importes y textos.
        - Identifica conciliaciones exactas.
        - Detecta pagos parciales, diferencias, duplicados y posibles coincidencias.
        - Señala facturas sin cobro y depósitos sin factura.
        - Calcula seis indicadores clave de conciliación y cobranza.
        - Genera un **Top 10 de excepciones por importe**.
        - Permite descargar un reporte Excel con resultados e indicadores.
        """
    )

col1, col2 = st.columns(2)

with col1:
    archivo_facturas = st.file_uploader(
        "1. Cargar base de Facturas",
        type=["xlsx"],
        key="facturas",
    )

with col2:
    archivo_movimientos = st.file_uploader(
        "2. Cargar base de Movimientos Bancarios",
        type=["xlsx"],
        key="movimientos",
    )

if archivo_facturas is not None and archivo_movimientos is not None:
    try:
        facturas = leer_excel(archivo_facturas)
        movimientos = leer_excel(archivo_movimientos)

        validar_columnas(facturas, FACTURAS_REQUERIDAS, "Facturas")
        validar_columnas(movimientos, MOVIMIENTOS_REQUERIDOS, "Movimientos Bancarios")

        resultado_facturas, no_asignados = conciliar(facturas, movimientos)
        indicadores = construir_indicadores(resultado_facturas, no_asignados)
        top_excepciones = construir_top_excepciones(resultado_facturas, no_asignados)

        st.success("Conciliación ejecutada correctamente.")

        st.subheader("Indicadores de conciliación y cobranza")

        # Primera fila: 3 indicadores principales.
        k1, k2, k3 = st.columns(3)
        k1.metric(
            "% facturas conciliadas",
            f"{indicadores['% facturas conciliadas']:.1f}%",
        )
        k2.metric(
            "% monto facturado conciliado",
            f"{indicadores['% monto facturado conciliado']:.1f}%",
        )
        k3.metric(
            "Monto pendiente de cobro",
            f"${indicadores['Monto pendiente de cobro']:,.2f}",
        )

        # Segunda fila: 3 indicadores adicionales.
        k4, k5, k6 = st.columns(3)
        k4.metric(
            "Días promedio de cobro",
            f"{indicadores['Días promedio de cobro']:.1f}",
        )
        k5.metric(
            "Depósitos no identificados",
            f"{indicadores['Depósitos no identificados']}",
        )
        k6.metric(
            "Monto depósitos no identificados",
            f"${indicadores['Monto depósitos no identificados']:,.2f}",
        )

        st.subheader("Top 10 excepciones por importe")
        st.dataframe(
            top_excepciones.style.format({"Importe_Impacto": "${:,.2f}"}),
            use_container_width=True,
            hide_index=True,
        )
        st.bar_chart(
            top_excepciones.set_index("Identificador")[["Importe_Impacto"]]
        )

        st.subheader("Resultado por factura")
        st.dataframe(
            resultado_facturas,
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("Movimientos bancarios no conciliados")
        st.dataframe(
            no_asignados,
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("Distribución de resultados")
        distribucion = (
            resultado_facturas["Estatus"]
            .value_counts()
            .rename_axis("Estatus")
            .reset_index(name="Número de facturas")
        )
        st.bar_chart(distribucion.set_index("Estatus"))

        reporte = generar_excel(
            resultado_facturas,
            no_asignados,
            indicadores,
            top_excepciones,
        )
        st.download_button(
            label="Descargar reporte de conciliación",
            data=reporte,
            file_name="Reporte_Conciliacion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as exc:
        st.error(f"No fue posible ejecutar la conciliación: {exc}")
else:
    st.info("Carga ambos archivos Excel para ejecutar el modelo.")

st.divider()
st.caption(
    "Ejercicio didáctico. Las reglas son deliberadamente transparentes para que puedan ser revisadas, "
    "modificadas y ampliadas por los participantes."
)
