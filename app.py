import io
import re
import unicodedata
from pathlib import Path

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


def conciliar(facturas, movimientos, tolerancia=TOLERANCIA):
    """
    Conciliación didáctica basada en reglas transparentes.

    Prioridad de clasificación por factura:
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

        if pd.isna(total):
            resultados.append(
                {
                    "Folio": folio,
                    "Cliente": factura["Cliente"],
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

        # 1. Duplicado: más de un movimiento con la misma referencia y el mismo importe esperado.
        exactos_ref = por_referencia[
            (por_referencia["Importe"] - total).abs() <= tolerancia
        ]
        if len(exactos_ref) >= 2:
            ids = exactos_ref["ID_Movimiento"].astype(str).tolist()
            movimientos_asignados.update(ids)
            importe_banco = float(exactos_ref["Importe"].sum())
            resultados.append(
                {
                    "Folio": folio,
                    "Cliente": factura["Cliente"],
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
            resultados.append(
                {
                    "Folio": folio,
                    "Cliente": factura["Cliente"],
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
            # Se toma el movimiento más cercano al importe de la factura.
            idx = (por_referencia["Importe"] - total).abs().idxmin()
            mov = por_referencia.loc[idx]
            importe = float(mov["Importe"])
            movimientos_asignados.add(str(mov["ID_Movimiento"]))

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
                    "Total_Factura": total,
                    "Estatus": estatus,
                    "Movimientos_Relacionados": str(mov["ID_Movimiento"]),
                    "Importe_Banco": importe,
                    "Diferencia": round(total - importe, 2),
                    "Comentario": comentario,
                }
            )
            continue

        # 5. Posible coincidencia: mismo importe y el nombre del cliente aparece en el concepto.
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
            resultados.append(
                {
                    "Folio": folio,
                    "Cliente": factura["Cliente"],
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


def construir_resumen(resultado_facturas, no_asignados):
    total_facturas = len(resultado_facturas)
    conciliadas = int((resultado_facturas["Estatus"] == "Conciliada").sum())
    porcentaje = (conciliadas / total_facturas * 100) if total_facturas else 0.0

    excepciones = int((resultado_facturas["Estatus"] != "Conciliada").sum()) + len(no_asignados)

    importe_pendiente = resultado_facturas.loc[
        resultado_facturas["Diferencia"] > 0, "Diferencia"
    ].sum()

    return {
        "Facturas revisadas": total_facturas,
        "Facturas conciliadas": conciliadas,
        "% conciliado": porcentaje,
        "Excepciones": excepciones,
        "Importe pendiente": float(importe_pendiente),
        "Depósitos sin factura": len(no_asignados),
    }


def generar_excel(resultado_facturas, no_asignados, resumen):
    buffer = io.BytesIO()
    resumen_df = pd.DataFrame(
        {"Indicador": list(resumen.keys()), "Valor": list(resumen.values())}
    )

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        resumen_df.to_excel(writer, index=False, sheet_name="Resumen")
        resultado_facturas.to_excel(writer, index=False, sheet_name="Conciliacion")
        no_asignados.to_excel(writer, index=False, sheet_name="Mov_No_Conciliados")

    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------
# Interfaz Streamlit
# ---------------------------------------------------------
st.title("Conciliación automática: Facturación vs. Bancos")
st.caption(
    "Modelo didáctico en Python para comparar facturas y movimientos bancarios, "
    "identificar coincidencias y generar excepciones para revisión contable."
)

with st.expander("¿Qué hace el modelo?", expanded=False):
    st.markdown(
        """
        - Lee dos archivos Excel: **Facturas** y **Movimientos Bancarios**.
        - Estandariza referencias, importes y textos.
        - Identifica conciliaciones exactas.
        - Detecta pagos parciales, diferencias, duplicados y posibles coincidencias.
        - Señala facturas sin cobro y depósitos sin factura.
        - Genera indicadores y un archivo Excel descargable con los resultados.
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
        resumen = construir_resumen(resultado_facturas, no_asignados)

        st.success("Conciliación ejecutada correctamente.")

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Facturas", resumen["Facturas revisadas"])
        k2.metric("Conciliadas", resumen["Facturas conciliadas"])
        k3.metric("% conciliado", f"{resumen['% conciliado']:.1f}%")
        k4.metric("Excepciones", resumen["Excepciones"])
        k5.metric("Importe pendiente", f"${resumen['Importe pendiente']:,.2f}")

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

        # Distribución de estatus: usa capacidades analíticas de pandas + visualización Streamlit.
        st.subheader("Distribución de resultados")
        distribucion = (
            resultado_facturas["Estatus"]
            .value_counts()
            .rename_axis("Estatus")
            .reset_index(name="Número de facturas")
        )
        st.bar_chart(distribucion.set_index("Estatus"))

        reporte = generar_excel(resultado_facturas, no_asignados, resumen)
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
