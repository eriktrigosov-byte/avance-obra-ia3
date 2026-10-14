"""Panel web de seguimiento de avance de obra.

Ejecutar desde la raíz del proyecto:
    streamlit run app.py
"""

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st


RAIZ = Path(__file__).resolve().parent
ARCHIVO_EXCEL = RAIZ / "salidas" / "datos_sinteticos_avance_obra.xlsx"

MORADO = "#A855F7"
MORADO_CLARO = "#D8B4FE"
MAGENTA = "#E879F9"
FONDO = "#090611"
FONDO_EJES = "#110A20"
TEXTO = "#F5F3FF"
REJILLA = "#3B2754"


st.set_page_config(page_title="Avance de obra", page_icon="🏗️", layout="wide")


def aplicar_estilo() -> None:
    """Define el aspecto oscuro y morado del panel."""
    st.markdown(
        f"""
        <style>
          .stApp {{ background: {FONDO}; color: {TEXTO}; }}
          [data-testid="stSidebar"] {{ background: #110A20; }}
          [data-testid="stMetric"] {{
            background: {FONDO_EJES}; border: 1px solid {REJILLA};
            border-radius: 10px; padding: 14px;
          }}
          [data-testid="stMetricValue"] {{ color: {MORADO_CLARO}; }}
          h1, h2, h3, p, label {{ color: {TEXTO}; }}
          .stTabs [data-baseweb="tab"] {{ color: {MORADO_CLARO}; }}
          .stTabs [aria-selected="true"] {{ color: {MAGENTA}; border-bottom-color: {MAGENTA}; }}
        </style>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data
def cargar_datos(ruta: str) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Lee las hojas de datos producidas por generar_avance_obra.py."""
    return (
        pd.read_excel(ruta, sheet_name="Curva S", header=2),
        pd.read_excel(ruta, sheet_name="Avance partidas", header=2),
        pd.read_excel(ruta, sheet_name="Resumen", header=2),
    )


def formatear_metrica(resumen: pd.DataFrame, indicador: str) -> str:
    fila = resumen.loc[resumen["Indicador"] == indicador].iloc[0]
    valor, unidad = fila["Valor"], fila["Unidad"]
    if unidad == "PEN":
        return f"S/ {valor:,.0f}"
    if unidad == "%":
        return f"{valor:.1f}%"
    if unidad == "pp":
        return f"{valor:+.1f} pp"
    return f"{int(valor)}"


def configurar_ejes(ax: plt.Axes) -> None:
    ax.set_facecolor(FONDO_EJES)
    ax.tick_params(colors=TEXTO)
    ax.xaxis.label.set_color(TEXTO)
    ax.yaxis.label.set_color(TEXTO)
    for borde in ax.spines.values():
        borde.set_color(REJILLA)
    ax.grid(axis="y", color=REJILLA, alpha=0.65)
    ax.set_axisbelow(True)


def grafico_curva_s(curva_s: pd.DataFrame) -> plt.Figure:
    fig, ax = plt.subplots(figsize=(12, 5.5), facecolor=FONDO)
    ax.plot(
        curva_s["Periodo"],
        curva_s["Avance planificado acumulado (%)"],
        color=MORADO_CLARO,
        linestyle="--",
        linewidth=2.7,
        marker="o",
        label="Planificado acumulado",
    )
    ax.plot(
        curva_s["Periodo"],
        curva_s["Avance real acumulado (%)"],
        color=MAGENTA,
        linewidth=3,
        marker="o",
        label="Real acumulado",
    )
    configurar_ejes(ax)
    ax.set_title("Curva S: avance acumulado de obra", color=TEXTO, loc="left", fontweight="bold", fontsize=15)
    ax.set_xlabel("Mes")
    ax.set_ylabel("Avance acumulado (%)")
    ax.set_ylim(0, 105)
    ax.tick_params(axis="x", rotation=40)
    leyenda = ax.legend(facecolor=FONDO_EJES, edgecolor=REJILLA, labelcolor=TEXTO, loc="upper left")
    for texto in leyenda.get_texts():
        texto.set_color(TEXTO)
    fig.tight_layout()
    return fig


def grafico_partidas(partidas: pd.DataFrame) -> plt.Figure:
    partidas_ordenadas = partidas.sort_values("Avance real (%)", ascending=True)
    fig, ax = plt.subplots(figsize=(12, 6), facecolor=FONDO)
    y = np.arange(len(partidas_ordenadas))
    alto = 0.36
    ax.barh(y - alto / 2, partidas_ordenadas["Avance planificado (%)"], alto, color=MORADO_CLARO, label="Planificado")
    ax.barh(y + alto / 2, partidas_ordenadas["Avance real (%)"], alto, color=MORADO, label="Real")
    configurar_ejes(ax)
    ax.set_title("Avance de obra por partidas", color=TEXTO, loc="left", fontweight="bold", fontsize=15)
    ax.set_xlabel("Avance (%)")
    ax.set_yticks(y, partidas_ordenadas["Partida"])
    ax.set_xlim(0, 110)
    for indice, (_, fila) in enumerate(partidas_ordenadas.iterrows()):
        ax.text(fila["Avance real (%)"] + 1.5, indice + alto / 2, f'{fila["Avance real (%)"]:.0f}%', color=TEXTO, va="center")
    leyenda = ax.legend(facecolor=FONDO_EJES, edgecolor=REJILLA, labelcolor=TEXTO, loc="lower right")
    for texto in leyenda.get_texts():
        texto.set_color(TEXTO)
    fig.tight_layout()
    return fig


def main() -> None:
    aplicar_estilo()
    st.title("Panel de avance de obra")
    st.caption("Datos sintéticos cargados desde el libro Excel generado por el proyecto.")

    if not ARCHIVO_EXCEL.exists():
        st.error("No se encontró el Excel. Ejecuta primero: python generar_avance_obra.py --recrear")
        st.stop()

    curva_s, partidas, resumen = cargar_datos(str(ARCHIVO_EXCEL))

    with st.sidebar:
        st.header("Fuente de datos")
        st.write("Libro activo")
        st.code(ARCHIVO_EXCEL.name)
        if st.button("Recargar datos"):
            st.cache_data.clear()
            st.rerun()
        st.divider()
        st.caption("Corte de información: noviembre de 2026")

    columnas = st.columns(4)
    metricas = [
        "Presupuesto contractual",
        "Avance físico planificado al corte",
        "Avance físico real al corte",
        "Partidas en retraso",
    ]
    for columna, indicador in zip(columnas, metricas):
        columna.metric(indicador, formatear_metrica(resumen, indicador))

    pestaña_resumen, pestaña_curva, pestaña_partidas = st.tabs(["Resumen", "Curva S", "Partidas"])

    with pestaña_resumen:
        st.subheader("Indicadores al corte")
        resumen_mostrar = resumen.copy()
        resumen_mostrar["Valor"] = resumen_mostrar.apply(
            lambda fila: formatear_metrica(resumen, fila["Indicador"]), axis=1
        )
        st.dataframe(resumen_mostrar, hide_index=True, use_container_width=True)
        st.info("El avance físico real ponderado es 6.7 puntos porcentuales menor que el planificado al corte.")

    with pestaña_curva:
        st.pyplot(grafico_curva_s(curva_s), use_container_width=True)
        columnas_curva = [
            "Periodo",
            "Avance planificado mensual (%)",
            "Avance real mensual (%)",
            "Avance planificado acumulado (%)",
            "Avance real acumulado (%)",
        ]
        st.dataframe(curva_s[columnas_curva], hide_index=True, use_container_width=True)

    with pestaña_partidas:
        st.pyplot(grafico_partidas(partidas), use_container_width=True)
        st.dataframe(partidas, hide_index=True, use_container_width=True)


if __name__ == "__main__":
    main()
