"""Genera datos sintéticos de una obra y sus gráficos de avance.

Ejecutar:
    python generar_avance_obra.py --recrear
"""

from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import matplotlib.pyplot as plt


RAIZ = Path(__file__).resolve().parent
CARPETA_SALIDA = RAIZ / "salidas"
ARCHIVO_EXCEL = CARPETA_SALIDA / "datos_sinteticos_avance_obra.xlsx"
GRAFICO_CURVA_S = CARPETA_SALIDA / "curva_s_avance_obra.png"
GRAFICO_PARTIDAS = CARPETA_SALIDA / "avance_por_partidas.png"

MONEDA = "PEN"
PRESUPUESTO_TOTAL = 12_500_000

MORADO = "#A855F7"
MORADO_CLARO = "#D8B4FE"
MORADO_INTENSO = "#7E22CE"
MAGENTA = "#E879F9"
FONDO = "#090611"
FONDO_EJES = "#110A20"
TEXTO = "#F5F3FF"
REJILLA = "#3B2754"


def crear_datos_sinteticos() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Crea una curva S y avances por partida físicamente consistentes."""
    meses = pd.date_range("2026-01-01", periods=12, freq="MS")

    # Distribución mensual de valor planificado: acelera en la etapa central y
    # desacelera al cierre, por lo que genera una curva S típica.
    porcentaje_plan_mensual = np.array(
        [1.5, 3.5, 6.5, 10.0, 13.0, 14.5, 14.0, 12.5, 10.0, 7.5, 4.5, 2.5]
    )
    porcentaje_real_mensual = np.array(
        [1.0, 2.8, 5.6, 8.8, 11.9, 13.8, 14.7, 13.6, 11.6, 9.0, 5.4, 2.8]
    )

    plan_acum = np.cumsum(porcentaje_plan_mensual)
    real_acum = np.cumsum(porcentaje_real_mensual)
    costo_plan_mensual = PRESUPUESTO_TOTAL * porcentaje_plan_mensual / 100
    # En la realidad se incorpora una desviación de costo progresiva de 2.1%.
    costo_real_mensual = (
        PRESUPUESTO_TOTAL
        * porcentaje_real_mensual
        / 100
        * np.linspace(0.995, 1.021, len(meses))
    )

    curva_s = pd.DataFrame(
        {
            "Mes": meses,
            "Periodo": meses.strftime("%b-%Y"),
            "Avance planificado mensual (%)": porcentaje_plan_mensual,
            "Avance real mensual (%)": porcentaje_real_mensual,
            "Avance planificado acumulado (%)": plan_acum,
            "Avance real acumulado (%)": real_acum,
            "Costo planificado mensual (PEN)": costo_plan_mensual.round(2),
            "Costo real mensual (PEN)": costo_real_mensual.round(2),
            "Costo planificado acumulado (PEN)": np.cumsum(costo_plan_mensual).round(2),
            "Costo real acumulado (PEN)": np.cumsum(costo_real_mensual).round(2),
        }
    )

    partidas = pd.DataFrame(
        {
            "Código": ["01", "02", "03", "04", "05", "06", "07", "08", "09"],
            "Partida": [
                "Obras preliminares",
                "Movimiento de tierras",
                "Cimentación",
                "Estructura de concreto",
                "Albañilería",
                "Instalaciones eléctricas",
                "Instalaciones sanitarias",
                "Acabados",
                "Obras exteriores",
            ],
            "Peso presupuestal (%)": [5, 8, 16, 25, 13, 9, 7, 14, 3],
            "Inicio planificado": pd.to_datetime(
                [
                    "2026-01-01", "2026-01-01", "2026-02-01", "2026-03-01",
                    "2026-05-01", "2026-06-01", "2026-06-01", "2026-08-01",
                    "2026-09-01",
                ]
            ),
            "Fin planificado": pd.to_datetime(
                [
                    "2026-02-28", "2026-03-31", "2026-05-31", "2026-08-31",
                    "2026-09-30", "2026-10-31", "2026-10-31", "2026-12-15",
                    "2026-12-31",
                ]
            ),
            "Avance planificado (%)": [100, 100, 100, 100, 100, 95, 95, 75, 55],
            "Avance real (%)": [100, 100, 100, 98, 93, 84, 86, 57, 38],
        }
    )
    partidas["Presupuesto (PEN)"] = (
        PRESUPUESTO_TOTAL * partidas["Peso presupuestal (%)"] / 100
    ).round(2)
    partidas["Desviación de avance (pp)"] = (
        partidas["Avance real (%)"] - partidas["Avance planificado (%)"]
    )
    partidas["Estado"] = np.select(
        [
            partidas["Avance real (%)"] >= partidas["Avance planificado (%)"],
            partidas["Desviación de avance (pp)"] >= -5,
        ],
        ["En plazo", "Atención"],
        default="Retraso",
    )

    avance_fisico_ponderado = (
        partidas["Peso presupuestal (%)"] * partidas["Avance real (%)"] / 100
    ).sum()
    resumen = pd.DataFrame(
        {
            "Indicador": [
                "Presupuesto contractual",
                "Avance físico planificado al corte",
                "Avance físico real al corte",
                "Desviación de avance",
                "Costo real acumulado",
                "Desviación de costo acumulado",
                "Partidas en retraso",
            ],
            "Valor": [
                PRESUPUESTO_TOTAL,
                95.0,
                round(avance_fisico_ponderado, 1),
                round(avance_fisico_ponderado - 95.0, 1),
                float(curva_s.loc[10, "Costo real acumulado (PEN)"].round(2)),
                float(
                    (
                        curva_s.loc[10, "Costo real acumulado (PEN)"]
                        - curva_s.loc[10, "Costo planificado acumulado (PEN)"]
                    ).round(2)
                ),
                int((partidas["Estado"] == "Retraso").sum()),
            ],
            "Unidad": ["PEN", "%", "%", "pp", "PEN", "PEN", "partidas"],
        }
    )
    return curva_s, partidas, resumen


def escribir_excel(
    curva_s: pd.DataFrame, partidas: pd.DataFrame, resumen: pd.DataFrame
) -> None:
    """Guarda las tres tablas como un libro Excel con formato legible."""
    CARPETA_SALIDA.mkdir(exist_ok=True)
    libro = Workbook()
    libro.remove(libro.active)
    estilo_cabecera = PatternFill("solid", fgColor="5B21B6")
    estilo_titulo = PatternFill("solid", fgColor="2E1065")

    for nombre, datos in {
        "Curva S": curva_s,
        "Avance partidas": partidas,
        "Resumen": resumen,
    }.items():
        hoja = libro.create_sheet(nombre)
        hoja.sheet_view.showGridLines = False
        hoja.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(datos.columns))
        celda_titulo = hoja.cell(1, 1, f"Datos sintéticos - {nombre}")
        celda_titulo.fill = estilo_titulo
        celda_titulo.font = Font(color="FFFFFF", bold=True, size=14)
        celda_titulo.alignment = Alignment(horizontal="left")

        for col, etiqueta in enumerate(datos.columns, 1):
            celda = hoja.cell(3, col, etiqueta)
            celda.fill = estilo_cabecera
            celda.font = Font(color="FFFFFF", bold=True)
            celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for fila, registro in enumerate(datos.itertuples(index=False), 4):
            for col, valor in enumerate(registro, 1):
                celda = hoja.cell(fila, col, valor.to_pydatetime() if isinstance(valor, pd.Timestamp) else valor)
                celda.alignment = Alignment(vertical="center")
                encabezado = str(datos.columns[col - 1])
                if "(%)" in encabezado:
                    celda.number_format = '0.0"%"'
                elif "(PEN)" in encabezado or encabezado == "Valor" and nombre == "Resumen":
                    celda.number_format = '#,##0.00'
                elif "Inicio" in encabezado or "Fin" in encabezado or encabezado == "Mes":
                    celda.number_format = "mmm-yy"
                elif encabezado == "Código":
                    celda.number_format = "@"

        hoja.freeze_panes = "A4"
        hoja.auto_filter.ref = f"A3:{get_column_letter(len(datos.columns))}{len(datos) + 3}"
        for col in range(1, len(datos.columns) + 1):
            ancho = max(14, min(34, max(len(str(x)) for x in [datos.columns[col - 1], *datos.iloc[:, col - 1].astype(str)]) + 2))
            hoja.column_dimensions[get_column_letter(col)].width = ancho
        hoja.row_dimensions[1].height = 24
        hoja.row_dimensions[3].height = 34

    libro.save(ARCHIVO_EXCEL)


def configurar_ejes(ax: plt.Axes) -> None:
    """Aplica el lenguaje visual morado sobre negro solicitado."""
    ax.set_facecolor(FONDO_EJES)
    ax.tick_params(colors=TEXTO, labelsize=10)
    ax.xaxis.label.set_color(TEXTO)
    ax.yaxis.label.set_color(TEXTO)
    ax.title.set_color(TEXTO)
    for borde in ax.spines.values():
        borde.set_color(REJILLA)
    ax.grid(axis="y", color=REJILLA, alpha=0.65, linewidth=0.8)
    ax.set_axisbelow(True)


def generar_graficos_desde_excel() -> None:
    """Lee el Excel generado y crea los dos gráficos solicitados."""
    curva_s = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Curva S", header=2)
    partidas = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Avance partidas", header=2)

    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10})
    fig, ax = plt.subplots(figsize=(12, 6.5), facecolor=FONDO)
    ax.plot(
        curva_s["Periodo"],
        curva_s["Avance planificado acumulado (%)"],
        color=MORADO_CLARO,
        linestyle="--",
        linewidth=2.8,
        marker="o",
        markersize=5,
        label="Planificado acumulado",
    )
    ax.plot(
        curva_s["Periodo"],
        curva_s["Avance real acumulado (%)"],
        color=MAGENTA,
        linewidth=3.2,
        marker="o",
        markersize=5,
        label="Real acumulado",
    )
    configurar_ejes(ax)
    ax.set_title(
        "Curva S: avance acumulado de obra",
        loc="left",
        fontweight="bold",
        fontsize=16,
        pad=16,
        color=TEXTO,
    )
    ax.set_xlabel("Mes")
    ax.set_ylabel("Avance acumulado (%)")
    ax.set_ylim(0, 105)
    ax.set_yticks(np.arange(0, 101, 10))
    ax.tick_params(axis="x", rotation=40)
    leyenda = ax.legend(facecolor=FONDO_EJES, edgecolor=REJILLA, labelcolor=TEXTO, loc="upper left")
    for texto in leyenda.get_texts():
        texto.set_color(TEXTO)
    fig.tight_layout()
    fig.savefig(GRAFICO_CURVA_S, dpi=180, facecolor=FONDO, bbox_inches="tight")
    plt.close(fig)

    partidas = partidas.sort_values("Avance real (%)", ascending=True)
    fig, ax = plt.subplots(figsize=(12, 7), facecolor=FONDO)
    y = np.arange(len(partidas))
    alto = 0.36
    ax.barh(y - alto / 2, partidas["Avance planificado (%)"], alto, color=MORADO_CLARO, label="Planificado")
    ax.barh(y + alto / 2, partidas["Avance real (%)"], alto, color=MORADO, label="Real")
    configurar_ejes(ax)
    ax.set_title(
        "Avance de obra por partidas",
        loc="left",
        fontweight="bold",
        fontsize=16,
        pad=16,
        color=TEXTO,
    )
    ax.set_xlabel("Avance (%)")
    ax.set_yticks(y, partidas["Partida"])
    ax.set_xlim(0, 110)
    ax.set_xticks(np.arange(0, 101, 10))
    for indice, (_, fila) in enumerate(partidas.iterrows()):
        ax.text(fila["Avance real (%)"] + 1.5, indice + alto / 2, f'{fila["Avance real (%)"]:.0f}%', color=TEXTO, va="center", fontsize=9)
    leyenda = ax.legend(facecolor=FONDO_EJES, edgecolor=REJILLA, labelcolor=TEXTO, loc="lower right")
    for texto in leyenda.get_texts():
        texto.set_color(TEXTO)
    fig.tight_layout()
    fig.savefig(GRAFICO_PARTIDAS, dpi=180, facecolor=FONDO, bbox_inches="tight")
    plt.close(fig)


def main() -> None:
    parser = argparse.ArgumentParser(description="Curva S y avance por partidas de una obra sintética.")
    parser.add_argument("--recrear", action="store_true", help="Regenera el Excel sintético antes de graficar.")
    args = parser.parse_args()

    if args.recrear or not ARCHIVO_EXCEL.exists():
        escribir_excel(*crear_datos_sinteticos())
    generar_graficos_desde_excel()
    print("Archivos generados:")
    for archivo in (ARCHIVO_EXCEL, GRAFICO_CURVA_S, GRAFICO_PARTIDAS):
        print(f"- {archivo.relative_to(RAIZ)}")


if __name__ == "__main__":
    main()
