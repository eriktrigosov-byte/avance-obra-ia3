# Curva S y avance de obra por partidas

El script `generar_avance_obra.py` construye datos sintéticos para una obra de doce meses con presupuesto contractual de PEN 12,500,000. La información se guarda primero en `salidas/datos_sinteticos_avance_obra.xlsx` y luego se lee desde ese mismo libro para generar los gráficos.

## Datos incluidos

- `Curva S`: avance mensual y acumulado planificado y real, además de costos mensuales y acumulados.
- `Avance partidas`: nueve partidas de obra con peso presupuestal, fechas planificadas, avance físico y estado.
- `Resumen`: indicadores principales al corte de noviembre de 2026.

La programación se concentra en los meses centrales, creando una curva S. El avance real presenta retrasos moderados en instalaciones, acabados y obras exteriores; por ello el avance físico ponderado queda por debajo de lo planificado al corte.

## Instalación y ejecución

En PowerShell, desde la raíz del proyecto:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe generar_avance_obra.py --recrear
```

`--recrear` vuelve a generar el libro Excel y ambos PNG. Sin esa opción, el script conserva el Excel existente y crea los gráficos leyéndolo.

## Gráficos

- `curva_s_avance_obra.png` compara el avance acumulado planificado con el real.
- `avance_por_partidas.png` compara cada partida contra su avance planificado.

La paleta usa fondo negro y tonos morados para mantener el estilo solicitado. Los valores, presupuesto, cronograma y nombres de partidas se pueden ajustar en `crear_datos_sinteticos()`.

## Panel web

La aplicación Streamlit carga el mismo Excel y presenta indicadores, tablas y gráficos en una interfaz de fondo negro y estilo morado.

```powershell
.\.venv\Scripts\streamlit.exe run app.py
```

Luego abre la dirección que Streamlit muestra normalmente en `http://localhost:8501`.
