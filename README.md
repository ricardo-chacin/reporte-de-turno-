# PackTrack – Reporte de Turno (versión de escritorio)

Aplicación de escritorio para **estandarizar el reporte de fin de turno** de una línea de envasado: el operador ingresa los datos del turno y la app calcula automáticamente los indicadores clave (KPIs), los muestra en un tablero y los guarda en un historial consultable por fecha.

> **Nota:** proyecto personal con fines de aprendizaje y portafolio. No tiene relación oficial con ninguna empresa. Los formatos, velocidades y datos que aparecen son **valores de ejemplo**.

## Características

- Reporte por **turno (T1, T2, T3)** o por **día completo** (24 h).
- Hasta **dos formatos por turno** (para cuando hay cambio de formato).
- Cálculo automático de:
  - Hectolitros programados y envasados, y **cumplimiento**
  - **GLY**, **LEF** y **OSE**
  - **Merma** por formato y general (a partir de los másicos, explosiones, rotura y rechazos)
  - **TE (EPT)** – tiempo efectivo de producción
  - Ratios de consumo de **agua, vapor y CO₂** por hectolitro envasado
- Tablero de resultados con tarjetas de KPIs y gráfico de eficiencias.
- Comentarios del turno: fallas, actividades correctivas, análisis y avisos.
- **Historial** con calendario para consultar reportes anteriores.
- Guardado de los reportes en **Excel**.

## Tecnologías

Python · Tkinter · Matplotlib · tkcalendar · pandas · openpyxl

## Estructura del proyecto

```
├── main.py       # Punto de entrada y tabla de formatos (factor hl/botella y velocidad nominal)
├── ui.py         # Interfaz gráfica (formulario, resultados e historial)
├── logic.py      # Lógica de cálculo de KPIs
├── storage.py    # Guardado y lectura de reportes en Excel
├── styles.py     # Estilos de la interfaz
└── utils.py      # Utilidades (carga de logo opcional)
```

## Instalación y ejecución

1. Clona el repositorio:

   ```bash
   git clone https://github.com/TU_USUARIO/NOMBRE_REPO.git
   cd NOMBRE_REPO
   ```

2. (Opcional) crea un entorno virtual:

   ```bash
   python -m venv .venv
   .venv\Scripts\activate        # Windows
   source .venv/bin/activate     # Mac / Linux
   ```

3. Instala las dependencias:

   ```bash
   pip install matplotlib tkcalendar pandas openpyxl
   ```

4. Ejecuta la aplicación:

   ```bash
   python main.py
   ```

La app funciona sin logo. Si quieres mostrar uno, colócalo en la ruta que usa `utils.cargar_logo`.

## Cómo se calculan los indicadores

Con `tt` = 8 h para un turno o 24 h para el día completo:

| Concepto | Fórmula |
|---|---|
| Hectolitros | `botellas × factor` (hl por botella del formato) |
| Cumplimiento | `botellas envasadas / botellas programadas × 100` |
| EPT (h) | `Σ botellas del formato / velocidad nominal del formato` |
| OST | `tt − NST / 60` |
| ST | `OST` |
| LT | `ST − DPA / 60` |
| EC | `minutos perdidos externos / 60` |
| **GLY** | `EPT / ST × 100` |
| **LEF** | `EPT / (LT − EC) × 100` |
| **OSE** | `EPT / OST × 100` |
| **Merma** | `(másicos − hl envasados + explosiones + rotura + rechazos) / másicos × 100` |

NST, DPA y minutos perdidos se ingresan en el formulario, en minutos.

## Configurar los formatos

En `main.py`, cada formato se define con su factor de conversión y su velocidad nominal (botellas por hora):

```python
FORMATS = {
    "Formato 330": {"factor": 0.00330, "vel": 72000},
    "Formato 1000": {"factor": 0.01000, "vel": 19000},
}
```

El `factor` es el volumen de la botella en hectolitros (1 L = 0,01 hl).


Ricardo Chacin – Estudiante de Ingeniería de Software, con experiencia en operaciones industriales.
[LinkedIn](www.linkedin.com/in/ricardo-chacin-n) · [GitHub](Ricardo-Chacin)
