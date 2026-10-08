# RepartoPersonal

> Sistema inteligente y equitativo para la planificación, optimización y asignación mensual de personal y turnos de trabajo complejos.

[![Python](https://img.shields.io/badge/Python-3.11%20%7C%203.12-blue.svg)](https://www.python.org/)
[![OR-Tools](https://img.shields.io/badge/Solver-Google%20OR--Tools%20CP--SAT-orange.svg)](https://developers.google.com/optimization)
[![PySide6](https://img.shields.io/badge/GUI-PySide6%20%2F%20Qt-green.svg)](https://wiki.qt.io/Qt_for_Python)
[![ReportLab](https://img.shields.io/badge/PDF-ReportLab%20Autonomous-red.svg)](https://www.reportlab.com/)

---

## 📋 ¿Qué es RepartoPersonal?

**RepartoPersonal** es una aplicación diseñada para resolver el problema clásico de cuadrantes y rotación de puestos de trabajo en organizaciones con turnos rotativos, múltiples puestos especializados y restricciones operativas estrictas.

A partir de:
1. Una **secuencia base de turnos** por trabajador (mañana, tarde, refuerzo, descanso) a lo largo de un ciclo continuo (ej. ciclo de 24 semanas).
2. Una **ficha de competencias, vetos y preferencias** de cada trabajador.
3. Un **catálogo de puestos** con dotaciones mínimas y objetivos por turno y día de la semana.

El sistema formula un modelo de programación por restricciones con **Google OR-Tools CP-SAT** para obtener una distribución mensual matemáticamente óptima, equitativa y verificada.

---

## ✨ Características Principales

- 🧠 **Optimización con CP-SAT (3 Fases)**:
  - **Fase 1 (Viabilidad y Cobertura)**: Maximiza la cobertura de puestos obligatorios y dotaciones mínimas.
  - **Fase 2 (Equidad de Esfuerzo - Minimax D016)**: Minimiza la dispersión (*spread*) entre el trabajador con mayor carga promedio del mes y el de menor carga.
  - **Fase 3 (Preferencias Individuales)**: Maximiza las afinidades de cada trabajador dentro del rango óptimo de equidad garantizado.
- 🛡️ **Respeto Estricto de Restricciones**:
  - Nadie cubre un puesto sin tener acreditado el conocimiento.
  - Prohibición estricta de puestos vetados.
  - Un trabajador solo puede ser asignado en su turno disponible (Mañana o Tarde).
- 🏖️ **Gestión de Bajas Laborales**:
  - Soporte para marcar trabajadores en estado `BAJA`, excluyéndolos automáticamente del cálculo mensual.
- 🔍 **Detección y Auditoría de Vacantes**:
  - Si en un turno no hay suficiente personal o competencias para cubrir todos los puestos requeridos, el sistema **no bloquea la operativa**: asigna óptimamente el personal disponible, deja los huecos descubiertos en blanco e indica las vacantes para revisión manual.
- 📊 **Exportación Profesional a Excel (`.xlsx`)**:
  - Formato oficial con códigos cromáticos y siglas normalizadas (Mañanas en azul, Tardes en verde, domingos destacados).
  - Filas resumen de puestos descubiertos en la hoja principal.
  - Hoja de estadísticas individuales: días trabajados, puntos de esfuerzo acumulados, media por día y distribución por puesto.
- 📑 **Generación Autónoma de PDF (`.pdf`)**:
  - Maquetación directa en formato apaisado A4 mediante **ReportLab**, lista para imprimir sin requerir herramientas externas ni suites ofimáticas instaladas.
- 🖥️ **Interfaz Gráfica Sencilla (PySide6)**:
  - Selección intuitiva de mes, año y alineación de ciclo con validación previa de compatibilidad de fechas.

---

## 📁 Estructura del Proyecto

```text
RepartoPersonal/
├── iniciar.py               # Punto de entrada de la interfaz gráfica (PySide6)
├── compilar_exe.py          # Script para empaquetar en ejecutable portable (.exe)
│
├── reparto/                 # Lógica de dominio y cálculo
│   ├── datos.py             # Lectura de secuencias, competencias y calendario laboral
│   ├── solver.py            # Modelo de optimización matemática CP-SAT en 3 fases
│   ├── auditoria.py         # Verificación independiente de reglas y recuento de vacantes
│   └── exportadores.py      # Generador de Excel (.xlsx) y PDF nativo (.pdf)
│
├── configuracion_ejemplo/   # Plantillas Excel de ejemplo para probar el sistema
│   ├── Plantilla_Conocimientos_y_Preferencias.xlsx
│   └── Secuencias_Turnos_Limpio.xlsx
│
└── tests/                   # Pruebas automatizadas (unitarias y de integración)
    ├── fixtures.py
    ├── test_datos.py
    └── test_solver_pipeline.py
```

---

## 🚀 Instalación y Puesta en Marcha

### Requisitos
- **Python 3.11** o superior.
- Recomendado: gestor de paquetes [`uv`](https://docs.astral.sh/uv/) o `pip`.

### 1. Clonar el repositorio
```bash
git clone https://github.com/tu-usuario/reparto-personal.git
cd reparto-personal
```

### 2. Crear entorno virtual e instalar dependencias
Con `uv`:
```bash
uv venv
uv pip install ortools pyside6 reportlab pillow openpyxl
```

Con `pip` estándar:
```bash
python -m venv .venv
source .venv/bin/activate   # En Windows: .venv\Scripts\activate
pip install ortools pyside6 reportlab pillow openpyxl
```

### 3. Ejecutar la aplicación
```bash
python RepartoPersonal/iniciar.py
```

---

## ⚙️ Configuración de Entrada

Para planificar un equipo de trabajo, se utilizan dos archivos Excel en la carpeta `configuracion/` (puedes tomar como base los archivos de `configuracion_ejemplo/`):

1. **`Plantilla_Conocimientos_y_Preferencias.xlsx`**:
   - **Columna `BAJA`**: Marcar con `X` si el trabajador está de baja laboral para excluirlo de la planificación.
   - **Competencias**: Puestos que cada trabajador sabe desempeñar (marcados con `X` o `1`).
   - **Vetos**: Puestos que el trabajador tiene restringidos (marcados con `V`).
   - **Preferencias**: Escala de puntuación o ranking numérico para priorizar puestos preferidos.

2. **`Secuencias_Turnos_Limpio.xlsx`**:
   - Secuencia de turnos de cada trabajador a lo largo del ciclo (24 semanas / 168 días):
     - `M`: Turno de Mañana.
     - `T`: Turno de Tarde.
     - `R`: Turno de Refuerzo.
     - `D` (o celda vacía): Descanso / Libre.

---

## 🧪 Pruebas Automatizadas

El proyecto cuenta con una batería completa de pruebas unitarias y de integración:

```bash
python -m unittest discover RepartoPersonal/tests
```

Las pruebas verifican:
- Lectura de archivos y exclusión de trabajadores en baja.
- Cálculo de festivos y alineación de días de ciclo con el calendario.
- Solución factible y equitativa del solver sin violaciones de competencias ni vetos.
- Exportación íntegra de los archivos Excel y PDF.

---

## 📦 Compilación a Ejecutable Portable (Windows / Linux)

Para generar una versión autónoma sin requerir instalación previa de Python:

```bash
pip install pyinstaller
python RepartoPersonal/compilar_exe.py
```

El ejecutable listo para distribuir se generará en la carpeta `RepartoPersonal/dist/RepartoPersonal/`.

---

## 📄 Licencia

Este proyecto está disponible bajo licencia [MIT](LICENSE).
