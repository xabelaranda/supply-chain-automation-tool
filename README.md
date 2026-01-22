# Supply Chain Automated Planner (SCAP)

Herramienta de escritorio desarrollada en Python para la automatización, análisis y simulación de planes de abastecimiento (Supply Planning). Diseñada para optimizar la toma de decisiones en cadenas de suministro internacionales con múltiples nodos de distribución.

## 🚀 Impacto y Resultados
El despliegue de esta herramienta en un entorno de producción real logró:
* **Reducción del 95% en tiempos operativos:** El proceso de generación de reportes pasó de **60 minutos a 3 minutos** por ciclo.
* **Eliminación de error humano:** Automatización completa del cálculo de KPIs críticos (DoS, OOS, USTN).
* **Simulación de Escenarios:** Capacidad de proyectar escenarios de "Recuperación Ideal" moviendo inventario virtualmente entre semanas para cubrir quiebres de stock.

## 🛠️ Características Técnicas

### 1. Procesamiento de Datos (Pandas)
* Ingesta de "Flat Files" masivos de ERPs.
* Cálculo vectorial de inventarios proyectados semana a semana.
* Lógica de detección de riesgos basada en cobertura (Days of Supply):
    * **OOS (Out of Stock):** Inventario agotado.
    * **USTN (Under Stock Target):** Inventario por debajo del nivel de seguridad.
    * **OSTN (Over Stock Target):** Exceso de inventario (High/Med).

### 2. Motor de Simulación (Algoritmo Greedy)
El sistema incluye un algoritmo de reabastecimiento inteligente que:
1.  Identifica semanas con quiebre de stock futuro.
2.  Busca envíos planificados en el futuro (Planned Shipments).
3.  Calcula si es factible "adelantar" esos envíos considerando Lead Times y restricciones de cierre de fábrica.
4.  Genera un "Plan de Acción" automático sugerido al planeador.

### 3. Reportes Automáticos (ReportLab & OpenPyXL)
* **Excel Formateado:** Genera hojas de cálculo con formato condicional nativo (colores, fuentes) aplicado desde Python para facilitar la lectura visual inmediata.
* **PDF Ejecutivo:** Genera un reporte PDF listo para imprimir con gráficos de evolución de riesgos y resumen ejecutivo.

### 4. Interfaz Gráfica (Tkinter)
Interfaz de usuario amigable que permite al usuario seleccionar archivos, configurar parámetros de simulación y elegir tipos de salida sin tocar código.

## 📦 Estructura del Proyecto
* `main.py`: Punto de entrada y orquestador de la interfaz gráfica.
* `reports.py`: Motor de generación de PDFs usando ReportLab.
* `styles.py`: Definiciones de estilos y formateo condicional para Excel (OpenPyXL).
* `utils.py`: Funciones auxiliares de limpieza y normalización de datos.
* `app_config.py`: Archivo de configuración centralizado (Mapeos, Colores, Parámetros).

## 🔧 Requisitos
* Python 3.8+
* Pandas
* OpenPyXL
* ReportLab

## 📄 Nota sobre Confidencialidad
Este repositorio contiene una versión sanitizada del código. Todos los nombres de productos, códigos de mercado y datos sensibles han sido reemplazados por valores genéricos o eliminados para cumplir con acuerdos de confidencialidad. La lógica funcional permanece intacta.
