# CUADROpz - Sistema de Gestión de Tareas Diarias

**CUADROpz** es una aplicación web interactiva desarrollada en Python utilizando **Streamlit**. Está diseñada para automatizar, organizar y optimizar la gestión de objetivos diarios y semanales de los colaboradores en una empresa, basándose en la estructura clásica de un cuadro de producción y persistiendo los datos de manera robusta.

## 🚀 Características Principales

1. **Pizarra Virtual por Persona:**
   - Cada colaborador (**MARY CRUZ**, **CPC.SHEYLA**, **CPC.HECTOR**) posee su propia pizarra con sus actividades del día.
   - Posibilidad de marcar tareas como "cumplidas" o "no cumplidas".
   - Control de prioridad ("Alta", "Media", "Baja", "Normal") y horarios.
   - **Carry-over automático:** Al finalizar el día laboral, las tareas no completadas se transfieren automáticamente al día hábil siguiente (omitiendo domingos), evitando duplicidades si ya existiesen tareas idénticas.

2. **Detección de Finalizaciones Pendientes:**
   - Si un usuario olvida cerrar su jornada, el sistema le recordará y le pedirá finalizar los días anteriores pendientes al iniciar sesión o cambiar de fecha, garantizando que el flujo de tareas acumuladas no se pierda.

3. **Informes Diarios Asistidos por IA:**
   - Al finalizar el día, se genera automáticamente un informe detallado con las actividades resueltas e irresueltas.
   - Para las tareas no resueltas, el sistema utiliza la API de OpenAI (modelo `gpt-3.5-turbo`) para proponer de forma autónoma **Alternativas de Solución** concisas y accionables.
   - Si no se proporciona una clave de OpenAI, la app activa de forma inteligente un **motor experto de sugerencias locales basado en reglas** de negocio contables y administrativas.

4. **Exportación Semanal a Excel con Formato Original:**
   - Permite descargar el cuadro de tareas semanal de todos los colaboradores con el formato idéntico del Excel original:
     - Fusión de celdas por colaborador en la columna `NOMBRES`.
     - Encabezados de días formateados (`LUNES 20/04`, etc.).
     - Codificación de color en base al estado de la tarea (Fondo **Amarillo** si fue cumplida, y colores pastel correspondientes si está pendiente: **Rojo** para Alta, **Naranja** para Media, **Verde** para Baja).
     - Incorporación automática de la **Leyenda de Colores** en la base del Excel generado.

5. **Calendario Flotante e Interactivo:**
   - Un botón flotante interactivo en forma de calendario en la esquina inferior derecha permite a los usuarios acceder rápidamente a una vista mensual.
   - Muestra de un vistazo los días con tareas pendientes mediante badges visuales.
   - Posibilidad de deslizar y explorar las tareas concretas de cualquier fecha seleccionada.

---

## 🛠️ Estructura del Código

El proyecto está diseñado bajo una arquitectura modular y limpia:
- **`app.py`**: Interfaz principal en Streamlit, manejo de estados, controles dinámicos e inyección de estilos CSS premium.
- **`database.py`**: Motor de almacenamiento SQLite local (`cuadropz.db`), importador del histórico inicial de Excel, operaciones CRUD de tareas e informes, y lógica de carry-over.
- **`excel_helper.py`**: Generador de archivos de Excel semanales formateados y colorizados utilizando `openpyxl`.
- **`ai_helper.py`**: Conexión con la API de OpenAI y motor de fallback para sugerencias empresariales locales.
- **`requirements.txt`**: Librerías necesarias para el entorno.

---

## 💻 Requisitos e Instalación

### 1. Clonar o descargar el repositorio
Asegúrate de colocar los archivos del código en tu carpeta de trabajo.

### 2. Crear un entorno virtual (Recomendado)
Abre tu consola de comandos en el directorio del proyecto y ejecuta:
```bash
python -m venv venv
venv\Scripts\activate      # En Windows
source venv/bin/activate   # En macOS/Linux
```

### 3. Instalar Dependencias
Instala los paquetes necesarios definidos en `requirements.txt`:
```bash
pip install -r requirements.txt
```

### 4. Inicializar e Importar Datos Históricos
Al ejecutar la aplicación por primera vez, el sistema detectará el archivo Excel original (`Excel de datos pizarra - copia.xlsx`) en el mismo directorio e **importará automáticamente las tareas de las 12 semanas históricas** a la base de datos local SQLite (`cuadropz.db`).

---

## 🏃‍♂️ Ejecución

Para iniciar el servidor local y abrir la aplicación web en tu navegador:
```bash
streamlit run app.py
```

La aplicación se abrirá por defecto en `http://localhost:8501`.

---

## ⚙️ Configuración de la API Key de OpenAI

- Puedes ingresar tu **OpenAI API Key** directamente en el cuadro de texto seguro ubicado en la parte inferior del menú lateral de la aplicación.
- Una vez ingresada, se guardará de forma persistente en la sesión para realizar las solicitudes.
- Si decides no usarla, el sistema de fallback rule-based seguirá sugiriendo soluciones inteligentes para tus actividades pendientes sin ningún problema.
