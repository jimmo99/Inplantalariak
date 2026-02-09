📋 README - Buscador de Aplicaciones TIC
🚀 Descripción
Aplicación desktop desarrollada en Python para buscar y gestionar soluciones TIC dentro de archivos Excel. Permite búsquedas en tiempo real, selección de columnas, exportación de resultados y gestión de selecciones persistentes.

🎯 Características Principales
🔍 Búsqueda Avanzada
Búsqueda en tiempo real mientras escribes

Búsqueda en todas las columnas simultáneamente

Normalización de texto (elimina acentos, insensible a mayúsculas)

Resultados en tabla con columnas seleccionables

📊 Gestión de Datos
Selección múltiple de columnas para visualización

Tabla interactiva con scroll horizontal y vertical

Selección persistente entre búsquedas

Treeview responsive que se adapta al contenido

💾 Exportación y Copiado
Exportación a Excel/CSV solo con columnas seleccionadas

Copia al portapapeles formateada para Excel

Metadatos automáticos en exportaciones

Formato listo para pegar sin encabezados

🎨 Interfaz Responsive
Diseño adaptable a cualquier tamaño de ventana

Grid system profesional con pesos configurables

Elementos que se redimensionan automáticamente

Scrollbars integradas en todos los componentes

📋 Requisitos del Sistema
🔧 Software Necesario
Python 3.8 o superior

Sistema operativo: Windows, macOS o Linux

📦 Dependencias Python
bash
pip install pandas openpyxl
⚙️ Instalación y Configuración
1. 📥 Instalar Python
Descargar e instalar Python desde python.org

2. 🔧 Instalar Dependencias
Abrir la terminal o línea de comandos y ejecutar:

bash
pip install pandas openpyxl
3. 📁 Preparar Archivos
Colocar el archivo buscador_aplicaciones.py en una carpeta de tu elección

Tener acceso a los archivos Excel que quieras buscar

4. 🏃‍♂️ Ejecutar la Aplicación
Opción A - Doble click:

Hacer doble click en buscador_aplicaciones.py

Opción B - Línea de comandos:

bash
python buscador_aplicaciones.py
🎮 Guía de Uso
🏁 Inicio Rápido
Ejecutar la aplicación

Cargar archivo: Click en "📂 Cargar Excel"

Seleccionar columnas: Elige las columnas a mostrar en el panel izquierdo

Buscar: Escribe en el campo de búsqueda - los resultados aparecen automáticamente

🔍 Cómo Buscar
Escribe directamente en el campo de búsqueda

Los resultados se actualizan automáticamente

Busca en todas las columnas a la vez

No importan mayúsculas, minúsculas o acentos

📋 Cómo Seleccionar Datos
Click en cualquier fila de la tabla para añadir a la selección

Las selecciones se mantienen entre búsquedas

Usa "Seleccionar Todo Visible" para añadir todos los resultados de una vez

Gestiona individualmente con los botones de la columna derecha

💾 Cómo Exportar
"📋 Copiar": Copia solo los valores de columnas seleccionadas (listo para pegar en Excel)

"💾 Exportar": Guarda en archivo Excel o CSV manteniendo la estructura de columnas

🧹 Cómo Limpiar
"🧹 Limpiar": Limpia la búsqueda y resultados

"🧹 Limpiar" (panel derecho): Elimina todas las líneas seleccionadas

"🗑️ Quitar": Elimina una línea específica de la selección

🎯 Flujo de Trabajo Recomendado
🔄 Proceso Estándar
Cargar el archivo Excel con los datos

Seleccionar las columnas que necesitas ver

Buscar términos específicos escribiendo en el campo

Seleccionar las filas de interés haciendo click en la tabla

Exportar o copiar las selecciones según necesites

💡 Consejos Prácticos
Selecciona solo las columnas necesarias para mayor claridad

Usa "Seleccionar Todo Visible" para agilizar después de una búsqueda

Las selecciones son persistentes - puedes acumular resultados de múltiples búsquedas

El copiado al portapapeles mantiene el formato de columnas para Excel

🛠️ Solución de Problemas
❌ Error: "Archivo no encontrado"
Usa el botón "Cargar Excel" para seleccionar manualmente el archivo

Verifica que la ruta del archivo sea correcta

❌ Error: "Permiso denegado"
Cierra el archivo Excel si está abierto en otro programa

Asegúrate de tener permisos de lectura en el archivo

❌ La aplicación no inicia
Verifica que Python esté instalado correctamente

Ejecuta desde línea de comandos para ver mensajes de error específicos

❌ No se muestran los resultados
Asegúrate de haber seleccionado al menos una columna

Verifica que el término de búsqueda exista en los datos