BackUp_creator_v3.py
====================

Descripción
-----------
Script de copia de seguridad configurable para múltiples carpetas de origen, con soporte para exclusión de patrones, barra de progreso y notificación en escritorio.

Requisitos
----------
- Python 3.x
- Paquetes externos: instalar con
    pip install tqdm plyer

Uso
---
1. Ejecuta el script en tu entorno con:
   python BackUp_creator_v3.py

2. Selecciona una o varias carpetas de origen y la carpeta de destino usando el diálogo gráfico.
3. Opcional: introduce las carpetas o patrones que deseas excluir del respaldo (separados por comas, ejemplo: venv,node_modules,temp).
4. El proceso mostrará una barra de progreso y al finalizar recibirás una notificación, además de un log detallado en 'backuplog.txt'.

Archivos generados
------------------
- backupconfig.json: Guarda tu configuración para futuros respaldos.
- backuplog.txt: Log con detalles de cada copia, eliminaciones y errores.

Notas
-----
- El script copia archivos que hayan cambiado/que no existan en destino y elimina archivos en destino que ya no existen en origen, simulando un modo 'mirror'.
- Si modificas la configuración, puedes añadir nuevas carpetas de origen o cambiar patrones excluidos en cada ejecución.

Autor y licencia
----------------
Creador: Sergio Saavedra
Licencia: Libre para uso personal y educativo.

