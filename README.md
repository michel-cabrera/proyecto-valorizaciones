# Proyecto de valorizaciones

Estructura inicial para ejecutar el proyecto desde Jenkins en macOS.

- `data/entrada/`: archivos de entrada de valorizaciones.
- `src/`: scripts que ejecutará Jenkins.
- `tests/`: pruebas automatizadas.
- `requirements.txt`: paquetes de Python requeridos.

Ya existían `Leer_valorizaciones.py` y `crear_pedidos_sap.py` en la raíz del proyecto; los dejé intactos. El Pipeline compartido espera scripts con otros nombres dentro de `src/`. Antes de ejecutarlo hay que revisar qué hace cada script y adaptar el Pipeline o los archivos.
