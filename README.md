\# Pipeline de valorizaciones y pedidos SAP – U.M. Inmaculada



Automatiza la lectura, validación y carga a SQLite de las valorizaciones

de contratistas y la generación de pedidos en SAP (modo simulado).

Jenkins ejecuta el pipeline tomando el código desde la rama main de este repositorio.



\## Etapas del pipeline (Jenkinsfile)

1\. Preparar Python: crea .venv e instala requirements.txt

2\. Probar lectura de valorizaciones: pytest y src/extraer.py

3\. Validación de calidad: cuadre de totales, RUC, contrato, acreedor, pedido de referencia y duplicados

4\. Carga a base de datos: SQLite (valorizaciones.db)

5\. Generación de pedidos SAP: modo simulado (SAP\_MODO = simulado)



Al terminar, Jenkins envía un correo con el resultado (OK o FALLÓ).



\## Estructura

\- data/entrada/: valorizaciones de prueba (datos simulados)

\- src/: scripts del pipeline (extraer, validar, cargar\_bd, generar\_pedidos)

\- tests/: pruebas automáticas con pytest

\- Leer\_valorizaciones.py y crear\_pedidos\_sap.py: lógica original de lectura y de SAP

\- Jenkinsfile: definición del pipeline



\## Modelo de ramas

\- main: versión estable; la ejecuta Jenkins. Solo recibe pull requests.

\- develop: integración de mejoras.

\- feature/\*: una rama por mejora.

\- hotfix/\*: correcciones urgentes sobre main.

\- docs/\*: cambios de documentación.



\## Cómo contribuir

1\. Crear una rama desde develop: git checkout -b feature/nombre

2\. Hacer commits con mensajes claros.

3\. Subir la rama y abrir un pull request hacia develop.

4\. Revisar e integrar; luego abrir un pull request de develop hacia main.



\## Ejecución local (Windows)

python -m venv .venv

.venv\\Scripts\\python.exe -m pip install -r requirements.txt

.venv\\Scripts\\python.exe -m pytest tests



\## Versiones

\- v1.0.0: pipeline con validación de calidad probada.



\## Integrantes

Cabrera Yauri Michel · Vega Herrera Gianella – DataOps, CERTUS

