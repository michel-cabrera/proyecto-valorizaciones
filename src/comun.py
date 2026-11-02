"""Rutas y utilidades compartidas por los scripts del pipeline."""

import json
import sys
from pathlib import Path

# Carpeta raíz del proyecto (la que contiene Leer_valorizaciones.py)
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))          # permite importar los scripts de la raíz

ENTRADA = ROOT / "data" / "entrada"        # valorizaciones a procesar
SALIDA = ROOT / "data" / "salida"          # resultados intermedios de cada etapa
EXTRACCION = SALIDA / "extraccion.json"    # etapa 2 -> etapa 3
VALIDADAS = SALIDA / "validadas.json"      # etapa 3 -> etapas 4 y 5
BD = ROOT / "valorizaciones.db"            # base de datos SQLite


def guardar_json(ruta, datos):
    ruta.parent.mkdir(parents=True, exist_ok=True)
    ruta.write_text(json.dumps(datos, ensure_ascii=False, indent=2, default=str), encoding="utf-8")


def leer_json(ruta, etapa_previa):
    if not ruta.exists():
        sys.exit(f"ERROR: no existe {ruta.name}. Ejecuta primero {etapa_previa}.")
    return json.loads(ruta.read_text(encoding="utf-8"))
