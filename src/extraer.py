"""Etapa 2 del pipeline: lee las valorizaciones de data/entrada.

Usa las funciones de Leer_valorizaciones.py (sin cambiar su lógica) y guarda
las cabeceras y líneas en data/salida/extraccion.json para la etapa de validación.
Los archivos que no se pueden leer se registran; la etapa de validación los rechaza.
"""

import sys

import comun
import Leer_valorizaciones as lector


def main():
    print(f"Extracción de valorizaciones | carpeta: {comun.ENTRADA}")
    if not comun.ENTRADA.exists():
        sys.exit(f"ERROR: no existe la carpeta de entrada {comun.ENTRADA}")

    archivos = lector.buscar_archivos(comun.ENTRADA)
    if not archivos:
        sys.exit(f"ERROR: no hay archivos '{lector.PATRON_ARCHIVOS}' en {comun.ENTRADA}")

    cabeceras, lineas, errores = [], [], []
    for ruta in archivos:
        try:
            cab, ls = lector.leer_valorizacion(ruta)
        except Exception as exc:
            mensaje = str(exc).splitlines()[0]
            errores.append({"Archivo": ruta.name, "Error": mensaje})
            print(f"  {ruta.name}: ERROR DE LECTURA -> {mensaje}")
            continue
        cabeceras.append(cab)
        lineas.extend(ls)
        print(f"  {ruta.name}: {len(ls)} línea(s) | S/ {cab['Suma Líneas']:,.2f} | cuadre: {cab['Cuadre']}")

    comun.guardar_json(comun.EXTRACCION, {"cabeceras": cabeceras, "lineas": lineas,
                                          "errores_lectura": errores})
    print(f"Leídas: {len(cabeceras)} | errores de lectura: {len(errores)} | salida: {comun.EXTRACCION}")


if __name__ == "__main__":
    main()
