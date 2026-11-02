"""Etapa 3 del pipeline: validación de calidad antes de llegar a SAP.

Reglas (si alguna falla, el script termina con código 1 y Jenkins detiene el pipeline):
  - El archivo se pudo leer.
  - El total de la valorización cuadra con la suma de sus partidas.
  - RUC de 11 dígitos.
  - Contrato con formato 0000-0000-0000.
  - Código de acreedor SAP (10 dígitos que empiezan con 1).
  - Pedido de referencia presente (para copiarlo en SAP).
  - Sin valorizaciones duplicadas (mismo contrato, mes y año) en la misma ejecución.

No son errores, solo se omiten: las que ya tienen "Pedido del Mes", las de pedido
anual y las que ya tienen un pedido registrado en la base de datos.
"""

import re
import sqlite3
import sys

import comun


def validar_cabecera(cab):
    """Devuelve la lista de errores de una valorización (vacía si está correcta)."""
    errores = []
    if cab.get("Cuadre") != "OK":
        errores.append(f"el total no cuadra con las partidas ({cab.get('Cuadre')})")
    if not re.fullmatch(r"\d{11}", str(cab.get("RUC") or "")):
        errores.append(f"RUC inválido: '{cab.get('RUC')}'")
    if not re.fullmatch(r"\d{4}-\d{4}-\d{4}", str(cab.get("Contrato") or "")):
        errores.append(f"contrato con formato inválido: '{cab.get('Contrato')}'")
    if not re.fullmatch(r"1\d{9}", str(cab.get("Cód. Acreedor") or "")):
        errores.append("falta el código de acreedor SAP")
    if not cab.get("Pedido Referencia"):
        errores.append("falta el pedido de referencia")
    return errores


def motivo_omision(cab, con_pedido):
    if cab.get("Pedido del Mes"):
        return "ya tiene Pedido del Mes"
    if cab.get("Observación"):
        return "pedido anual"
    if cab.get("Archivo") in con_pedido:
        return "ya tiene pedido registrado en la base de datos"
    return ""


def archivos_con_pedido():
    """Archivos que ya tienen pedido en SQLite (control de duplicados entre ejecuciones)."""
    if not comun.BD.exists():
        return set()
    with sqlite3.connect(comun.BD) as con:
        existe = con.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='pedidos'").fetchone()
        if not existe:
            return set()
        return {r[0] for r in con.execute("SELECT archivo FROM pedidos")}


def main():
    datos = comun.leer_json(comun.EXTRACCION, "src/extraer.py")
    con_pedido = archivos_con_pedido()
    errores, aptas, omitidas, vistas = [], [], [], {}

    for e in datos["errores_lectura"]:
        errores.append(f"{e['Archivo']}: no se pudo leer ({e['Error']})")

    for cab in datos["cabeceras"]:
        nombre = cab["Archivo"]
        clave = (cab.get("Contrato"), cab.get("Mes"), cab.get("Año"))
        if clave in vistas:
            errores.append(f"{nombre}: duplicada con {vistas[clave]} (mismo contrato, mes y año)")
            continue
        vistas[clave] = nombre

        motivo = motivo_omision(cab, con_pedido)
        if motivo:
            omitidas.append({"Archivo": nombre, "Motivo": motivo})
            print(f"  Omitida {nombre}: {motivo}")
            continue
        problemas = validar_cabecera(cab)
        if problemas:
            errores.extend(f"{nombre}: {p}" for p in problemas)
        else:
            aptas.append(cab)
            print(f"  OK {nombre}")

    print(f"Aptas: {len(aptas)} | omitidas: {len(omitidas)} | errores: {len(errores)}")
    if errores:
        print("\nVALIDACIÓN FALLIDA. Corrige estas valorizaciones antes de generar pedidos:")
        for e in errores:
            print(f"  - {e}")
        sys.exit(1)

    archivos_aptos = {c["Archivo"] for c in aptas}
    lineas = [l for l in datos["lineas"] if l["Archivo"] in archivos_aptos]
    comun.guardar_json(comun.VALIDADAS, {"cabeceras": aptas, "lineas": lineas, "omitidas": omitidas})
    print(f"Validación correcta. Salida: {comun.VALIDADAS}")


if __name__ == "__main__":
    main()
