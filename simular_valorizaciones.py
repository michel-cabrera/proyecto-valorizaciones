"""Previsualiza valorizaciones y elegibilidad para pedidos, sin conectarse a SAP."""

import csv
from pathlib import Path

from openpyxl import load_workbook
import Leer_valorizaciones as lector


ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "data" / "entrada"
REPORT = ROOT / "SIMULACION_PEDIDOS.csv"
HISTORY = ROOT / "HISTORIAL_PEDIDOS.xlsx"


def pedidos_creados():
    if not HISTORY.exists():
        return set()
    wb = load_workbook(HISTORY, data_only=True, read_only=True)
    ws = wb.active
    return {
        str(row[1])
        for row in ws.iter_rows(min_row=2, values_only=True)
        if len(row) > 7 and row[7] == "CREADO" and row[1]
    }


def motivo_omision(cabecera, creados):
    if cabecera.get("Pedido del Mes"):
        return "Ya tiene Pedido del Mes"
    if cabecera.get("Observación"):
        return "Pedido anual"
    if cabecera.get("Archivo") in creados:
        return "Ya figura creado en el historial"
    if cabecera.get("Cuadre") != "OK":
        return "El cuadre no es OK"
    if not cabecera.get("Cód. Acreedor"):
        return "Falta código de acreedor"
    if not cabecera.get("Pedido Referencia"):
        return "Falta pedido de referencia"
    return "Elegible en simulación; no se ejecutó SAP"


def main():
    if not INPUT.exists():
        raise SystemExit(f"No existe la carpeta de entrada: {INPUT}")

    archivos = sorted(
        p for p in INPUT.rglob("Valorizaci*.xls*")
        if p.is_file() and not p.name.startswith("~$")
    )
    seleccionados = set(lector.buscar_archivos(INPUT))
    creados = pedidos_creados()
    filas = []
    errores = []
    candidatos = []

    for archivo in archivos:
        if archivo not in seleccionados:
            filas.append([archivo.name, "Omitida", 0, "", "No pasó las reglas de selección del lector"])
            continue
        try:
            cabecera, lineas = lector.leer_valorizacion(archivo)
        except Exception as exc:
            mensaje = str(exc).splitlines()[0]
            filas.append([archivo.name, "Error de lectura", 0, "", mensaje])
            errores.append((archivo.name, mensaje))
            continue

        razon = motivo_omision(cabecera, creados)
        importe = cabecera.get("Suma Líneas") or 0
        filas.append([archivo.name, "Revisión simulada", len(lineas), f"{importe:.2f}", razon])
        if razon.startswith("Elegible"):
            candidatos.append((archivo.name, len(lineas), importe))

    with REPORT.open("w", newline="", encoding="utf-8-sig") as salida:
        writer = csv.writer(salida)
        writer.writerow(["Valorización", "Estado", "Líneas", "Importe calculado", "Resultado de simulación"])
        writer.writerows(filas)

    print("SIMULACIÓN LOCAL: no se conectó a SAP ni se creó ningún pedido.")
    print(f"Archivos encontrados: {len(archivos)}")
    print(f"Candidatos para pedido simulado: {len(candidatos)}")
    print(f"Errores de lectura: {len(errores)}")
    print(f"Reporte: {REPORT}")
    for nombre, lineas, importe in candidatos:
        print(f"Candidato: {nombre} | líneas: {lineas} | importe: {importe:,.2f}")
    for nombre, mensaje in errores:
        print(f"Error: {nombre} | {mensaje}")

    if errores:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
