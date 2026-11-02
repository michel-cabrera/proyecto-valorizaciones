"""Etapa 5 del pipeline: genera los pedidos SAP de las valorizaciones validadas.

SAP_MODO (variable de entorno, la define el Jenkinsfile):
  simulado -> no se conecta a SAP; asigna un número de pedido simulado y lo registra
              en SQLite y en data/salida/pedidos_simulados.csv.
  qa       -> usa crear_pedidos_sap.py con SAP GUI abierto (solo Windows, ambiente de
              pruebas). Respeta el MODO_PRUEBA configurado en ese script.
Cada pedido queda vinculado a su valorización en la tabla 'pedidos', y la etapa de
validación lo usa para no volver a generar pedidos duplicados.
"""

import csv
import os
import sqlite3
import sys
from datetime import datetime

import comun
import Leer_valorizaciones as lector


def pendientes(con, cabeceras):
    hechos = {r[0] for r in con.execute("SELECT archivo FROM pedidos")}
    return [c for c in cabeceras if c["Archivo"] not in hechos]


def registrar(con, archivo, nro, modo, estado, mensaje):
    con.execute("INSERT OR REPLACE INTO pedidos VALUES (?,?,?,?,?,?)",
                (archivo, nro, modo, estado, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), mensaje))


def modo_simulado(con, lista):
    filas = []
    sello = datetime.now().strftime("%Y%m%d%H%M%S")
    for i, c in enumerate(lista, 1):
        nro = f"SIM-{sello}-{i:03d}"
        registrar(con, c["Archivo"], nro, "simulado", "SIMULADO", "Pedido simulado; no se conectó a SAP")
        filas.append([c["Archivo"], c["Proveedor"], c["Pedido Referencia"], c["Suma Líneas"], nro])
        print(f"  {c['Archivo']} -> pedido simulado {nro}")
    salida = comun.SALIDA / "pedidos_simulados.csv"
    salida.parent.mkdir(parents=True, exist_ok=True)
    with salida.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Valorización", "Proveedor", "Pedido Referencia", "Importe", "Pedido generado"])
        w.writerows(filas)
    print(f"Reporte: {salida}")
    return 0


def modo_qa(con, lista, lineas):
    if os.name != "nt":
        sys.exit("ERROR: el modo 'qa' necesita Windows con SAP GUI abierto.")
    import crear_pedidos_sap as sap            # requiere pywin32
    sap.DEJAR_EN_PANTALLA = False              # en el pipeline no se deja el pedido abierto
    sap.session = sap.conectar_sap()
    errores = 0
    for c in lista:
        grupos = lector.agrupar_para_pedido([l for l in lineas if l["Archivo"] == c["Archivo"]])
        p = {"Clase Pedido": sap.CLASE_PEDIDO, "Proveedor": str(c["Cód. Acreedor"]),
             "Org. Compras": sap.ORG_COMPRAS, "Grupo Compras": sap.GRUPO_COMPRAS,
             "Sociedad": sap.SOCIEDAD, "Pedido Referencia": str(c["Pedido Referencia"]),
             "Fecha Pedido": "", "Fondo Garantía %": c.get("Fondo Garantía %") or 0,
             "Mes": str(c.get("Mes") or "").upper(), "Año": c.get("Año")}
        try:
            estado, nro, msg = sap.crear_pedido(p, grupos)
        except Exception as exc:
            estado, nro, msg = "ERROR", "", str(exc)
            errores += 1
            try:
                sap.salir_sin_guardar()
            except Exception:
                pass
        print(f"  {c['Archivo']} -> {estado} {nro} | {msg}")
        if estado == "CREADO":                 # solo un pedido grabado evita repetirlo
            registrar(con, c["Archivo"], nro, "qa", estado, msg)
    return 1 if errores else 0


def main():
    modo = os.environ.get("SAP_MODO", "simulado").lower()
    usuario = os.environ.get("SAP_USER", "(sin usuario)")
    print(f"Generación de pedidos | modo: {modo} | usuario SAP: {usuario}")
    datos = comun.leer_json(comun.VALIDADAS, "src/validar.py")
    with sqlite3.connect(comun.BD) as con:
        lista = pendientes(con, datos["cabeceras"])
        if not lista:
            print("No hay valorizaciones pendientes de pedido.")
            return
        if modo == "simulado":
            codigo = modo_simulado(con, lista)
        elif modo == "qa":
            codigo = modo_qa(con, lista, datos["lineas"])
        else:
            sys.exit(f"ERROR: SAP_MODO desconocido '{modo}' (usa 'simulado' o 'qa')")
    print("Listo.")
    sys.exit(codigo)


if __name__ == "__main__":
    main()
