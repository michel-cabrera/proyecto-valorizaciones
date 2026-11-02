"""Etapa 4 del pipeline: registra las valorizaciones validadas en SQLite."""

import sqlite3
from datetime import datetime

import comun

ESQUEMA = """
CREATE TABLE IF NOT EXISTS valorizaciones (
    archivo TEXT PRIMARY KEY, proveedor TEXT, cod_acreedor TEXT, ruc TEXT,
    contrato TEXT, mes TEXT, anio INTEGER, pedido_referencia TEXT,
    fondo_garantia REAL, total REAL, suma_lineas REAL, cuadre TEXT, fecha_carga TEXT
);
CREATE TABLE IF NOT EXISTS detalle (
    id INTEGER PRIMARY KEY AUTOINCREMENT, archivo TEXT REFERENCES valorizaciones(archivo),
    descripcion TEXT, n_servicio TEXT, imputacion TEXT, tipo_imputacion TEXT,
    clase_coste TEXT, cantidad REAL, tarifa REAL, importe REAL
);
CREATE TABLE IF NOT EXISTS pedidos (
    archivo TEXT PRIMARY KEY REFERENCES valorizaciones(archivo),
    nro_pedido TEXT, modo TEXT, estado TEXT, fecha TEXT, mensaje TEXT
);
"""


def numero(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def main():
    datos = comun.leer_json(comun.VALIDADAS, "src/validar.py")
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with sqlite3.connect(comun.BD) as con:
        con.executescript(ESQUEMA)
        for c in datos["cabeceras"]:
            con.execute("INSERT OR REPLACE INTO valorizaciones VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", (
                c["Archivo"], c["Proveedor"], c["Cód. Acreedor"], c["RUC"], c["Contrato"],
                c["Mes"], c["Año"] or None, c["Pedido Referencia"], numero(c["Fondo Garantía %"]),
                numero(c["Total Valorización"]), numero(c["Suma Líneas"]), c["Cuadre"], ahora))
            con.execute("DELETE FROM detalle WHERE archivo = ?", (c["Archivo"],))
        for l in datos["lineas"]:
            con.execute("INSERT INTO detalle (archivo, descripcion, n_servicio, imputacion, tipo_imputacion,"
                        " clase_coste, cantidad, tarifa, importe) VALUES (?,?,?,?,?,?,?,?,?)", (
                l["Archivo"], l["Descripción"], l["N° Servicio"], l["Imputación"], l["Tipo Imputación"],
                l["Clase de Coste"], numero(l["Cantidad"]), numero(l["Tarifa"]), numero(l["Importe"])))
        total = con.execute("SELECT COUNT(*) FROM valorizaciones").fetchone()[0]
    print(f"Cargadas {len(datos['cabeceras'])} valorización(es) y {len(datos['lineas'])} línea(s).")
    print(f"Total en la base de datos: {total} valorización(es) | archivo: {comun.BD}")


if __name__ == "__main__":
    main()
