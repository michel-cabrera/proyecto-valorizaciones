"""Prueba de ejemplo de la lectura de una valorización Excel."""

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from Leer_valorizaciones import leer_valorizacion


class LeerValorizacionesTest(unittest.TestCase):
    def test_extrae_y_cuadra_una_linea_de_ejemplo(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            archivo = Path(temp_dir) / "Valorizacion Proveedor OCTUBRE 2026.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.append(["VALORIZACION PROVEEDOR PRUEBA"])
            ws.append(["Servicio de prueba OCTUBRE 2026"])
            ws.append(["Contrato: 1013-0386-2025"])
            ws.append(["Acreedor: 1000001066"])
            ws.append(["RUC: 12345678901"])
            ws.append(["Fondo de Garantía 10%"] )
            ws.append(["Pedido", "4500000001"])
            ws.append([])
            ws.append([
                "Descripción", "Unidad", "Cantidad", "Grupo Artículo",
                "Número Servicio", "Imputación", "Clase de Coste",
                "Tarifa Unitaria", "Importe",
            ])
            ws.append([
                "Servicio de prueba", "UND", 1, "SERV", "12345678",
                "IN91240003", "12345", 100, 100,
            ])
            ws.append(["TOTAL VALORIZACION", None, None, None, None, None, None, None, 100])
            wb.save(archivo)

            cabecera, lineas = leer_valorizacion(archivo)

        self.assertEqual(cabecera["Proveedor"], "PROVEEDOR PRUEBA")
        self.assertEqual(cabecera["Contrato"], "1013-0386-2025")
        self.assertEqual(cabecera["Pedido Referencia"], "4500000001")
        self.assertEqual(cabecera["Cuadre"], "OK")
        self.assertEqual(len(lineas), 1)
        self.assertEqual(lineas[0]["Tipo Imputación"], "K")
        self.assertEqual(lineas[0]["Importe"], 100)


if __name__ == "__main__":
    unittest.main()
