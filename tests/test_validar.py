"""Pruebas de las reglas de calidad de src/validar.py."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from validar import validar_cabecera  # noqa: E402

CORRECTA = {"Cuadre": "OK", "RUC": "20123456789", "Contrato": "1013-0386-2025",
            "Cód. Acreedor": "1000001066", "Pedido Referencia": "4500000001"}


def test_valorizacion_correcta_no_tiene_errores():
    assert validar_cabecera(CORRECTA) == []


def test_detecta_total_que_no_cuadra():
    errores = validar_cabecera({**CORRECTA, "Cuadre": "DIFERENCIA 50.00"})
    assert any("no cuadra" in e for e in errores)


def test_detecta_ruc_invalido():
    errores = validar_cabecera({**CORRECTA, "RUC": "12345"})
    assert any("RUC" in e for e in errores)


def test_detecta_contrato_invalido():
    errores = validar_cabecera({**CORRECTA, "Contrato": "1013-386"})
    assert any("contrato" in e for e in errores)
