"""
Lee todas las valorizaciones (Excel) de una carpeta y arma la base de datos
BD_VALORIZACIONES.xlsx con 3 pestañas:

  Cabeceras   -> una fila por valorización (proveedor, contrato, mes, pedidos, fondo %)
  Lineas      -> todas las líneas con importe (descripción, servicio, imputación...)
  Para pedido -> importes agrupados por N° servicio + imputación (lo que va a SAP)

Uso:
    python leer_valorizaciones.py                 -> lee la carpeta CARPETA_VALORIZACIONES
    python leer_valorizaciones.py "C:\\otra\\ruta"  -> lee otra carpeta

Requisito: pip install openpyxl
Las valorizaciones deben estar guardadas desde Excel (para que las fórmulas
tengan su valor calculado).
"""

import re
import sys
from datetime import date, timedelta
import unicodedata
from collections import OrderedDict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils import range_boundaries
from openpyxl.styles import Alignment, Font, PatternFill

# =====================================================================
# CONFIGURACIÓN
# =====================================================================
# Carpeta de SharePoint sincronizada (sin el año ni el mes; el script los completa).
# Estructura esperada:  ...\Backup 2026\2. Valorizaciones\10. Octubre\(ok) Belen\Valorización....xlsx
CARPETA_SHAREPOINT = (Path.home() / "Hochschild Mining PLC" / "UM Inmaculada - Documentos"
                      / "Control de Gestion")
SUBCARPETA_ANIO = "Backup {anio}"            # {anio} se reemplaza por el año
SUBCARPETA_VALOS = "2. Valorizaciones"
DIAS_CIERRE = 5                              # primeros días del mes -> se toma el mes anterior
SOLO_CARPETAS_OK = True                      # True = solo carpetas que empiezan con "(ok)"

# Contratistas que NO se leen (su pedido se genera de otra manera).
# Escribe el nombre como aparece en la carpeta, sin el "(ok)".
EXCLUIDOS = ["Altemec", "Zevallos", "Smart"]

# Carpetas con un solo pedido para todo el año (una posición por mes).
# En estas se toma el pedido más cercano al área de impresión de la hoja.
CARPETAS_PEDIDO_ANUAL = ["Cisternas MMAA"]
PATRON_ARCHIVOS = "Valorizaci*.xls*"         # solo archivos cuyo nombre empiece con "Valorizaci"
ARCHIVO_BD = Path(__file__).parent / "BD_VALORIZACIONES.xlsx"
VERSION = "v7 (29/10/2026)"

MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO",
         "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

# Encabezados de la tabla de la valorización (texto normalizado -> campo)
COLUMNAS_TABLA = OrderedDict([
    ("descripcion", "Descripción"), ("unidad", "Unidad"), ("cant", "Cantidad"),
    ("grupoarticulo", "Grupo Artículo"), ("numeroservicio", "N° Servicio"),
    ("imputacion", "Imputación"), ("clasedecoste", "Clase de Coste"),
    ("tarifaunitaria", "Tarifa"), ("importe", "Importe"),
])


def norm(t):
    t = unicodedata.normalize("NFKD", str(t or "")).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]", "", t.lower())


def fila_texto(celdas, fila):
    """Todo el contenido de una fila como texto (para datos en celdas vecinas)."""
    return " ".join(str(v.__int__() if isinstance(v, float) and v.is_integer() else v)
                    for r, c, v in sorted(celdas, key=lambda x: x[1]) if r == fila)


def nro_pedido(v):
    """Devuelve el N° de pedido (10 dígitos que empiezan con 45) o ''."""
    if isinstance(v, (int, float)):
        v = f"{v:.0f}"
    m = re.fullmatch(r"\s*(45\d{8})\s*", str(v or ""))
    return m.group(1) if m else ""


def es_pedido_anual(ruta):
    partes = [norm(re.sub(r"^\s*\(.*?\)\s*", "", x)) for x in ruta.parts]
    return any(norm(x) in partes for x in CARPETAS_PEDIDO_ANUAL)


def pedido_cercano_area_impresion(ws, celdas):
    """N° de pedido (45xxxxxxxx) más cercano al área de impresión de la hoja."""
    pedidos = [(r, c, nro_pedido(v)) for r, c, v in celdas if nro_pedido(v)]
    if not pedidos:
        return ""
    try:
        area = ws.print_area
        area = area[0] if isinstance(area, (list, tuple)) else area
        area = area.split(",")[0].split("!")[-1].replace("$", "").replace("'", "")
        c1, f1, c2, f2 = range_boundaries(area)
    except Exception:
        return pedidos[0][2]                  # sin área de impresión: el primero

    def distancia(r, c):
        dc = 0 if c1 <= c <= c2 else min(abs(c - c1), abs(c - c2))
        df = 0 if f1 <= r <= f2 else min(abs(r - f1), abs(r - f2))
        return dc + df
    return min(pedidos, key=lambda x: (distancia(x[0], x[1]), x[0], x[1]))[2]


def tipo_imputacion(obj):
    """K = centro de costo (IN + 8 dígitos), P = PEP, ? = revisar."""
    obj = str(obj or "").strip().upper()
    if re.fullmatch(r"IN\d{8}", obj):
        return "K"
    if re.fullmatch(r"IN\d{8}-[A-Z]", obj):
        return "?"            # ej. IN91240003-S: confirmar qué tipo es
    return "P" if obj else ""


# =====================================================================
# LECTURA DE UNA VALORIZACIÓN
# =====================================================================
def leer_valorizacion(ruta):
    ws = load_workbook(ruta, data_only=True).worksheets[0]
    celdas = [(c.row, c.column, c.value) for fila in ws.iter_rows(max_row=min(ws.max_row, 200))
              for c in fila if c.value not in (None, "")]
    textos = [(r, c, str(v).strip()) for r, c, v in celdas if isinstance(v, str)]

    cab = {"Carpeta": ruta.parent.name, "Archivo": ruta.name, "Proveedor": "", "Cód. Acreedor": "", "RUC": "",
           "Contrato": "", "Servicio": "", "Mes": "", "Año": "", "Pedido Referencia": "",
           "Pedido del Mes": "", "Fondo Garantía %": 0, "Total Valorización": None,
           "Suma Líneas": 0, "Cuadre": "", "Texto Corto": "", "Observación": ""}

    for r, c, t in textos:
        tu = t.upper()
        if tu.startswith("VALORIZACION ") and not cab["Proveedor"]:
            cab["Proveedor"] = t[len("VALORIZACION "):].strip()
            # la línea siguiente es el servicio con el mes
            sig = next((x for rr, cc, x in textos if rr == r + 1 and cc == c), "")
            # Texto corto = las 2 líneas del título: "VALORIZACION ... / SERVICIO ... MES AÑO"
            cab["Texto Corto"] = " / ".join(re.sub(r"\s+", " ", x).strip() for x in (t, sig) if x)          # título del encabezado de la valorización
            cab["Servicio"] = re.sub(r"[\s-]*(" + "|".join(MESES) + r")\s+\d{4}\s*$", "",
                                     sig, flags=re.I).strip()
            m = re.search(r"(" + "|".join(MESES + ["SEPTIEMBRE"]) + r")\s+(\d{4})", sig, re.I)
            if m:
                cab["Mes"] = m.group(1).upper().replace("SEPTIEMBRE", "SETIEMBRE")
                cab["Año"] = int(m.group(2))
        elif "CONTRA" in tu and not cab["Contrato"]:
            # acepta "1013-0386-2025" y "1013 - 0276 - 2025"
            m = re.search(r"(\d{4})\s*-\s*(\d{4})\s*-\s*(\d{4})", t)
            if m:
                cab["Contrato"] = "-".join(m.groups())
            else:
                m = re.search(r"(\d{4})\s*-\s*(\d{4})\s*-\s*(\d{4})", fila_texto(celdas, r))
                cab["Contrato"] = "-".join(m.groups()) if m else t
        if "ACREEDOR" in tu and not cab["Cód. Acreedor"]:
            # "Anexo Acreedor 1000001066", "Acreedor: 1000...", "Acreedor N° 1000..."
            # o el número en otra celda de la misma fila
            m = re.search(r"ACREEDOR\D{0,15}?(1\d{9})\b", tu) or \
                re.search(r"\b(1\d{9})\b", fila_texto(celdas, r))
            cab["Cód. Acreedor"] = m.group(1) if m else ""
        if re.search(r"R\.?\s*U\.?\s*C", tu) and not cab["RUC"]:
            m = re.search(r"\b(\d{11})\b", t) or re.search(r"\b(\d{11})\b", fila_texto(celdas, r))
            cab["RUC"] = m.group(1) if m else ""
        if "FONDO DE GARANT" in tu:
            m = re.search(r"(\d+(?:[.,]\d+)?)\s*%", t)
            cab["Fondo Garantía %"] = float(m.group(1).replace(",", ".")) if m else 0
        elif norm(t) == "pedido":
            nums = sorted((cc, nro_pedido(v)) for rr, cc, v in celdas
                          if rr == r and cc > c and nro_pedido(v))
            if nums:
                cab["Pedido Referencia"] = nums[-1][1]                # columna más a la derecha
                if len(nums) > 1:
                    cab["Pedido del Mes"] = nums[0][1]
                fila_ped, col_ref = r, nums[-1][0]

    # Mes y año: primero del nombre del archivo ("... Octubre 2026.xlsx"),
    # porque el texto de algunas valorizaciones trae fechas de inicio de contrato
    m = re.search(r"(" + "|".join(MESES + ["SEPTIEMBRE"]) + r")\W*(\d{4})", ruta.stem, re.I)
    if m:
        cab["Mes"] = m.group(1).upper().replace("SEPTIEMBRE", "SETIEMBRE")
        cab["Año"] = int(m.group(2))

    # Pedido anual (ej. Cisternas MMAA): el pedido más cercano al área de impresión
    if es_pedido_anual(ruta):
        anual = pedido_cercano_area_impresion(ws, celdas)
        if anual:
            cab["Pedido Referencia"] = anual
            cab["Pedido del Mes"] = anual
            cab["Observación"] = "Pedido anual (una posición por mes): no se crea pedido nuevo"

    # Tabla de líneas: fila del encabezado "DESCRIPCION"
    fila_enc = next((r for r, c, t in textos if norm(t) == "descripcion"), None)
    if not fila_enc:
        raise RuntimeError("No encontré la fila de encabezados (DESCRIPCION)")
    cols = {}
    for r, c, t in textos:
        if r == fila_enc:
            n = norm(t)
            for clave in COLUMNAS_TABLA:
                if n.startswith(clave) and clave not in cols:
                    cols[clave] = c
    faltan = [COLUMNAS_TABLA[k] for k in COLUMNAS_TABLA if k not in cols]
    if faltan:
        raise RuntimeError(f"Faltan columnas en la tabla: {faltan}")

    lineas = []
    for r in range(fila_enc + 1, ws.max_row + 1):
        desc = ws.cell(r, cols["descripcion"]).value
        if isinstance(desc, str) and desc.strip().upper().startswith("TOTAL"):
            cab["Total Valorización"] = ws.cell(r, cols["importe"]).value
            break
        servicio = ws.cell(r, cols["numeroservicio"]).value
        importe = ws.cell(r, cols["importe"]).value
        if servicio in (None, "") or not isinstance(importe, (int, float)):
            continue
        if round(importe, 2) == 0:
            continue
        imput = str(ws.cell(r, cols["imputacion"]).value or "").strip()
        lineas.append({
            "Archivo": ruta.name, "Proveedor": cab["Proveedor"], "Mes": cab["Mes"],
            "Descripción": str(desc or "").strip(),
            "Unidad": ws.cell(r, cols["unidad"]).value,
            "Cantidad": ws.cell(r, cols["cant"]).value,
            "Grupo Artículo": str(ws.cell(r, cols["grupoarticulo"]).value or ""),
            "N° Servicio": str(servicio).replace(".0", ""),
            "Imputación": imput, "Tipo Imputación": tipo_imputacion(imput),
            "Clase de Coste": str(ws.cell(r, cols["clasedecoste"]).value or "").replace(".0", ""),
            "Tarifa": ws.cell(r, cols["tarifaunitaria"]).value,
            "Importe": round(importe, 2),
        })

    cab["Suma Líneas"] = round(sum(l["Importe"] for l in lineas), 2)
    total = cab["Total Valorización"]
    if total is None:
        cab["Cuadre"] = "SIN TOTAL (¿archivo guardado desde Excel?)"
    else:
        cab["Total Valorización"] = round(total, 2)
        dif = round(total - cab["Suma Líneas"], 2)
        cab["Cuadre"] = "OK" if abs(dif) < 0.05 else f"DIFERENCIA {dif:,.2f}"
    return cab, lineas


def agrupar_para_pedido(lineas):
    """Suma importes por valorización + N° servicio + imputación."""
    grupos = OrderedDict()
    for l in lineas:
        k = (l["Archivo"], l["Tipo Imputación"], l["N° Servicio"], l["Imputación"])
        g = grupos.setdefault(k, {"Archivo": l["Archivo"], "Proveedor": l["Proveedor"],
                                  "Tipo Imputación": l["Tipo Imputación"],
                                  "N° Servicio": l["N° Servicio"], "Imputación": l["Imputación"],
                                  "Clase de Coste": l["Clase de Coste"], "Importe": 0.0,
                                  "N° Líneas": 0})
        g["Importe"] = round(g["Importe"] + l["Importe"], 2)
        g["N° Líneas"] += 1
    orden = {"K": 0, "?": 1, "P": 2, "": 3}
    return sorted(grupos.values(), key=lambda g: (g["Archivo"], orden.get(g["Tipo Imputación"], 9)))


# =====================================================================
# EXCEL DE SALIDA
# =====================================================================
def escribir_hoja(wb, titulo, filas, formatos):
    ws = wb.create_sheet(titulo)
    if not filas:
        return
    columnas = list(filas[0].keys())
    ws.append(columnas)
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for f in filas:
        ws.append([f[k] for k in columnas])
    for i, nombre in enumerate(columnas, 1):
        ancho = max(len(str(nombre)), *(len(str(f[nombre] or "")) for f in filas))
        ws.column_dimensions[ws.cell(1, i).column_letter].width = min(max(ancho + 2, 10), 55)
        for fila in ws.iter_rows(min_row=2, min_col=i, max_col=i):
            for c in fila:
                c.font = Font(name="Arial")
                if nombre in formatos:
                    c.number_format = formatos[nombre]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return ws


def carpeta_del_mes():
    """Carpeta Backup AAAA / 2. Valorizaciones / MM. Mes (con regla de cierre)."""
    hoy = date.today()
    if hoy.day <= DIAS_CIERRE:
        hoy = hoy.replace(day=1) - timedelta(days=1)
    base = CARPETA_SHAREPOINT / SUBCARPETA_ANIO.format(anio=hoy.year) / SUBCARPETA_VALOS
    if not base.exists():
        sys.exit(f"No encuentro la carpeta de SharePoint sincronizada:\n  {base}\n"
                 "Revisa CARPETA_SHAREPOINT en el script o pasa la carpeta al ejecutar.")
    for sub in base.iterdir():
        m = re.match(r"\s*0?(\d{1,2})\s*[.\-]", sub.name)
        if sub.is_dir() and m and int(m.group(1)) == hoy.month:
            return sub
    sys.exit(f"No encontré la carpeta del mes {hoy.month} dentro de {base}")


def buscar_archivos(carpeta):
    """Busca valorizaciones en la carpeta y sus subcarpetas (una por contratista)."""
    por_carpeta = {}
    omitidas = set()
    excluidas = set()
    lista_excluir = {norm(x) for x in EXCLUIDOS}
    for p in carpeta.rglob(PATRON_ARCHIVOS):
        if p.name.startswith("~$"):
            continue
        sub = p.parent
        partes = [re.sub(r"^\s*\(.*?\)\s*", "", x) for x in p.relative_to(carpeta).parts[:-1]]
        en_nombre = any(re.search(rf"\b{re.escape(x)}\b", p.stem, re.I) for x in EXCLUIDOS)
        if en_nombre or any(norm(x) in lista_excluir for x in partes):
            excluidas.add(str(p.relative_to(carpeta).parent))
            continue
        if SOLO_CARPETAS_OK and sub != carpeta and not sub.name.lower().startswith("(ok)"):
            omitidas.add(sub.name)
            continue
        por_carpeta.setdefault(sub, []).append(p)
    archivos = []
    for sub, lista in sorted(por_carpeta.items()):
        lista.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        if len(lista) > 1 and sub != carpeta:
            print(f"  Aviso: {sub.name} tiene {len(lista)} valorizaciones; uso la más reciente: {lista[0].name}")
            archivos.append(lista[0])
        else:
            archivos.extend(lista)
    for x in sorted(excluidas):
        print(f"  Excluida (lista EXCLUIDOS): {x}")
    for o in sorted(omitidas):
        print(f"  Omitida (no empieza con '(ok)'): {o}")
    return archivos


def main():
    carpeta = Path(sys.argv[1]) if len(sys.argv) > 1 else carpeta_del_mes()
    print(f"Leer valorizaciones {VERSION} | carpeta: {carpeta}")
    if not carpeta.exists():
        sys.exit(f"No existe la carpeta: {carpeta}")
    archivos = buscar_archivos(carpeta)
    if not archivos:
        sys.exit(f"No hay archivos '{PATRON_ARCHIVOS}' en la carpeta ni en sus subcarpetas.")

    cabeceras, lineas = [], []
    for ruta in archivos:
        try:
            cab, ls = leer_valorizacion(ruta)
            cabeceras.append(cab)
            lineas.extend(ls)
            print(f"  {ruta.parent.name} / {ruta.name}: {len(ls)} línea(s) | S/ {cab['Suma Líneas']:,.2f} | {cab['Cuadre']}")
        except Exception as e:
            print(f"  {ruta.name}: ERROR -> {e}")

    wb = Workbook()
    wb.remove(wb.active)
    dinero = "#,##0.00"
    escribir_hoja(wb, "Cabeceras", cabeceras,
                  {"Total Valorización": dinero, "Suma Líneas": dinero, "Fondo Garantía %": "0.##"})
    escribir_hoja(wb, "Lineas", lineas, {"Importe": dinero, "Tarifa": dinero, "Cantidad": "#,##0.000"})
    escribir_hoja(wb, "Para pedido", agrupar_para_pedido(lineas), {"Importe": dinero})
    try:
        wb.save(ARCHIVO_BD)
    except PermissionError:
        sys.exit(f"Cierra {ARCHIVO_BD.name} y vuelve a ejecutar.")
    print(f"Listo. Base de datos en: {ARCHIVO_BD}")


if __name__ == "__main__":
    main()