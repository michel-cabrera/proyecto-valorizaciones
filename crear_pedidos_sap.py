"""
Creación de pedidos de servicios en SAP (ME21N) a partir de BD_VALORIZACIONES.xlsx
(la base de datos que arma leer_valorizaciones.py).

Por cada valorización SIN "Pedido del Mes":
  cabecera -> copia el pedido de referencia (el del mes anterior) -> cambia el mes
  en los textos -> carga las líneas de servicio agrupadas por servicio + imputación
  (cantidad 1 x importe) en la posición de su tipo de imputación -> IVA -> fondo
  de garantía (ZF01) -> graba.
Los resultados se guardan en HISTORIAL_PEDIDOS.xlsx.

Requisitos (una sola vez):
    pip install pywin32 openpyxl

Uso:
    python crear_pedidos_sap.py "Trascawi"   -> solo valorizaciones que contengan ese texto
    python crear_pedidos_sap.py              -> todas las pendientes

IMPORTANTE: con MODO_PRUEBA = True el script llena el pedido pero NO lo graba.
"""

import re
import sys
import time
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

import win32com.client
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill

# =====================================================================
# CONFIGURACIÓN
# =====================================================================
MODO_PRUEBA = True          # True = NO graba los pedidos (solo los llena)
DEJAR_EN_PANTALLA = True    # En modo prueba: deja el pedido abierto en SAP (sin grabar)
                            # para que lo revises. Procesa solo UNA valorización.
ARCHIVO_BD = Path(__file__).with_name("BD_VALORIZACIONES.xlsx")
ARCHIVO_HISTORIAL = Path(__file__).with_name("HISTORIAL_PEDIDOS.xlsx")
VERSION = "v10 (29/10/2026)"

# Datos de cabecera iguales para todas las valorizaciones (de tus grabaciones)
CLASE_PEDIDO = "ZP07"
ORG_COMPRAS = "0400"
GRUPO_COMPRAS = "401"
SOCIEDAD = "1013"

# Tipo de imputación de la valorización -> tipo de imputación (KNTTP) de la
# posición del pedido donde se cargan sus líneas
POSICION_SEGUN_TIPO = {"K": ["K"], "P": ["P", "Q"], "F": ["F"]}
DIAS_CIERRE = 5             # primeros días del mes -> fecha = último día del mes anterior
# Condición de fondo de garantía según el porcentaje de la valorización.
# Todas las condiciones de fondo copiadas del pedido de referencia se BORRAN
# (vienen bloqueadas) y se agrega la correcta de nuevo: ClCd -> Enter -> Importe.
FONDO_POR_PORCENTAJE = {10: "ZF01", 5: "ZF03"}
IVA_POR_DEFECTO = "18"

# Nombre del menú "Selección" del resumen de documentos (Pedidos).
# Se prueba por texto y, si no funciona, con los códigos de tus grabaciones.
TEXTOS_VARIANTE = ["Pedidos", "Mis pedidos", "Pedido"]
CODIGOS_VARIANTE = ["CB863658240C1FE1B4CA34A5B6A22CE9NEW:PO_QUERY",
                    "CB863658240C1FE1B4AAD07F943E8CD8NEW:PO_QUERY"]

# Columna de imputación en las líneas de servicio según el tipo
CAMPO_IMPUTACION = {"P": "RM11P-PS_PSP_PNR",   # PEP
                    "K": "RM11P-KOSTL",        # Centro de costo
                    "F": "RM11P-AUFNR"}        # Orden

# Posición de columnas vista en tu grabación (respaldo si no se ubican por nombre)
COL_RESPALDO = {"ESLL-SRVPOS": 2, "ESLL-MENGE": 4, "ESLL-TBTWR": 6,
                "RM11P-PS_PSP_PNR": 11, "MEPO1211-TXZ01": 5, "MEPO1211-EEIND": 9,
                "KOMV-KSCHL": 1, "KOMV-KBETR": 3}

# =====================================================================
# IDs DE PANTALLA (tomados de tu grabación)
# El número "SAPLMEGUI:00xx" cambia según la vista; fid() prueba todos.
# =====================================================================
U = "wnd[0]/usr/subSUB0:SAPLMEGUI:0013"
CAB = f"{U}/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105"
TABS_CAB = (f"{U}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200"
            "/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL")
ORG = f"{TABS_CAB}/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221"
TEXTO_CAB = (f"{TABS_CAB}/tabpTABHDT3/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230"
             "/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell")
COND = f"{TABS_CAB}/tabpTABHDT2/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1236/subSUB0:SAPLV69A:6201"
TABLA_COND = f"{COND}/tblSAPLV69ATCTRL_KONDITIONEN"
BTN_INS_COND = f"{COND}/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KOLO"
TABLA_POS = (f"{U}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200"
             "/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")
DET = f"{U}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301"
COMBO_POS = f"{DET}/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST"
TABS_DET = f"{DET}/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL"
SERV = f"{TABS_DET}/tabpTABIDT1/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1328/subSUB0:SAPLMLSP:0400"
TABLA_SERV = f"{SERV}/tblSAPLMLSPTC_VIEW"
BTN_BORRAR_SERV = f"{SERV}/btnLOESCH"
IVA = f"{TABS_DET}/tabpTABIDT8/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1317/ctxtMEPO1317-MWSKZ"
RESUMEN = "wnd[0]/shellcont/shell/shellcont[1]/shell[0]"
ARBOL = "wnd[0]/shellcont/shell/shellcont[1]/shell[1]"
VARIANTES_PANTALLA = ["0013", "0010", "0020", "0017", "0015", "0016", "0019", "0012", "0014", "0018"]

session = None


# =====================================================================
# FUNCIONES SAP
# =====================================================================
def conectar_sap():
    try:
        sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
        return sap.Children(0).Children(0)
    except Exception:
        sys.exit("No encontré SAP abierto. Inicia sesión en SAP y vuelve a ejecutar.")


def fid(id_):
    """findById que prueba las distintas variantes de pantalla de la ME21N."""
    if "subSUB0:SAPLMEGUI:" not in id_:
        return session.findById(id_)
    for v in VARIANTES_PANTALLA:
        try:
            return session.findById(re.sub(r"subSUB0:SAPLMEGUI:00\d\d",
                                           f"subSUB0:SAPLMEGUI:{v}", id_, count=1))
        except Exception:
            pass
    raise RuntimeError(f"No encontré en pantalla: ...{id_.split('/')[-1]}")


def existe(id_):
    try:
        fid(id_)
        return True
    except Exception:
        return False


def barra_estado():
    sbar = session.findById("wnd[0]/sbar")
    return sbar.MessageType, sbar.Text      # S=éxito, W=aviso, E=error


def enter(ventana="wnd[0]"):
    session.findById(ventana).sendVKey(0)
    time.sleep(0.6)
    tipo, texto = barra_estado()
    if tipo in ("E", "A"):
        raise RuntimeError(texto)


def comando(cmd):
    session.findById("wnd[0]/tbar[0]/okcd").text = cmd
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)


def salir_sin_guardar():
    for _ in range(3):                                    # cerrar ventanas emergentes
        if existe("wnd[1]") and not existe("wnd[1]/usr/btnSPOP-OPTION2"):
            session.findById("wnd[1]").close()
    comando("/n")
    for boton in ("wnd[1]/usr/btnSPOP-OPTION2", "wnd[1]/usr/btnSPOP-VAROPTION2"):
        if existe(boton):
            session.findById(boton).press()               # "No" a grabar
            break


def col(tabla, campo):
    """Índice de la columna de una tabla SAP por el nombre de su campo."""
    try:
        for i in range(tabla.Columns.Count):
            if tabla.Columns.ElementAt(i).Name.upper().endswith(campo.upper()):
                return i
    except Exception:
        pass
    if campo in COL_RESPALDO:
        return COL_RESPALDO[campo]
    raise RuntimeError(f"La columna {campo} no está visible en la tabla de SAP")


def celda(id_tabla, fila_abs, campo):
    """Devuelve la celda de la fila absoluta, desplazando la tabla si hace falta."""
    tabla = fid(id_tabla)
    pos = tabla.VerticalScrollbar.Position
    if not (pos <= fila_abs < pos + tabla.VisibleRowCount):
        tabla.VerticalScrollbar.Position = fila_abs
        time.sleep(0.4)
        tabla = fid(id_tabla)
        pos = tabla.VerticalScrollbar.Position
    return tabla.GetCell(fila_abs - pos, col(tabla, campo))


# =====================================================================
# FORMATOS
# =====================================================================
def num(v):
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace(",", "")
    return float(t) if t else None


def fmt_precio(x):
    return f"{x:,.2f}"


def fmt_cant(x):
    t = f"{x:,.3f}".rstrip("0")
    return t + "0" if t.endswith(".") or len(t.split(".")[1]) < 2 else t


def a_fecha(v):
    if isinstance(v, (datetime, date)):
        return v
    for fmt in ("%d/%m/%Y", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(v).strip(), fmt)
        except ValueError:
            pass
    raise RuntimeError(f"Fecha no válida: {v} (usa dd/mm/aaaa)")


def fecha_documento(p):
    if p.get("Fecha Pedido"):
        return a_fecha(p["Fecha Pedido"])
    hoy = date.today()
    if hoy.day <= DIAS_CIERRE:
        return hoy.replace(day=1) - timedelta(days=1)
    return None


def fmt_fecha_sap(campo, fecha):
    actual = campo.text or ""
    sep = next((c for c in actual if not c.isdigit()), ".")
    return fecha.strftime(f"%d{sep}%m{sep}%Y")


MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO",
         "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
PATRON_MES = re.compile(r"\b(ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SETIEMBRE|"
                        r"SEPTIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE)(\s*(?:DEL?\s*)?\d{4})?\b",
                        re.IGNORECASE)


def actualizar_mes(texto, fecha):
    """Cambia el mes y año del texto copiado por los de la valorización.
    Ej.: 'VALOR.CAAS SETIEMBRE 2026 ORDEN' -> 'VALOR.CAAS OCTUBRE 2026 ORDEN'."""
    nuevo = f"{MESES[fecha.month - 1]}"
    def cambio(m):
        return f"{nuevo} {fecha.year}" if m.group(2) else nuevo
    return PATRON_MES.sub(cambio, texto or "")


# =====================================================================
# PASOS DEL PEDIDO
# =====================================================================
def paso_cabecera(p):
    comando("/nME21N")
    # Si quedó un pedido de una prueba anterior sin grabar, SAP pregunta si grabar: "No"
    for boton in ("wnd[1]/usr/btnSPOP-OPTION2", "wnd[1]/usr/btnSPOP-VAROPTION2"):
        if existe(boton):
            session.findById(boton).press()
            time.sleep(1)
            break
    fid(f"{CAB}/cmbMEPO_TOPLINE-BSART").key = p["Clase Pedido"]
    fid(f"{CAB}/ctxtMEPO_TOPLINE-SUPERFIELD").text = p["Proveedor"]
    f_doc = fecha_documento(p)
    if f_doc:
        c = fid(f"{CAB}/ctxtMEPO_TOPLINE-BEDAT")
        c.text = fmt_fecha_sap(c, f_doc)
    # Igual que en tu grabación: proveedor y datos organizativos juntos, y recién Enter
    try:
        fid(f"{TABS_CAB}/tabpTABHDT8").select()        # pestaña "Datos organizativos"
        time.sleep(0.4)
    except Exception:
        pass
    fid(f"{ORG}/ctxtMEPO1222-EKORG").text = p["Org. Compras"]
    fid(f"{ORG}/ctxtMEPO1222-EKGRP").text = p["Grupo Compras"]
    fid(f"{ORG}/ctxtMEPO1222-BUKRS").text = p["Sociedad"]
    enter()
    return f_doc


def elegir_variante_pedidos():
    resumen = session.findById(RESUMEN)
    resumen.pressContextButton("SELECT")
    for t in TEXTOS_VARIANTE:
        try:
            resumen.selectContextMenuItemByText(t)
            return
        except Exception:
            pass
    for c in CODIGOS_VARIANTE:
        try:
            resumen.selectContextMenuItem(c)
            return
        except Exception:
            pass
    raise RuntimeError("No encontré la opción 'Pedidos' en Selección del resumen de documentos")


def manejar_ventanas_copia(peps):
    """Ventanas que aparecen al copiar (avisos o pedido de PEP)."""
    for _ in range(6):
        time.sleep(0.5)
        if existe("wnd[3]"):
            ok = "wnd[3]/tbar[0]/btn[0]"
            (session.findById(ok).press() if existe(ok) else session.findById("wnd[3]").close())
            continue
        pep_id = "wnd[2]/usr/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-PS_POSID"
        if existe(pep_id):
            if not peps:
                raise RuntimeError("SAP pide un PEP al copiar y no hay PEP en el Detalle")
            session.findById(pep_id).text = peps[0]
            session.findById("wnd[2]").sendVKey(0)
            continue
        if existe("wnd[2]"):
            session.findById("wnd[2]").sendVKey(0)
            continue
        break


def paso_copiar_referencia(p, peps):
    if not existe(RESUMEN):
        session.findById("wnd[0]/tbar[1]/btn[8]").press()     # abrir resumen de documentos
    elegir_variante_pedidos()
    session.findById("wnd[0]/usr/ctxtSP$00014-LOW").text = p["Pedido Referencia"]
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    time.sleep(1.5)
    arbol = session.findById(ARBOL)
    claves = list(arbol.GetAllNodeKeys())
    if not claves:
        raise RuntimeError(f"No se encontró el pedido de referencia {p['Pedido Referencia']}")
    arbol.selectNode(claves[0])                                 # el pedido completo
    session.findById(RESUMEN).pressButton("COPY")
    manejar_ventanas_copia(peps)


def paso_resumen_posiciones(posiciones, f_entrega, f_mes):
    """Fecha de entrega y texto breve de TODAS las posiciones copiadas.
    Si el Excel no trae texto, se actualiza el mes del texto copiado."""
    i = 0
    while True:
        try:
            c_txt = celda(TABLA_POS, i, "MEPO1211-TXZ01")
        except Exception:
            break
        actual = (c_txt.text or "").strip()
        if not actual:
            break
        texto_excel = posiciones.get(i + 1, {}).get("texto", "")
        c_txt.text = (texto_excel or actualizar_mes(actual, f_mes))[:40]
        c = celda(TABLA_POS, i, "MEPO1211-EEIND")
        c.text = fmt_fecha_sap(c, f_entrega)
        i += 1
    if i == 0:
        raise RuntimeError("No se copiaron posiciones del pedido de referencia")
    enter()


def paso_texto_cabecera(texto, f_mes):
    """Si el Excel trae texto, se usa ese; si no, se toma el texto copiado
    del pedido de referencia y se le cambia el mes."""
    fid(f"{TABS_CAB}/tabpTABHDT3").select()
    time.sleep(0.4)
    editor = fid(TEXTO_CAB)
    if texto:
        lineas = str(texto).replace("\r\n", "\n").split("\n")
        editor.text = "\r".join(lineas) + "\r"
    else:
        actual = editor.text or ""
        if actual.strip():
            editor.text = actualizar_mes(actual, f_mes)


def paso_servicios(pos, lineas, iva):
    fid(COMBO_POS).key = f"{pos:>4}"                            # ir a la posición
    time.sleep(0.6)
    fid(f"{TABS_DET}/tabpTABIDT1").select()
    time.sleep(0.4)

    reemplazar = all(l["servicio"] for l in lineas)
    if reemplazar:                                             # borrar líneas copiadas
        tabla = fid(TABLA_SERV)
        for r in range(tabla.RowCount):
            try:
                tabla.getAbsoluteRow(r).selected = True
            except Exception:
                break
        fid(BTN_BORRAR_SERV).press()
        if existe("wnd[1]/usr/btnSPOP-OPTION1"):
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        time.sleep(0.5)

    for r, l in enumerate(lineas):
        if l["servicio"]:
            celda(TABLA_SERV, r, "ESLL-SRVPOS").text = l["servicio"]
        if l["cantidad"] is not None:
            celda(TABLA_SERV, r, "ESLL-MENGE").text = fmt_cant(l["cantidad"])
        if l["precio"] is not None:
            celda(TABLA_SERV, r, "ESLL-TBTWR").text = fmt_precio(l["precio"])
        if l["objeto"]:
            tipo = (l["tipo"] or "P").upper()[:1]
            if tipo not in CAMPO_IMPUTACION:
                raise RuntimeError(f"Tipo de imputación no válido: {l['tipo']} (usa P, K o F)")
            celda(TABLA_SERV, r, CAMPO_IMPUTACION[tipo]).text = l["objeto"]
    enter()

    if iva:
        fid(f"{TABS_DET}/tabpTABIDT8").select()
        fid(IVA).text = iva
        enter()


def filas_condiciones():
    """[(fila_visible, fila_absoluta, clase_condicion)] de la tabla de condiciones."""
    tabla = fid(TABLA_COND)
    i_cl = col(tabla, "KOMV-KSCHL")
    pos = tabla.VerticalScrollbar.Position
    filas = []
    for r in range(tabla.VisibleRowCount):
        try:
            filas.append((r, pos + r, (tabla.GetCell(r, i_cl).text or "").strip().upper()))
        except Exception:
            break
    return tabla, i_cl, filas


def boton_condiciones(*palabras):
    """Botón de la tabla de condiciones cuyo texto de ayuda contiene alguna palabra."""
    barra = fid(f"{COND}/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000")
    for i in range(barra.Children.Count):
        b = barra.Children.ElementAt(i)
        texto = f"{getattr(b, 'Tooltip', '')} {getattr(b, 'Text', '')}".lower()
        if any(p in texto for p in palabras):
            return b
    return None


def borrar_condicion(fila_abs, clase):
    tabla = fid(TABLA_COND)
    tabla.getAbsoluteRow(fila_abs).selected = True
    boton = boton_condiciones("borrar", "suprimir", "eliminar")
    if boton is None:
        raise RuntimeError(f"No encontré el botón para borrar la condición {clase}; bórrala a mano")
    boton.press()
    time.sleep(0.5)
    if existe("wnd[1]/usr/btnSPOP-OPTION1"):
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()


def paso_fondo_garantia(porcentaje):
    porcentaje = round(porcentaje or 0, 2)
    clase = FONDO_POR_PORCENTAJE.get(porcentaje) if porcentaje else None
    if porcentaje and not clase:
        raise RuntimeError(f"No hay condición configurada para fondo de garantía {porcentaje:g}% "
                           f"(revisa FONDO_POR_PORCENTAJE)")
    fid(f"{TABS_CAB}/tabpTABHDT2").select()
    time.sleep(0.5)

    # 1) Borrar TODAS las condiciones de fondo copiadas (vienen con el importe bloqueado)
    otras = set(FONDO_POR_PORCENTAJE.values())
    for _ in range(6):
        _, _, filas = filas_condiciones()
        sobra = next(((ab, cl) for _, ab, cl in filas if cl in otras), None)
        if not sobra:
            break
        borrar_condicion(*sobra)
        enter()
    if not clase:
        return

    # 2) Agregar la condición nueva en una fila vacía: ClCd + Enter
    tabla, i_cl, filas = filas_condiciones()
    vacia = next((r for r, _, cl in filas if not cl and tabla.GetCell(r, i_cl).Changeable), None)
    if vacia is None:
        fid(BTN_INS_COND).press()
        tabla, i_cl, filas = filas_condiciones()
        vacia = next(r for r, _, cl in filas if not cl)
    tabla.GetCell(vacia, i_cl).text = clase
    enter()

    # 3) Escribir el porcentaje en "Importe" (se vuelve a ubicar la fila: SAP reordena)
    tabla, i_cl, filas = filas_condiciones()
    fila = next((r for r, _, cl in filas if cl == clase), None)
    if fila is None:
        raise RuntimeError(f"No se pudo agregar la condición {clase}")
    importe = tabla.GetCell(fila, col(tabla, "KOMV-KBETR"))
    if not importe.Changeable:
        raise RuntimeError(f"El importe de {clase} sigue bloqueado; revisa la pestaña Condiciones")
    importe.text = f"{porcentaje:g}"
    enter()


def paso_grabar():
    if MODO_PRUEBA:
        _, texto = barra_estado()
        if DEJAR_EN_PANTALLA:
            return "PRUEBA OK", "", "Pedido llenado y dejado en pantalla para revisión (NO grabado)"
        salir_sin_guardar()
        return "PRUEBA OK", "", texto or "Pedido llenado sin errores (no se grabó)"
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    time.sleep(1)
    if existe("wnd[1]/usr/btnSPOP-VAROPTION1"):                  # "¿Grabar documento?"
        session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
        time.sleep(1.5)
    tipo, texto = barra_estado()
    numero = re.search(r"\d{10}", texto or "")
    if tipo == "S" and numero:
        return "CREADO", numero.group(), texto
    raise RuntimeError(texto or "SAP no confirmó la creación del pedido")


def leer_tipos_posiciones():
    """Tipo de imputación (KNTTP) de cada posición copiada: ['K', 'P', ...]."""
    tipos, i = [], 0
    while True:
        try:
            txt = (celda(TABLA_POS, i, "MEPO1211-TXZ01").text or "").strip()
        except Exception:
            break
        if not txt:
            break
        try:
            tipos.append((celda(TABLA_POS, i, "MEPO1211-KNTTP").text or "").strip().upper())
        except Exception:
            tipos.append("")
        i += 1
    return tipos


def crear_pedido(p, grupos):
    """p: datos de cabecera; grupos: líneas de 'Para pedido' de esa valorización."""
    peps = [g["Imputación"] for g in grupos if g["Tipo Imputación"] == "P"]
    f_doc = paso_cabecera(p)
    paso_copiar_referencia(p, peps)
    # Mes de los textos = mes de la valorización (de la BD); si no hay, el de la fecha del pedido
    f_mes = f_doc or date.today()
    if p.get("Mes") in MESES and p.get("Año"):
        f_mes = date(int(p["Año"]), MESES.index(p["Mes"]) + 1, 1)
    paso_resumen_posiciones({}, f_doc or date.today(), f_mes)
    paso_texto_cabecera("", f_mes)

    # Repartir las líneas en las posiciones según su tipo de imputación
    tipos_pos = leer_tipos_posiciones()
    if tipos_pos and not any(tipos_pos):
        # No se pudo leer el tipo de imputación: se asume el orden CeCo/Orden y luego PEP
        orden = sorted({g["Tipo Imputación"] for g in grupos}, key=lambda t: "KFP".find(t))
        tipos_pos = [POSICION_SEGUN_TIPO.get(t, [t])[0] for t in orden][:len(tipos_pos)]
    por_posicion = defaultdict(list)
    for g in grupos:
        tipo = g["Tipo Imputación"]
        if tipo not in POSICION_SEGUN_TIPO:
            raise RuntimeError(f"Imputación sin clasificar: {g['Imputación']} (tipo '{tipo}')")
        destino = next((i + 1 for i, t in enumerate(tipos_pos) if t in POSICION_SEGUN_TIPO[tipo]), None)
        if destino is None:
            raise RuntimeError(f"El pedido de referencia no tiene posición para imputación {tipo} "
                               f"(posiciones copiadas: {tipos_pos})")
        por_posicion[destino].append({"servicio": g["N° Servicio"], "cantidad": 1.0,
                                      "precio": float(g["Importe"]), "tipo": tipo,
                                      "objeto": g["Imputación"]})
    for pos in sorted(por_posicion):
        paso_servicios(pos, por_posicion[pos], IVA_POR_DEFECTO)
    paso_fondo_garantia(num(p.get("Fondo Garantía %")))
    enter()
    return paso_grabar()


# =====================================================================
# BASE DE DATOS E HISTORIAL
# =====================================================================
def leer_tabla(ws):
    enc = [str(c.value or "").strip() for c in ws[1]]
    return [dict(zip(enc, [c.value for c in fila])) for fila in ws.iter_rows(min_row=2)
            if any(c.value not in (None, "") for c in fila)]


COLS_HIST = ["Fecha", "Archivo", "Proveedor", "Cód. Acreedor", "Pedido Referencia",
             "Importe", "Modo", "Estado", "Pedido Creado", "Mensaje SAP"]


def abrir_historial():
    if ARCHIVO_HISTORIAL.exists():
        wb = load_workbook(ARCHIVO_HISTORIAL)
        return wb, wb.active
    wb = Workbook()
    ws = wb.active
    ws.title = "Historial"
    ws.append(COLS_HIST)
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center")
    for col_, ancho in zip("ABCDEFGHIJ", (17, 50, 45, 14, 17, 14, 10, 12, 15, 60)):
        ws.column_dimensions[col_].width = ancho
    ws.freeze_panes = "A2"
    for f, color in (('$H2="CREADO"', "C6EFCE"), ('$H2="PRUEBA OK"', "FFEB9C"), ('$H2="ERROR"', "FFC7CE")):
        ws.conditional_formatting.add("A2:J5000", FormulaRule(formula=[f], fill=PatternFill("solid", fgColor=color)))
    return wb, ws


def main():
    global session
    filtro = " ".join(sys.argv[1:]).strip().lower()
    print(f"Crear pedidos SAP {VERSION} | MODO_PRUEBA = {MODO_PRUEBA}"
          + (f" | filtro: '{filtro}'" if filtro else ""))
    if not ARCHIVO_BD.exists():
        sys.exit("No encuentro BD_VALORIZACIONES.xlsx. Ejecuta primero: python leer_valorizaciones.py")
    wb_bd = load_workbook(ARCHIVO_BD, data_only=True)
    cabeceras = leer_tabla(wb_bd["Cabeceras"])
    para_pedido = leer_tabla(wb_bd["Para pedido"])

    wb_h, ws_h = abrir_historial()
    creados = {str(r[1]) for r in ws_h.iter_rows(min_row=2, values_only=True) if r[7] == "CREADO"}

    pendientes = []
    for c in cabeceras:
        texto = f"{c.get('Carpeta', '')} {c.get('Archivo', '')} {c.get('Proveedor', '')}".lower()
        if filtro and filtro not in texto:
            continue
        motivo = ("ya tiene Pedido del Mes" if c.get("Pedido del Mes") else
                  "pedido anual" if c.get("Observación") else
                  "ya creado (historial)" if c.get("Archivo") in creados else
                  f"no cuadra ({c.get('Cuadre')})" if c.get("Cuadre") != "OK" else
                  "sin Cód. Acreedor" if not c.get("Cód. Acreedor") else
                  "sin Pedido Referencia" if not c.get("Pedido Referencia") else "")
        if motivo:
            print(f"  Se omite {c.get('Archivo')}: {motivo}")
            continue
        pendientes.append(c)
    if not pendientes:
        sys.exit("No hay valorizaciones pendientes para procesar con ese filtro.")

    revisar = MODO_PRUEBA and DEJAR_EN_PANTALLA
    if revisar and len(pendientes) > 1:
        print(f"  Modo revisión: se procesa solo la primera de {len(pendientes)} valorizaciones "
              "(usa un filtro para elegir cuál).")
        pendientes = pendientes[:1]

    session = conectar_sap()
    for c in pendientes:
        grupos = [g for g in para_pedido if g.get("Archivo") == c["Archivo"]]
        p = {"Clase Pedido": CLASE_PEDIDO, "Proveedor": str(c["Cód. Acreedor"]),
             "Org. Compras": ORG_COMPRAS, "Grupo Compras": GRUPO_COMPRAS, "Sociedad": SOCIEDAD,
             "Pedido Referencia": str(c["Pedido Referencia"]), "Fecha Pedido": "",
             "Fondo Garantía %": c.get("Fondo Garantía %") or 0,
             "Mes": str(c.get("Mes") or "").upper(), "Año": c.get("Año")}
        print(f"{c['Archivo']} | {len(grupos)} línea(s) | S/ {c.get('Suma Líneas') or 0:,.2f} ... ",
              end="", flush=True)
        try:
            estado, numero, msg = crear_pedido(p, grupos)
        except Exception as e:
            estado, numero, msg = "ERROR", "", str(e)
            if not revisar:                    # en modo revisión se deja la pantalla tal cual
                try:
                    salir_sin_guardar()
                except Exception:
                    pass
        print(estado, numero, "|", msg)
        ws_h.append([datetime.now().strftime("%d/%m/%Y %H:%M"), c["Archivo"], c.get("Proveedor"),
                     str(c["Cód. Acreedor"]), str(c["Pedido Referencia"]), c.get("Suma Líneas"),
                     "PRUEBA" if MODO_PRUEBA else "REAL", estado, numero, msg])
        ws_h.cell(ws_h.max_row, 6).number_format = "#,##0.00"

    try:
        wb_h.save(ARCHIVO_HISTORIAL)
    except PermissionError:
        sys.exit("Cierra HISTORIAL_PEDIDOS.xlsx y vuelve a ejecutar para guardar los resultados.")
    print(f"Listo. Resultados en {ARCHIVO_HISTORIAL.name}")
    if revisar:
        print("\n>>> El pedido quedó ABIERTO en SAP y NO está grabado. Revísalo con calma.")
        print(">>> Al terminar, sal con la flecha verde (atrás) y responde 'No' cuando pregunte si grabar.")


if __name__ == "__main__":
    main()