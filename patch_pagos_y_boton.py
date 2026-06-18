# -*- coding: utf-8 -*-
"""
Parche sobre el .xlsm vivo de Nomina por proyecto (editado en Excel por el
usuario, ya con vbaProject.bin y el boton CARGAR GASTO). Hace dos cosas SIN
tocar el VBA ni romper el boton:

  1) ARREGLA EL BOTON "CARGAR GASTO".
     En la hoja Captura habia 4 botones encimados; solo uno (_x0000_s1025)
     tenia la macro CargarGasto asignada, pero quedaba DEBAJO de otro boton
     visible sin macro (_x0000_s1026), por eso el clic no hacia nada.
     Dejamos un unico boton (s1025, con la macro) y eliminamos los otros 3
     de forma coordinada: VML, drawing1.xml, <controls>, ctrlProps y rels.

  2) AGREGA LA HOJA "Pagos fijos" (cronograma de pagos fijos).
     Tabla con Fecha de pago, Concepto, Categoria, Proyecto (lista),
     Periodicidad, Monto, Estatus (Pendiente/Pagado) y Notas. Cada pago se
     asigna a un proyecto con la lista desplegable; cuando se marca "Pagado"
     su monto se suma al proyecto en la hoja Resumen (nueva columna
     "Pagos fijos"). No requiere macros nuevas.

Uso:  python patch_pagos_y_boton.py  <entrada.xlsm>  [salida.xlsm]
"""
import sys, re, zipfile

SRC = sys.argv[1] if len(sys.argv) > 1 else "Nomina_por_proyecto.xlsm"
OUT = sys.argv[2] if len(sys.argv) > 2 else "Nomina_por_proyecto.xlsm"

NEW_SHEET_FILE = "xl/worksheets/sheet10.xml"
NEW_SHEET_RID = "rId15"
NEW_SHEET_ID = "10"

# ---------------------------------------------------------------- leer zip
with zipfile.ZipFile(SRC) as z:
    parts = {n: z.read(n) for n in z.namelist()}


def get(n):
    return parts[n].decode("utf-8")


def put(n, s):
    parts[n] = s.encode("utf-8")


# ======================================================================
# 1) ARREGLAR EL BOTON
# ======================================================================

# --- 1a) VML: conservar solo la <v:shape> que tiene <x:FmlaMacro> ---
vml = get("xl/drawings/vmlDrawing1.vml")
head = vml[: vml.index("<v:shape")]
shapes = re.findall(r"<v:shape\b.*?</v:shape>", vml, re.S)
keep = [s for s in shapes if "FmlaMacro" in s]
assert keep, "No se encontro la shape VML con FmlaMacro"
put("xl/drawings/vmlDrawing1.vml", head + keep[0] + "</xml>")

# --- 1b) drawing1.xml: conservar solo el anchor del shape id=1025 ---
dr = get("xl/drawings/drawing1.xml")
dr_head = dr[: dr.index("<mc:AlternateContent")]
blocks = re.findall(r"<mc:AlternateContent\b.*?</mc:AlternateContent>", dr, re.S)
keepb = [b for b in blocks if 'id="1025"' in b]
assert keepb, "No se encontro el anchor de dibujo del boton 1025"
# El macro debe ser el nombre simple "CargarGasto" (igual que en el VML y en
# el controlPr, y que el que aparece en Alt+F8). Un prefijo tipo "[0]!" hace
# que Excel no resuelva la macro al pulsar el boton ("no se puede ejecutar...").
block = keepb[0].replace('macro="" textlink=""', 'macro="CargarGasto" textlink=""')
put("xl/drawings/drawing1.xml", dr_head + block + "</xdr:wsDr>")

# --- 1c) sheet2.xml (Captura): dejar solo el <control shapeId=1025> ---
s2 = get("xl/worksheets/sheet2.xml")
ctrls = re.search(r"<controls>(.*)</controls>", s2, re.S).group(1)
cblocks = re.findall(r"<mc:AlternateContent\b.*?</mc:AlternateContent>", ctrls, re.S)
keepc = [c for c in cblocks if 'shapeId="1025"' in c]
assert keepc, "No se encontro el control 1025"
s2 = re.sub(r"<controls>.*</controls>", "<controls>" + keepc[0] + "</controls>", s2, flags=re.S)
put("xl/worksheets/sheet2.xml", s2)

# --- 1d) rels de Captura: quitar ctrlProp2/3/4 (rId5,6,7) ---
r2 = get("xl/worksheets/_rels/sheet2.xml.rels")
for rid in ("rId5", "rId6", "rId7"):
    r2 = re.sub(r'<Relationship\b[^>]*Id="%s"[^>]*/>' % rid, "", r2)
put("xl/worksheets/_rels/sheet2.xml.rels", r2)

# --- 1e) borrar ctrlProps sobrantes ---
for n in ("xl/ctrlProps/ctrlProp2.xml", "xl/ctrlProps/ctrlProp3.xml", "xl/ctrlProps/ctrlProp4.xml"):
    parts.pop(n, None)

# ======================================================================
# 2) ESTILO NUEVO (fecha con borde, sin relleno)  -> indice 28
# ======================================================================
st = get("xl/styles.xml")
m = re.search(r'<cellXfs count="(\d+)">', st)
n_xf = int(m.group(1))
DATE_XF = n_xf  # nuevo indice
new_xf = '<xf numFmtId="14" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>'
st = st.replace("</cellXfs>", new_xf + "</cellXfs>")
st = st.replace('<cellXfs count="%d">' % n_xf, '<cellXfs count="%d">' % (n_xf + 1))
put("xl/styles.xml", st)

# Estilos reutilizados del libro:
S_TITLE = 21   # titulo blanco sobre azul, centrado
S_NOTE = 25    # nota gris, wrap
S_HEAD = 7     # encabezado de columna (blanco/azul, centrado, borde)
S_LBLR = 5     # etiqueta alineada a la derecha (negrita)
S_MONEY_HL = 6  # dinero, relleno claro, verde negrita (totales)
S_TXT = 8      # texto con borde
S_MONEY = 9    # dinero con borde

# ======================================================================
# 3) HOJA "Pagos fijos"  (sheet10.xml)
# ======================================================================
NOTA = ("Cronograma de pagos fijos (renta, servicios, impuestos, etc.). "
        "Elige el Proyecto en la lista; cuando marques el Estatus como "
        "'Pagado', el monto se carga a ese proyecto en la hoja Resumen.")

rows = []
# fila 1: titulo
c1 = ['<c r="A1" s="%d" t="inlineStr"><is><t>CRONOGRAMA DE PAGOS FIJOS</t></is></c>' % S_TITLE]
for col in "BCDEFGH":
    c1.append('<c r="%s1" s="%d"/>' % (col, S_TITLE))
rows.append('<row r="1" spans="1:8" ht="21">%s</row>' % "".join(c1))

# fila 2: nota (A2:D2) + total programado (F2 etiqueta, G2 valor)
rows.append(
    '<row r="2" spans="1:8">'
    '<c r="A2" s="%d" t="inlineStr"><is><t>%s</t></is></c>' % (S_NOTE, NOTA)
    + '<c r="B2" s="%d"/><c r="C2" s="%d"/><c r="D2" s="%d"/>' % (S_NOTE, S_NOTE, S_NOTE)
    + '<c r="F2" s="%d" t="inlineStr"><is><t>Total programado:</t></is></c>' % S_LBLR
    + '<c r="G2" s="%d"><f>SUM(F5:F1000)</f></c>' % S_MONEY_HL
    + "</row>"
)
# fila 3: total pagado
rows.append(
    '<row r="3" spans="1:8">'
    '<c r="F3" s="%d" t="inlineStr"><is><t>Total pagado:</t></is></c>' % S_LBLR
    + '<c r="G3" s="%d"><f>SUMIF(G5:G1000,"Pagado",F5:F1000)</f></c>' % S_MONEY_HL
    + "</row>"
)
# fila 4: encabezados
heads = ["Fecha de pago", "Concepto", "Categoria", "Proyecto",
         "Periodicidad", "Monto", "Estatus", "Notas"]
hc = []
for j, h in enumerate(heads):
    col = chr(ord("A") + j)
    hc.append('<c r="%s4" s="%d" t="inlineStr"><is><t>%s</t></is></c>' % (col, S_HEAD, h))
rows.append('<row r="4" spans="1:8">%s</row>' % "".join(hc))

# filas 5..64: vacias con formato (Fecha, txt, txt, txt, txt, Monto, txt, txt)
col_styles = [DATE_XF, S_TXT, S_TXT, S_TXT, S_TXT, S_MONEY, S_TXT, S_TXT]
for r in range(5, 65):
    cs = []
    for j, sidx in enumerate(col_styles):
        col = chr(ord("A") + j)
        cs.append('<c r="%s%d" s="%d"/>' % (col, r, sidx))
    rows.append('<row r="%d" spans="1:8">%s</row>' % (r, "".join(cs)))

dv = (
    '<dataValidations count="4">'
    '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="C5:C1000">'
    '<formula1>"Renta,Servicio,Impuesto,Predial,Mantenimiento,Otro"</formula1></dataValidation>'
    '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="D5:D1000">'
    '<formula1>ListaProyectos</formula1></dataValidation>'
    '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="E5:E1000">'
    '<formula1>"Unico,Mensual,Bimestral,Trimestral,Semestral,Anual"</formula1></dataValidation>'
    '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="G5:G1000">'
    '<formula1>"Pendiente,Pagado"</formula1></dataValidation>'
    "</dataValidations>"
)

sheet10 = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<dimension ref="A1:H64"/>'
    '<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>'
    '<sheetFormatPr defaultRowHeight="15"/>'
    "<cols>"
    '<col min="1" max="1" width="14" customWidth="1"/>'
    '<col min="2" max="2" width="28" customWidth="1"/>'
    '<col min="3" max="3" width="16" customWidth="1"/>'
    '<col min="4" max="4" width="18" customWidth="1"/>'
    '<col min="5" max="5" width="14" customWidth="1"/>'
    '<col min="6" max="6" width="14" customWidth="1"/>'
    '<col min="7" max="7" width="12" customWidth="1"/>'
    '<col min="8" max="8" width="30" customWidth="1"/>'
    "</cols>"
    "<sheetData>" + "".join(rows) + "</sheetData>"
    '<mergeCells count="2"><mergeCell ref="A1:H1"/><mergeCell ref="A2:D2"/></mergeCells>'
    + dv +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
    "</worksheet>"
)
put(NEW_SHEET_FILE, sheet10)

# ======================================================================
# 4) RESUMEN (sheet3.xml): nueva columna D "Pagos fijos", Costo total -> E
# ======================================================================
s3 = get("xl/worksheets/sheet3.xml")
s3 = s3.replace('<dimension ref="A1:D54"/>', '<dimension ref="A1:E54"/>')
s3 = s3.replace(
    '<cols><col min="1" max="1" width="22" customWidth="1"/><col min="2" max="4" width="18" customWidth="1"/></cols>',
    '<cols><col min="1" max="1" width="22" customWidth="1"/><col min="2" max="5" width="18" customWidth="1"/></cols>',
)
s3 = s3.replace('<mergeCell ref="A1:D1"/>', '<mergeCell ref="A1:E1"/>')

r1 = ('<row r="1" spans="1:5" ht="21" x14ac:dyDescent="0.35">'
      '<c r="A1" s="21" t="s"><v>22</v></c><c r="B1" s="21"/><c r="C1" s="21"/>'
      '<c r="D1" s="21"/><c r="E1" s="21"/></row>')
r3 = ('<row r="3" spans="1:5" x14ac:dyDescent="0.35">'
      '<c r="A3" s="7" t="s"><v>13</v></c>'
      '<c r="B3" s="7" t="s"><v>23</v></c>'
      '<c r="C3" s="7" t="s"><v>24</v></c>'
      '<c r="D3" s="7" t="inlineStr"><is><t>Pagos fijos (pagados)</t></is></c>'
      '<c r="E3" s="7" t="s"><v>25</v></c></row>')
data = [r1, r3]
for i in range(50):
    r = 4 + i
    cat = 4 + i
    # Total percibido: se LEE de la hoja del proyecto (col I = Total percibido).
    # Asi, si borras una fila en la hoja del proyecto, el Resumen baja solo.
    # INDIRECT arma la referencia 'NombreProyecto'!I6:I100000; IFERROR->0 si la
    # hoja aun no existe.
    bf = ('IF(Catalogos!$A$%d="","",IFERROR(SUM(INDIRECT("\'"&amp;'
          'Catalogos!$A$%d&amp;"\'!I6:I100000")),0))' % (cat, cat))
    data.append(
        '<row r="%d" spans="1:5" x14ac:dyDescent="0.35">' % r
        + '<c r="A%d" s="8" t="str"><f>Catalogos!$A$%d</f></c>' % (r, cat)
        + '<c r="B%d" s="9"><f>%s</f></c>' % (r, bf)
        + '<c r="C%d" s="9"><f>IF(Catalogos!$A$%d="","",B%d*FactorCargaSocial)</f></c>' % (r, cat, r)
        + '<c r="D%d" s="9"><f>IF(Catalogos!$A$%d="","",SUMIFS(PagosMonto,PagosProyecto,Catalogos!$A$%d,PagosEstatus,"Pagado"))</f></c>' % (r, cat, cat)
        + '<c r="E%d" s="10"><f>IF(Catalogos!$A$%d="","",B%d+C%d+D%d)</f></c>' % (r, cat, r, r, r)
        + "</row>"
    )
data.append(
    '<row r="54" spans="1:5" x14ac:dyDescent="0.35">'
    '<c r="A54" s="11" t="s"><v>26</v></c>'
    '<c r="B54" s="12"><f>SUM(B4:B53)</f></c>'
    '<c r="C54" s="12"><f>SUM(C4:C53)</f></c>'
    '<c r="D54" s="12"><f>SUM(D4:D53)</f></c>'
    '<c r="E54" s="12"><f>SUM(E4:E53)</f></c></row>'
)
s3 = re.sub(r"<sheetData>.*</sheetData>", "<sheetData>" + "".join(data) + "</sheetData>", s3, flags=re.S)
put("xl/worksheets/sheet3.xml", s3)

# --- 4b) _Movimientos (sheet8.xml): borrar filas de prueba (deja encabezado) ---
# El libro maestro es solo historico; el Resumen ya NO depende de el. Limpiamos
# las 2 filas de prueba ($30,000 Tabmex y Produccion) que quedaron cargadas.
s8 = get("xl/worksheets/sheet8.xml")
s8 = re.sub(r'<dimension ref="[^"]*"/>', '<dimension ref="A1:J1"/>', s8)
row1 = re.search(r'<row r="1".*?</row>', s8, re.S).group(0)
s8 = re.sub(r"<sheetData>.*</sheetData>", "<sheetData>" + row1 + "</sheetData>", s8, flags=re.S)
put("xl/worksheets/sheet8.xml", s8)

# ======================================================================
# 5) workbook.xml: nueva hoja, nombres definidos, recalculo total
# ======================================================================
wb = get("xl/workbook.xml")
wb = wb.replace(
    '<sheet name="Resumen" sheetId="3" r:id="rId3"/>',
    '<sheet name="Resumen" sheetId="3" r:id="rId3"/>'
    '<sheet name="Pagos fijos" sheetId="%s" r:id="%s"/>' % (NEW_SHEET_ID, NEW_SHEET_RID),
)
nuevos = (
    "<definedName name=\"PagosProyecto\">'Pagos fijos'!$D$5:$D$1000</definedName>"
    "<definedName name=\"PagosMonto\">'Pagos fijos'!$F$5:$F$1000</definedName>"
    "<definedName name=\"PagosEstatus\">'Pagos fijos'!$G$5:$G$1000</definedName>"
)
wb = wb.replace("</definedNames>", nuevos + "</definedNames>")
wb = wb.replace('<calcPr calcId="191029" iterateDelta="1E-4"/>',
                '<calcPr calcId="191029" iterateDelta="1E-4" fullCalcOnLoad="1"/>')
put("xl/workbook.xml", wb)

# --- workbook.xml.rels: agregar hoja nueva, quitar calcChain (rId13) ---
wr = get("xl/_rels/workbook.xml.rels")
wr = re.sub(r'<Relationship\b[^>]*Id="rId13"[^>]*/>', "", wr)  # calcChain
wr = wr.replace(
    "</Relationships>",
    '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet10.xml"/></Relationships>'
    % NEW_SHEET_RID,
)
put("xl/_rels/workbook.xml.rels", wr)

# ======================================================================
# 6) [Content_Types].xml: hoja nueva, quitar ctrlProp2/3/4 y calcChain
# ======================================================================
ct = get("[Content_Types].xml")
for n in (2, 3, 4):
    ct = ct.replace(
        '<Override PartName="/xl/ctrlProps/ctrlProp%d.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>' % n,
        "",
    )
ct = ct.replace(
    '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>',
    "",
)
ct = ct.replace(
    "</Types>",
    '<Override PartName="/xl/worksheets/sheet10.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
)
put("[Content_Types].xml", ct)

# borrar calcChain (se regenera al abrir; usamos fullCalcOnLoad)
parts.pop("xl/calcChain.xml", None)

# ======================================================================
# 7) docProps/app.xml: contar hoja y nombres nuevos (evita aviso de reparacion)
# ======================================================================
app = get("docProps/app.xml")
app = app.replace("<vt:lpstr>Hojas de cálculo</vt:lpstr></vt:variant><vt:variant><vt:i4>9</vt:i4>",
                  "<vt:lpstr>Hojas de cálculo</vt:lpstr></vt:variant><vt:variant><vt:i4>10</vt:i4>")
app = app.replace("<vt:lpstr>Rangos con nombre</vt:lpstr></vt:variant><vt:variant><vt:i4>5</vt:i4>",
                  "<vt:lpstr>Rangos con nombre</vt:lpstr></vt:variant><vt:variant><vt:i4>8</vt:i4>")
app = app.replace('<vt:vector size="14" baseType="lpstr">', '<vt:vector size="18" baseType="lpstr">')
app = app.replace("<vt:lpstr>Resumen</vt:lpstr>",
                  "<vt:lpstr>Resumen</vt:lpstr><vt:lpstr>Pagos fijos</vt:lpstr>")
app = app.replace("<vt:lpstr>ListaPuestos</vt:lpstr>",
                  "<vt:lpstr>ListaPuestos</vt:lpstr><vt:lpstr>PagosProyecto</vt:lpstr>"
                  "<vt:lpstr>PagosMonto</vt:lpstr><vt:lpstr>PagosEstatus</vt:lpstr>")
put("docProps/app.xml", app)

# ---------------------------------------------------------------- escribir
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for n, b in parts.items():
        z.writestr(n, b)
print("OK ->", OUT)
print("  estilo fecha nuevo (cellXfs idx):", DATE_XF)
