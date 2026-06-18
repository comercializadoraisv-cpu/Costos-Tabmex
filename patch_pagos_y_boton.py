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

COMPRAS_FILE = "xl/worksheets/sheet11.xml"
COMPRAS_RID = "rId16"
COMPRAS_ID = "11"

# ---------------------------------------------------------------- leer zip
with zipfile.ZipFile(SRC) as z:
    parts = {n: z.read(n) for n in z.namelist()}


def get(n):
    return parts[n].decode("utf-8")


def put(n, s):
    parts[n] = s.encode("utf-8")


# ======================================================================
# 1) ARREGLAR EL BOTON
#    Lo reconstruimos COMPLETO como un unico control de formulario "legacy"
#    (solo VML + legacyDrawing), descartando toda la maquinaria moderna
#    (drawing1.xml, <controls>, ctrlProps) que venia enredada de varios
#    intentos. En un form control legacy la macro vive en <x:FmlaMacro> del
#    VML; ese es el unico binding y es el mas compatible.
# ======================================================================

VML_LIMPIO = (
    '<xml xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:x="urn:schemas-microsoft-com:office:excel">'
    '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>'
    '<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" '
    'path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>'
    '<v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>'
    '<o:lock v:ext="edit" shapetype="t"/></v:shapetype>'
    '<v:shape id="_x0000_s1025" type="#_x0000_t201" '
    "style='position:absolute;margin-left:16.5pt;margin-top:327pt;width:125.5pt;"
    "height:23.5pt;z-index:1;mso-wrap-style:tight' o:button=\"t\" "
    'fillcolor="buttonFace [67]" strokecolor="windowText [64]" o:insetmode="auto">'
    '<v:fill color2="buttonFace [67]" o:detectmouseclick="t"/>'
    '<o:lock v:ext="edit" rotation="t"/>'
    "<v:textbox style='mso-direction-alt:auto' o:singleclick=\"f\">"
    "<div style='text-align:center'><font face=\"Calibri\" size=\"200\" "
    'color="auto"><b>CARGAR GASTO</b></font></div></v:textbox>'
    '<x:ClientData ObjectType="Button">'
    "<x:Anchor>1, 0, 22, 0, 2, 9, 23, 18</x:Anchor>"
    "<x:PrintObject>False</x:PrintObject>"
    "<x:AutoFill>False</x:AutoFill>"
    "<x:FmlaMacro>CargarGasto</x:FmlaMacro>"
    "<x:TextHAlign>Center</x:TextHAlign>"
    "<x:TextVAlign>Center</x:TextVAlign>"
    "</x:ClientData></v:shape></xml>"
)
put("xl/drawings/vmlDrawing1.vml", VML_LIMPIO)

# sheet2.xml: cortar desde <drawing .../> hasta el final y dejar solo el
# legacyDrawing (apunta al VML). Asi se elimina drawing1.xml y el <controls>.
s2 = get("xl/worksheets/sheet2.xml")
s2 = re.sub(r'<drawing r:id="[^"]*"/>.*</worksheet>',
            '<legacyDrawing r:id="rId3"/></worksheet>', s2, flags=re.S)
put("xl/worksheets/sheet2.xml", s2)

# rels de Captura: solo printerSettings (rId1) + vmlDrawing (rId3).
put(
    "xl/worksheets/_rels/sheet2.xml.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="../printerSettings/printerSettings1.bin"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>'
    "</Relationships>",
)

# borrar drawing1.xml y todos los ctrlProps (ya no se usan)
for n in list(parts):
    if n.startswith("xl/ctrlProps/") or n == "xl/drawings/drawing1.xml":
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
# 3b) HOJA "Compras"  (sheet11.xml) - facturas de compra por proyecto
# ======================================================================
# Columnas: Fecha | Folio/UUID | Proveedor | RFC | Concepto | Categoria |
#           Proyecto | Subtotal | IVA | Total | Forma de pago | Estatus | Notas
NOTA_C = ("Facturas de compra. Elige el Proyecto en la lista: el costo de cada "
          "factura registrada se carga a ese proyecto en el Resumen (Subtotal, "
          "IVA y Total).")
crows = []
# fila 1: titulo (A1:M1)
cc = ['<c r="A1" s="%d" t="inlineStr"><is><t>FACTURAS DE COMPRA POR PROYECTO</t></is></c>' % S_TITLE]
for col in "BCDEFGHIJKLM":
    cc.append('<c r="%s1" s="%d"/>' % (col, S_TITLE))
crows.append('<row r="1" spans="1:13" ht="21">%s</row>' % "".join(cc))
# fila 2: nota (A2:E2) + totales sobre las columnas H, I, J
crows.append(
    '<row r="2" spans="1:13">'
    '<c r="A2" s="%d" t="inlineStr"><is><t>%s</t></is></c>' % (S_NOTE, NOTA_C)
    + '<c r="B2" s="%d"/><c r="C2" s="%d"/><c r="D2" s="%d"/><c r="E2" s="%d"/>' % (S_NOTE, S_NOTE, S_NOTE, S_NOTE)
    + '<c r="G2" s="%d" t="inlineStr"><is><t>Totales &#8594;</t></is></c>' % S_LBLR
    + '<c r="H2" s="%d"><f>SUM(H5:H1000)</f></c>' % S_MONEY_HL
    + '<c r="I2" s="%d"><f>SUM(I5:I1000)</f></c>' % S_MONEY_HL
    + '<c r="J2" s="%d"><f>SUM(J5:J1000)</f></c>' % S_MONEY_HL
    + "</row>"
)
# fila 4: encabezados
cheads = ["Fecha", "Folio / UUID", "Proveedor", "RFC", "Concepto",
          "Categoria", "Proyecto", "Subtotal", "IVA", "Total",
          "Forma de pago", "Estatus", "Notas"]
chc = []
for j, h in enumerate(cheads):
    col = chr(ord("A") + j)
    chc.append('<c r="%s4" s="%d" t="inlineStr"><is><t>%s</t></is></c>' % (col, S_HEAD, h))
crows.append('<row r="4" spans="1:13">%s</row>' % "".join(chc))
# filas 5..64 vacias con formato
ccol_styles = [DATE_XF, S_TXT, S_TXT, S_TXT, S_TXT, S_TXT, S_TXT,
               S_MONEY, S_MONEY, S_MONEY, S_TXT, S_TXT, S_TXT]
for r in range(5, 65):
    cs = []
    for j, sidx in enumerate(ccol_styles):
        col = chr(ord("A") + j)
        cs.append('<c r="%s%d" s="%d"/>' % (col, r, sidx))
    crows.append('<row r="%d" spans="1:13">%s</row>' % (r, "".join(cs)))

dvc = (
    '<dataValidations count="4">'
    '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="F5:F1000">'
    '<formula1>"Material,Herramienta,Equipo,Servicio,Combustible,Flete,Renta de equipo,Otro"</formula1></dataValidation>'
    '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="G5:G1000">'
    '<formula1>ListaProyectos</formula1></dataValidation>'
    '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="K5:K1000">'
    '<formula1>"Transferencia,Efectivo,Tarjeta,Cheque,Credito"</formula1></dataValidation>'
    '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="L5:L1000">'
    '<formula1>"Pendiente,Pagada"</formula1></dataValidation>'
    "</dataValidations>"
)
cwidths = [14, 22, 24, 16, 30, 16, 18, 14, 14, 14, 16, 12, 28]
ccols = "".join(
    '<col min="%d" max="%d" width="%d" customWidth="1"/>' % (j + 1, j + 1, w)
    for j, w in enumerate(cwidths)
)
sheet11 = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<dimension ref="A1:M64"/>'
    '<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>'
    '<sheetFormatPr defaultRowHeight="15"/>'
    "<cols>" + ccols + "</cols>"
    "<sheetData>" + "".join(crows) + "</sheetData>"
    '<mergeCells count="2"><mergeCell ref="A1:M1"/><mergeCell ref="A2:E2"/></mergeCells>'
    + dvc +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
    "</worksheet>"
)
put(COMPRAS_FILE, sheet11)

# ======================================================================
# 4) RESUMEN (sheet3.xml): Pagos fijos (D) + Compras (E,F,G) + Costo total (H)
# ======================================================================
s3 = get("xl/worksheets/sheet3.xml")
s3 = s3.replace('<dimension ref="A1:D54"/>', '<dimension ref="A1:H54"/>')
s3 = s3.replace(
    '<cols><col min="1" max="1" width="22" customWidth="1"/><col min="2" max="4" width="18" customWidth="1"/></cols>',
    '<cols><col min="1" max="1" width="22" customWidth="1"/><col min="2" max="8" width="15" customWidth="1"/></cols>',
)
s3 = s3.replace('<mergeCell ref="A1:D1"/>', '<mergeCell ref="A1:H1"/>')

r1 = ('<row r="1" spans="1:8" ht="21" x14ac:dyDescent="0.35">'
      '<c r="A1" s="21" t="s"><v>22</v></c><c r="B1" s="21"/><c r="C1" s="21"/>'
      '<c r="D1" s="21"/><c r="E1" s="21"/><c r="F1" s="21"/><c r="G1" s="21"/>'
      '<c r="H1" s="21"/></row>')
r3 = ('<row r="3" spans="1:8" x14ac:dyDescent="0.35">'
      '<c r="A3" s="7" t="s"><v>13</v></c>'
      '<c r="B3" s="7" t="s"><v>23</v></c>'
      '<c r="C3" s="7" t="s"><v>24</v></c>'
      '<c r="D3" s="7" t="inlineStr"><is><t>Pagos fijos (pagados)</t></is></c>'
      '<c r="E3" s="7" t="inlineStr"><is><t>Compras subtotal</t></is></c>'
      '<c r="F3" s="7" t="inlineStr"><is><t>Compras IVA</t></is></c>'
      '<c r="G3" s="7" t="inlineStr"><is><t>Compras total</t></is></c>'
      '<c r="H3" s="7" t="s"><v>25</v></c></row>')
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
        '<row r="%d" spans="1:8" x14ac:dyDescent="0.35">' % r
        + '<c r="A%d" s="8" t="str"><f>Catalogos!$A$%d</f></c>' % (r, cat)
        + '<c r="B%d" s="9"><f>%s</f></c>' % (r, bf)
        + '<c r="C%d" s="9"><f>IF(Catalogos!$A$%d="","",B%d*FactorCargaSocial)</f></c>' % (r, cat, r)
        + '<c r="D%d" s="9"><f>IF(Catalogos!$A$%d="","",SUMIFS(PagosMonto,PagosProyecto,Catalogos!$A$%d,PagosEstatus,"Pagado"))</f></c>' % (r, cat, cat)
        + '<c r="E%d" s="9"><f>IF(Catalogos!$A$%d="","",SUMIF(ComprasProyecto,Catalogos!$A$%d,ComprasSubtotal))</f></c>' % (r, cat, cat)
        + '<c r="F%d" s="9"><f>IF(Catalogos!$A$%d="","",SUMIF(ComprasProyecto,Catalogos!$A$%d,ComprasIVA))</f></c>' % (r, cat, cat)
        + '<c r="G%d" s="9"><f>IF(Catalogos!$A$%d="","",SUMIF(ComprasProyecto,Catalogos!$A$%d,ComprasTotal))</f></c>' % (r, cat, cat)
        + '<c r="H%d" s="10"><f>IF(Catalogos!$A$%d="","",B%d+C%d+D%d+G%d)</f></c>' % (r, cat, r, r, r, r)
        + "</row>"
    )
data.append(
    '<row r="54" spans="1:8" x14ac:dyDescent="0.35">'
    '<c r="A54" s="11" t="s"><v>26</v></c>'
    '<c r="B54" s="12"><f>SUM(B4:B53)</f></c>'
    '<c r="C54" s="12"><f>SUM(C4:C53)</f></c>'
    '<c r="D54" s="12"><f>SUM(D4:D53)</f></c>'
    '<c r="E54" s="12"><f>SUM(E4:E53)</f></c>'
    '<c r="F54" s="12"><f>SUM(F4:F53)</f></c>'
    '<c r="G54" s="12"><f>SUM(G4:G53)</f></c>'
    '<c r="H54" s="12"><f>SUM(H4:H53)</f></c></row>'
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
    '<sheet name="Pagos fijos" sheetId="%s" r:id="%s"/>'
    '<sheet name="Compras" sheetId="%s" r:id="%s"/>'
    % (NEW_SHEET_ID, NEW_SHEET_RID, COMPRAS_ID, COMPRAS_RID),
)
nuevos = (
    "<definedName name=\"PagosProyecto\">'Pagos fijos'!$D$5:$D$1000</definedName>"
    "<definedName name=\"PagosMonto\">'Pagos fijos'!$F$5:$F$1000</definedName>"
    "<definedName name=\"PagosEstatus\">'Pagos fijos'!$G$5:$G$1000</definedName>"
    "<definedName name=\"ComprasProyecto\">Compras!$G$5:$G$1000</definedName>"
    "<definedName name=\"ComprasSubtotal\">Compras!$H$5:$H$1000</definedName>"
    "<definedName name=\"ComprasIVA\">Compras!$I$5:$I$1000</definedName>"
    "<definedName name=\"ComprasTotal\">Compras!$J$5:$J$1000</definedName>"
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
    '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet10.xml"/>'
    '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet11.xml"/>'
    "</Relationships>" % (NEW_SHEET_RID, COMPRAS_RID),
)
put("xl/_rels/workbook.xml.rels", wr)

# ======================================================================
# 6) [Content_Types].xml: hoja nueva, quitar ctrlProp2/3/4 y calcChain
# ======================================================================
ct = get("[Content_Types].xml")
# quitar overrides de drawing1 y de TODOS los ctrlProps (boton reconstruido)
ct = re.sub(r'<Override PartName="/xl/drawings/drawing1\.xml"[^>]*/>', "", ct)
ct = re.sub(r'<Override PartName="/xl/ctrlProps/ctrlProp\d+\.xml"[^>]*/>', "", ct)
ct = ct.replace(
    '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>',
    "",
)
ct = ct.replace(
    "</Types>",
    '<Override PartName="/xl/worksheets/sheet10.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/worksheets/sheet11.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
)
put("[Content_Types].xml", ct)

# borrar calcChain (se regenera al abrir; usamos fullCalcOnLoad)
parts.pop("xl/calcChain.xml", None)

# ======================================================================
# 7) docProps/app.xml: contar hoja y nombres nuevos (evita aviso de reparacion)
# ======================================================================
app = get("docProps/app.xml")
app = app.replace("<vt:lpstr>Hojas de cálculo</vt:lpstr></vt:variant><vt:variant><vt:i4>9</vt:i4>",
                  "<vt:lpstr>Hojas de cálculo</vt:lpstr></vt:variant><vt:variant><vt:i4>11</vt:i4>")
app = app.replace("<vt:lpstr>Rangos con nombre</vt:lpstr></vt:variant><vt:variant><vt:i4>5</vt:i4>",
                  "<vt:lpstr>Rangos con nombre</vt:lpstr></vt:variant><vt:variant><vt:i4>12</vt:i4>")
app = app.replace('<vt:vector size="14" baseType="lpstr">', '<vt:vector size="23" baseType="lpstr">')
app = app.replace("<vt:lpstr>Resumen</vt:lpstr>",
                  "<vt:lpstr>Resumen</vt:lpstr><vt:lpstr>Pagos fijos</vt:lpstr>"
                  "<vt:lpstr>Compras</vt:lpstr>")
app = app.replace("<vt:lpstr>ListaPuestos</vt:lpstr>",
                  "<vt:lpstr>ListaPuestos</vt:lpstr><vt:lpstr>PagosProyecto</vt:lpstr>"
                  "<vt:lpstr>PagosMonto</vt:lpstr><vt:lpstr>PagosEstatus</vt:lpstr>"
                  "<vt:lpstr>ComprasProyecto</vt:lpstr><vt:lpstr>ComprasSubtotal</vt:lpstr>"
                  "<vt:lpstr>ComprasIVA</vt:lpstr><vt:lpstr>ComprasTotal</vt:lpstr>")
put("docProps/app.xml", app)

# ---------------------------------------------------------------- escribir
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for n, b in parts.items():
        z.writestr(n, b)
print("OK ->", OUT)
print("  estilo fecha nuevo (cellXfs idx):", DATE_XF)
