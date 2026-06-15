# -*- coding: utf-8 -*-
"""
Genera el libro de Excel macro-habilitado para captura de gastos de nomina por
proyecto. Estructura faithful a la version entregada al usuario:
  Inicio, Captura, Resumen, Catalogos, <Proyectos>, _Movimientos, _Plantilla
"""
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter

PROYECTOS = ["Tabmex", "Contrato voceo", "Producción"]

# ----- Estilos reutilizables -----
AZUL = "1F4E79"
AZUL_CLARO = "DDEBF7"
GRIS = "F2F2F2"
VERDE = "548235"

titulo_font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
sub_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
label_font = Font(name="Calibri", size=11, bold=True, color="1F4E79")
normal_font = Font(name="Calibri", size=11)
nota_font = Font(name="Calibri", size=9, italic=True, color="808080")

header_fill = PatternFill("solid", fgColor=AZUL)
input_fill = PatternFill("solid", fgColor="FFF2CC")
light_fill = PatternFill("solid", fgColor=AZUL_CLARO)

thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

money = '#,##0.00 "MXN";[Red]-#,##0.00 "MXN"'
pct = '0.0%'

center = Alignment(horizontal="center", vertical="center")
left = Alignment(horizontal="left", vertical="center")
wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)

wb = Workbook()

# =====================================================================
# CATALOGOS  (se crea primero porque otras hojas lo referencian)
# =====================================================================
cat = wb.active
cat.title = "Catalogos"
cat.sheet_view.showGridLines = False

cat["A1"] = "CATALOGOS"
cat["A1"].font = titulo_font
cat["A1"].fill = header_fill
cat.merge_cells("A1:H1")
cat["A1"].alignment = center

# Proyectos
cat["A3"] = "Proyectos"
cat["A3"].font = sub_font
cat["A3"].fill = header_fill
for i, p in enumerate(PROYECTOS):
    c = cat.cell(row=4 + i, column=1, value=p)
    c.border = border
    c.font = normal_font

# Empleados
cat["C3"] = "Empleados"
cat["C3"].font = sub_font
cat["C3"].fill = header_fill
empleados = ["(captura aqui tus empleados)"]
for i, e in enumerate(empleados):
    c = cat.cell(row=4 + i, column=3, value=e)
    c.border = border
    c.font = normal_font

# Conceptos
cat["E3"] = "Conceptos"
cat["E3"].font = sub_font
cat["E3"].fill = header_fill
conceptos = ["Sueldo base", "Bono de productividad", "Bono puntualidad",
             "Aguinaldo", "Prima vacacional", "Comisiones", "Otra prestacion"]
for i, k in enumerate(conceptos):
    c = cat.cell(row=4 + i, column=5, value=k)
    c.border = border
    c.font = normal_font

# Factor de carga social
cat["G3"] = "Parametros"
cat["G3"].font = sub_font
cat["G3"].fill = header_fill
cat["G4"] = "Factor de carga social patronal (%)"
cat["G4"].font = label_font
cat["G4"].alignment = wrap
cat["H4"] = 0.0
cat["H4"].number_format = pct
cat["H4"].fill = input_fill
cat["H4"].border = border
cat["H4"].font = normal_font
cat["G6"] = ("Estimado de IMSS, Infonavit y provisiones (aguinaldo, prima "
             "vacacional). En 0% solo se cuenta lo pagado en mano. Subelo para "
             "reflejar el costo patronal real por proyecto.")
cat["G6"].font = nota_font
cat["G6"].alignment = wrap
cat.merge_cells("G6:H9")

cat.column_dimensions["A"].width = 20
cat.column_dimensions["B"].width = 3
cat.column_dimensions["C"].width = 28
cat.column_dimensions["D"].width = 3
cat.column_dimensions["E"].width = 24
cat.column_dimensions["F"].width = 3
cat.column_dimensions["G"].width = 24
cat.column_dimensions["H"].width = 14

# Rangos con nombre
wb.defined_names.add(DefinedName("FactorCargaSocial", attr_text="Catalogos!$H$4"))
wb.defined_names.add(DefinedName("ListaProyectos", attr_text="Catalogos!$A$4:$A$203"))
wb.defined_names.add(DefinedName("ListaEmpleados", attr_text="Catalogos!$C$4:$C$203"))
wb.defined_names.add(DefinedName("ListaConceptos", attr_text="Catalogos!$E$4:$E$23"))

# =====================================================================
# _MOVIMIENTOS  (libro maestro / auditoria)
# =====================================================================
mov = wb.create_sheet("_Movimientos")
mov_headers = ["Fecha", "Proyecto", "Empleado", "Concepto",
               "Sueldo", "Bonos", "Otras prestaciones", "Total percibido"]
for j, h in enumerate(mov_headers, start=1):
    c = mov.cell(row=1, column=j, value=h)
    c.font = sub_font
    c.fill = header_fill
    c.alignment = center
    c.border = border
widths = [12, 18, 24, 22, 14, 14, 16, 16]
for j, w in enumerate(widths, start=1):
    mov.column_dimensions[get_column_letter(j)].width = w
mov.freeze_panes = "A2"

# =====================================================================
# _PLANTILLA (oculta) -> base para hojas de proyecto nuevas
# =====================================================================
def build_project_sheet(ws, nombre):
    ws.sheet_view.showGridLines = False
    ws["A1"] = nombre
    ws["A1"].font = titulo_font
    ws["A1"].fill = header_fill
    ws.merge_cells("A1:H1")
    ws["A1"].alignment = center

    ws["A2"] = "GASTO DE NOMINA POR PROYECTO"
    ws["A2"].font = label_font
    ws.merge_cells("A2:H2")

    # Totales
    ws["F3"] = "Total percibido:"
    ws["F3"].font = label_font
    ws["F3"].alignment = Alignment(horizontal="right", vertical="center")
    ws["G3"] = "=SUM(G6:G100000)"
    ws["G3"].number_format = money
    ws["G3"].font = Font(bold=True, color=VERDE)
    ws["F4"] = "Costo total c/carga social:"
    ws["F4"].font = label_font
    ws["F4"].alignment = Alignment(horizontal="right", vertical="center")
    ws["G4"] = "=SUM(H6:H100000)"
    ws["G4"].number_format = money
    ws["G4"].font = Font(bold=True, color=VERDE)

    headers = ["Fecha", "Empleado", "Concepto", "Sueldo", "Bonos",
               "Otras prestaciones", "Total percibido", "Costo c/carga social"]
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=5, column=j, value=h)
        c.font = sub_font
        c.fill = header_fill
        c.alignment = center
        c.border = border
    ws_widths = [12, 24, 22, 14, 14, 16, 16, 18]
    for j, w in enumerate(ws_widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A6"

plant = wb.create_sheet("_Plantilla")
build_project_sheet(plant, "_Plantilla")

# Hojas de proyecto reales
for p in PROYECTOS:
    ws = wb.create_sheet(p)
    build_project_sheet(ws, p)

# =====================================================================
# CAPTURA
# =====================================================================
cap = wb.create_sheet("Captura", 0)  # primera posicion despues la movemos
cap.sheet_view.showGridLines = False

cap["B2"] = "CAPTURA DE GASTO DE NOMINA"
cap["B2"].font = titulo_font
cap["B2"].fill = header_fill
cap.merge_cells("B2:D2")
cap["B2"].alignment = center

campos = [
    (4, "Fecha", "C4", "fecha"),
    (5, "Proyecto", "C5", "proyecto"),
    (6, "Empleado", "C6", "empleado"),
    (7, "Concepto", "C7", "concepto"),
    (9, "Sueldo base", "C9", "money"),
    (10, "Bonos", "C10", "money"),
    (11, "Otras prestaciones", "C11", "money"),
]
for row, etiqueta, celda, tipo in campos:
    lc = cap.cell(row=row, column=2, value=etiqueta)
    lc.font = label_font
    lc.alignment = Alignment(horizontal="right", vertical="center")
    ic = cap[celda]
    ic.fill = input_fill
    ic.border = border
    ic.font = normal_font
    if tipo == "money":
        ic.number_format = money
    elif tipo == "fecha":
        ic.number_format = "dd/mm/yyyy"

# Total preview en C16
cap["B16"] = "TOTAL (vista previa)"
cap["B16"].font = Font(bold=True, color=AZUL, size=12)
cap["B16"].alignment = Alignment(horizontal="right", vertical="center")
cap["C16"] = "=N(C9)+N(C10)+N(C11)"
cap["C16"].number_format = money
cap["C16"].font = Font(bold=True, color=VERDE, size=12)
cap["C16"].fill = light_fill
cap["C16"].border = border

cap["B18"] = ("C16 es solo vista previa del registro; no es un boton. "
              "Llena los campos y ejecuta la macro CargarGasto (Alt+F8) o el boton.")
cap["B18"].font = nota_font
cap["B18"].alignment = wrap
cap.merge_cells("B18:D20")

cap.column_dimensions["A"].width = 3
cap.column_dimensions["B"].width = 22
cap.column_dimensions["C"].width = 26
cap.column_dimensions["D"].width = 6

# Validaciones (dropdowns)
dv_proj = DataValidation(type="list", formula1="=ListaProyectos", allow_blank=True)
dv_emp = DataValidation(type="list", formula1="=ListaEmpleados", allow_blank=True)
dv_con = DataValidation(type="list", formula1="=ListaConceptos", allow_blank=True)
cap.add_data_validation(dv_proj); dv_proj.add(cap["C5"])
cap.add_data_validation(dv_emp); dv_emp.add(cap["C6"])
cap.add_data_validation(dv_con); dv_con.add(cap["C7"])

# =====================================================================
# RESUMEN
# =====================================================================
res = wb.create_sheet("Resumen", 1)
res.sheet_view.showGridLines = False
res["A1"] = "RESUMEN CONSOLIDADO POR PROYECTO"
res["A1"].font = titulo_font
res["A1"].fill = header_fill
res.merge_cells("A1:D1")
res["A1"].alignment = center

rh = ["Proyecto", "Total percibido", "Carga social", "Costo total"]
for j, h in enumerate(rh, start=1):
    c = res.cell(row=3, column=j, value=h)
    c.font = sub_font
    c.fill = header_fill
    c.alignment = center
    c.border = border

# Una fila por proyecto del catalogo (hasta 50 filas con formula)
first = 4
last = first + 49
for i in range(50):
    r = first + i
    proj_ref = f"Catalogos!$A${4 + i}"
    res.cell(row=r, column=1, value=f"={proj_ref}").border = border
    # Total percibido
    tp = (f"=IF({proj_ref}=\"\",\"\","
          f"SUMIFS('_Movimientos'!$H:$H,'_Movimientos'!$B:$B,{proj_ref}))")
    c2 = res.cell(row=r, column=2, value=tp)
    c2.number_format = money; c2.border = border
    # Carga social
    cs = (f"=IF({proj_ref}=\"\",\"\",B{r}*FactorCargaSocial)")
    c3 = res.cell(row=r, column=3, value=cs)
    c3.number_format = money; c3.border = border
    # Costo total
    ct = (f"=IF({proj_ref}=\"\",\"\",B{r}+C{r})")
    c4 = res.cell(row=r, column=4, value=ct)
    c4.number_format = money; c4.border = border
    c4.font = Font(bold=True)

# Totales generales
res.cell(row=last + 1, column=1, value="TOTAL GENERAL").font = Font(bold=True, color=AZUL)
for col in (2, 3, 4):
    L = get_column_letter(col)
    c = res.cell(row=last + 1, column=col,
                 value=f"=SUM({L}{first}:{L}{last})")
    c.number_format = money
    c.font = Font(bold=True, color=VERDE)
    c.fill = light_fill
    c.border = border

res.column_dimensions["A"].width = 22
for L in ("B", "C", "D"):
    res.column_dimensions[L].width = 18

# =====================================================================
# INICIO  (guia)
# =====================================================================
ini = wb.create_sheet("Inicio", 0)
ini.sheet_view.showGridLines = False
ini["B2"] = "NOMINA POR PROYECTO  -  Guia rapida"
ini["B2"].font = titulo_font
ini["B2"].fill = header_fill
ini.merge_cells("B2:H2")
ini["B2"].alignment = center

pasos = [
    "1. Activa las macros: al abrir, si sale la barra amarilla arriba, pulsa 'Habilitar contenido'.",
    "2. (Solo la 1a vez) Importa la macro: Alt+F11 -> Archivo -> Importar archivo -> CargarGasto.bas.",
    "3. Ve a la hoja CAPTURA y llena: Fecha, Proyecto, Empleado, Concepto, Sueldo, Bonos, Otras prestaciones.",
    "4. Ejecuta la macro: Alt+F8 -> CargarGasto -> Ejecutar (o usa el boton si insertaste uno).",
    "5. El registro se manda a la hoja del proyecto (la crea si no existe) y al libro _Movimientos.",
    "6. RESUMEN y los totales por proyecto se actualizan solos (formulas SUMIFS).",
    "",
    "CATALOGOS: edita aqui tus Proyectos, Empleados, Conceptos y el Factor de carga social patronal (%).",
    "Si el factor es 0% solo se cuenta lo pagado en mano; subelo para reflejar el costo patronal real.",
    "",
    "IMPORTANTE: al guardar conserva el formato .xlsm (Excel a veces sugiere .xlsx, que borra las macros).",
]
for i, t in enumerate(pasos):
    c = ini.cell(row=4 + i, column=2, value=t)
    c.font = normal_font if t and not t[0].isdigit() else Font(size=11)
    if t.startswith(("IMPORTANTE", "CATALOGOS")):
        c.font = Font(bold=True, color=AZUL)
    ini.merge_cells(start_row=4 + i, start_column=2, end_row=4 + i, end_column=8)
    c.alignment = left
ini.column_dimensions["A"].width = 3
for L in "BCDEFGH":
    ini.column_dimensions[L].width = 16

# =====================================================================
# Orden de hojas y ocultar tecnicas
# =====================================================================
orden = ["Inicio", "Captura", "Resumen", "Catalogos"] + PROYECTOS + ["_Movimientos", "_Plantilla"]
wb._sheets.sort(key=lambda s: orden.index(s.title) if s.title in orden else 99)
wb["_Plantilla"].sheet_state = "hidden"
wb["_Movimientos"].sheet_state = "hidden"
wb.active = 0

wb.save("Nomina_por_proyecto.xlsm")
print("OK -> Nomina_por_proyecto.xlsm")
