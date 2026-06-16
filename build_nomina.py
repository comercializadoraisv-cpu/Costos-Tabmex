# -*- coding: utf-8 -*-
"""
Genera el libro de Excel macro-habilitado para captura de gastos de nomina por
proyecto.
  Hojas: Inicio, Captura, Resumen, Catalogos, <Proyectos>, _Movimientos, _Plantilla
  Campos de captura: Fecha, Quincena, Proyecto, Empleado, Puesto/Cargo, Concepto,
                     Sueldo base, Bonos, Otras prestaciones.
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
right = Alignment(horizontal="right", vertical="center")
wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)

QUINCENAS = ["1ra quincena", "2da quincena"]

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
cat.merge_cells("A1:K1")
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

# Puestos / Cargos
cat["G3"] = "Puestos / Cargos"
cat["G3"].font = sub_font
cat["G3"].fill = header_fill
puestos = ["(captura aqui tus puestos)"]
for i, pu in enumerate(puestos):
    c = cat.cell(row=4 + i, column=7, value=pu)
    c.border = border
    c.font = normal_font

# Parametros: factor de carga social
cat["I3"] = "Parametros"
cat["I3"].font = sub_font
cat["I3"].fill = header_fill
cat["I4"] = "Factor de carga social patronal (%)"
cat["I4"].font = label_font
cat["I4"].alignment = wrap
cat["J4"] = 0.0
cat["J4"].number_format = pct
cat["J4"].fill = input_fill
cat["J4"].border = border
cat["J4"].font = normal_font
cat["I6"] = ("Estimado de IMSS, Infonavit y provisiones (aguinaldo, prima "
             "vacacional). En 0% solo se cuenta lo pagado en mano. Subelo para "
             "reflejar el costo patronal real por proyecto.")
cat["I6"].font = nota_font
cat["I6"].alignment = wrap
cat.merge_cells("I6:J9")

anchos = {"A": 20, "B": 3, "C": 28, "D": 3, "E": 24, "F": 3,
          "G": 26, "H": 3, "I": 24, "J": 14}
for col, w in anchos.items():
    cat.column_dimensions[col].width = w

# Rangos con nombre
wb.defined_names.add(DefinedName("FactorCargaSocial", attr_text="Catalogos!$J$4"))
wb.defined_names.add(DefinedName("ListaProyectos", attr_text="Catalogos!$A$4:$A$203"))
wb.defined_names.add(DefinedName("ListaEmpleados", attr_text="Catalogos!$C$4:$C$203"))
wb.defined_names.add(DefinedName("ListaConceptos", attr_text="Catalogos!$E$4:$E$23"))
wb.defined_names.add(DefinedName("ListaPuestos", attr_text="Catalogos!$G$4:$G$203"))

# =====================================================================
# _MOVIMIENTOS  (libro maestro / auditoria)
# =====================================================================
mov = wb.create_sheet("_Movimientos")
mov_headers = ["Fecha", "Quincena", "Proyecto", "Empleado", "Puesto", "Concepto",
               "Sueldo", "Bonos", "Otras prestaciones", "Total percibido"]
for j, h in enumerate(mov_headers, start=1):
    c = mov.cell(row=1, column=j, value=h)
    c.font = sub_font
    c.fill = header_fill
    c.alignment = center
    c.border = border
widths = [12, 14, 18, 24, 18, 22, 14, 14, 16, 16]
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
    ws.merge_cells("A1:J1")
    ws["A1"].alignment = center

    ws["A2"] = "GASTO DE NOMINA POR PROYECTO"
    ws["A2"].font = label_font
    ws.merge_cells("A2:J2")

    # Totales (arriba de la tabla)
    ws.merge_cells("G3:H3")
    ws["G3"] = "Total percibido:"
    ws["G3"].font = label_font
    ws["G3"].alignment = right
    ws["I3"] = "=SUM(I6:I100000)"
    ws["I3"].number_format = money
    ws["I3"].font = Font(bold=True, color=VERDE)

    ws.merge_cells("G4:H4")
    ws["G4"] = "Costo total c/carga social:"
    ws["G4"].font = label_font
    ws["G4"].alignment = right
    ws["J4"] = "=SUM(J6:J100000)"
    ws["J4"].number_format = money
    ws["J4"].font = Font(bold=True, color=VERDE)

    headers = ["Fecha", "Quincena", "Empleado", "Puesto", "Concepto",
               "Sueldo", "Bonos", "Otras prestaciones",
               "Total percibido", "Costo c/carga social"]
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=5, column=j, value=h)
        c.font = sub_font
        c.fill = header_fill
        c.alignment = center
        c.border = border
    ws_widths = [12, 14, 24, 20, 22, 13, 13, 16, 16, 18]
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
cap = wb.create_sheet("Captura", 0)
cap.sheet_view.showGridLines = False

cap["B2"] = "CAPTURA DE GASTO DE NOMINA"
cap["B2"].font = titulo_font
cap["B2"].fill = header_fill
cap.merge_cells("B2:D2")
cap["B2"].alignment = center

campos = [
    (4, "Fecha", "C4", "fecha"),
    (5, "Quincena", "C5", "texto"),
    (6, "Proyecto", "C6", "texto"),
    (7, "Empleado", "C7", "texto"),
    (8, "Puesto / Cargo", "C8", "texto"),
    (9, "Concepto", "C9", "texto"),
    (11, "Sueldo base", "C11", "money"),
    (12, "Bonos", "C12", "money"),
    (13, "Otras prestaciones", "C13", "money"),
]
for row, etiqueta, celda, tipo in campos:
    lc = cap.cell(row=row, column=2, value=etiqueta)
    lc.font = label_font
    lc.alignment = right
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
cap["B16"].alignment = right
cap["C16"] = "=N(C11)+N(C12)+N(C13)"
cap["C16"].number_format = money
cap["C16"].font = Font(bold=True, color=VERDE, size=12)
cap["C16"].fill = light_fill
cap["C16"].border = border

cap["B19"] = ("C16 es solo vista previa del registro; no es un boton. "
              "Llena los campos y pulsa el boton CARGAR GASTO (o Alt+F8 -> CargarGasto).")
cap["B19"].font = nota_font
cap["B19"].alignment = wrap
cap.merge_cells("B19:D21")

cap.column_dimensions["A"].width = 3
cap.column_dimensions["B"].width = 22
cap.column_dimensions["C"].width = 26
cap.column_dimensions["D"].width = 6

# Validaciones (dropdowns)
dv_quin = DataValidation(type="list", formula1='"%s"' % ",".join(QUINCENAS),
                         allow_blank=True)
dv_proj = DataValidation(type="list", formula1="=ListaProyectos", allow_blank=True)
dv_emp = DataValidation(type="list", formula1="=ListaEmpleados", allow_blank=True)
dv_pue = DataValidation(type="list", formula1="=ListaPuestos", allow_blank=True)
dv_con = DataValidation(type="list", formula1="=ListaConceptos", allow_blank=True)
cap.add_data_validation(dv_quin); dv_quin.add(cap["C5"])
cap.add_data_validation(dv_proj); dv_proj.add(cap["C6"])
cap.add_data_validation(dv_emp); dv_emp.add(cap["C7"])
cap.add_data_validation(dv_pue); dv_pue.add(cap["C8"])
cap.add_data_validation(dv_con); dv_con.add(cap["C9"])

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

first = 4
last = first + 49
for i in range(50):
    r = first + i
    proj_ref = f"Catalogos!$A${4 + i}"
    res.cell(row=r, column=1, value=f"={proj_ref}").border = border
    # Total percibido: suma columna J de _Movimientos donde Proyecto (col C) = proyecto
    tp = (f"=IF({proj_ref}=\"\",\"\","
          f"SUMIFS('_Movimientos'!$J:$J,'_Movimientos'!$C:$C,{proj_ref}))")
    c2 = res.cell(row=r, column=2, value=tp)
    c2.number_format = money; c2.border = border
    cs = (f"=IF({proj_ref}=\"\",\"\",B{r}*FactorCargaSocial)")
    c3 = res.cell(row=r, column=3, value=cs)
    c3.number_format = money; c3.border = border
    ct = (f"=IF({proj_ref}=\"\",\"\",B{r}+C{r})")
    c4 = res.cell(row=r, column=4, value=ct)
    c4.number_format = money; c4.border = border
    c4.font = Font(bold=True)

res.cell(row=last + 1, column=1, value="TOTAL GENERAL").font = Font(bold=True, color=AZUL)
for col in (2, 3, 4):
    L = get_column_letter(col)
    c = res.cell(row=last + 1, column=col, value=f"=SUM({L}{first}:{L}{last})")
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
    "2. La macro YA viene incrustada en este libro: no hay que importar nada ni entrar al editor.",
    "3. Ve a la hoja CAPTURA y llena: Fecha, Quincena, Proyecto, Empleado, Puesto, Concepto, Sueldo, Bonos, Otras prestaciones.",
    "4. Pulsa el boton CARGAR GASTO (o Alt+F8 -> CargarGasto -> Ejecutar).",
    "5. El registro se manda a la hoja del proyecto (la crea si no existe) y al libro _Movimientos.",
    "6. RESUMEN y los totales por proyecto se actualizan solos (formulas SUMIFS).",
    "",
    "CATALOGOS: edita aqui tus Proyectos, Empleados, Puestos, Conceptos y el Factor de carga social patronal (%).",
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
