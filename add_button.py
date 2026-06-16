# -*- coding: utf-8 -*-
"""
Agrega el boton CARGAR GASTO al .xlsm y deja el archivo listo para Excel.

Requiere una instancia de LibreOffice escuchando en el socket 2002:
  soffice --headless --invisible --norestore --nodefault \
    -env:UserInstallation=file:///tmp/louser \
    --accept="socket,host=localhost,port=2002;urp;"

Pasos:
  1) LibreOffice abre el libro (con el VBA ya incrustado), agrega un boton de
     control de formulario en la hoja Captura y lo guarda. Al guardar, LO:
       - re-serializa (bendice) el vbaProject.bin con sus claves CMG/DPB/GC,
       - escribe la estructura COMPLETA del boton (drawing + ctrlProps +
         controls), que es lo que Excel necesita para que NO salga gris,
       - PERO pierde los nombres definidos y NO escribe el vinculo a la macro.
  2) Reparacion por edicion del zip:
       - re-inyecta los nombres definidos (FactorCargaSocial, Lista*...),
       - inyecta <x:FmlaMacro>CargarGasto</x:FmlaMacro> en el VML del boton.
"""
import sys, time, zipfile, os, re

SRC = "Nomina_por_proyecto.xlsm"
MODULE = "modNomina"     # modulo VBA
SUB = "CargarGasto"      # procedimiento que ejecuta el boton (FmlaMacro)
BTN_NAME = "btnCargar"

# nombres definidos que hay que restituir (deben coincidir con build_nomina.py)
DEFINED_NAMES = [
    ("FactorCargaSocial", "Catalogos!$J$4"),
    ("ListaProyectos", "Catalogos!$A$4:$A$203"),
    ("ListaEmpleados", "Catalogos!$C$4:$C$203"),
    ("ListaConceptos", "Catalogos!$E$4:$E$23"),
    ("ListaPuestos", "Catalogos!$G$4:$G$203"),
]


def lo_add_button(path):
    import uno
    from com.sun.star.beans import PropertyValue
    from com.sun.star.awt import Point, Size
    from com.sun.star.script import ScriptEventDescriptor

    def mk(n, v):
        p = PropertyValue(); p.Name = n; p.Value = v; return p

    lc = uno.getComponentContext()
    res = lc.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", lc)
    ctx = None
    for _ in range(40):
        try:
            ctx = res.resolve(
                "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
            break
        except Exception:
            time.sleep(1)
    if ctx is None:
        raise RuntimeError("No pude conectar con LibreOffice (socket 2002).")
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    url = "file://" + os.path.abspath(path)
    doc = desktop.loadComponentFromURL(
        url, "_blank", 0, (mk("Hidden", True),
                           mk("FilterName", "Calc MS Excel 2007 VBA XML")))
    cap = doc.Sheets.getByName("Captura")
    dp = cap.DrawPage
    while dp.Count > 0:
        dp.remove(dp.getByIndex(0))

    # posicionar el boton anclado a la celda B23 (debajo del formulario)
    anchor = cap.getCellByPosition(1, 22)   # col B, fila 23 (0-based)
    pos = anchor.Position
    shape = doc.createInstance("com.sun.star.drawing.ControlShape")
    shape.setPosition(Point(pos.X, pos.Y))
    shape.setSize(Size(5200, 1000))
    model = smgr.createInstanceWithContext(
        "com.sun.star.form.component.CommandButton", ctx)
    model.Label = "CARGAR GASTO"
    model.Name = BTN_NAME
    shape.setControl(model)
    dp.add(shape)

    forms = dp.Forms
    form = forms.getByIndex(0)
    idx = [i for i in range(form.Count)
           if form.getByIndex(i).Name == BTN_NAME][0]
    sed = ScriptEventDescriptor()
    sed.ListenerType = "XActionListener"
    sed.EventMethod = "actionPerformed"
    sed.ScriptType = "Script"
    sed.ScriptCode = ("vnd.sun.star.script:VBAProject.%s.%s"
                      "?language=Basic&location=document" % (MODULE, SUB))
    forms.registerScriptEvent(idx, sed)

    doc.storeToURL(url, (mk("FilterName", "Calc MS Excel 2007 VBA XML"),))
    doc.close(False)


def repair(path):
    with zipfile.ZipFile(path) as z:
        data = {n: z.read(n) for n in z.namelist()}

    # --- 1) restituir nombres definidos en workbook.xml ---
    wbxml = data["xl/workbook.xml"].decode("utf-8")
    wbxml = re.sub(r"<definedNames>.*?</definedNames>", "", wbxml, flags=re.S)
    dn = "<definedNames>" + "".join(
        '<definedName name="%s">%s</definedName>' % (n, ref)
        for n, ref in DEFINED_NAMES) + "</definedNames>"
    # los definedNames van despues de </sheets>
    wbxml = wbxml.replace("</sheets>", "</sheets>" + dn, 1)
    data["xl/workbook.xml"] = wbxml.encode("utf-8")

    # --- 2) inyectar FmlaMacro en el VML del boton ---
    vml_key = next((n for n in data if re.match(r"xl/drawings/vmlDrawing\d+\.vml$", n)), None)
    if vml_key is None:
        raise RuntimeError("No encontre el vmlDrawing del boton.")
    vml = data[vml_key].decode("utf-8")
    if "FmlaMacro" not in vml:
        fm = "<x:FmlaMacro>%s</x:FmlaMacro>" % SUB
        if "</x:Anchor>" in vml:
            vml = vml.replace("</x:Anchor>", "</x:Anchor>" + fm, 1)
        else:
            vml = vml.replace("</x:ClientData>", fm + "</x:ClientData>", 1)
    data[vml_key] = vml.encode("utf-8")

    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in data.items():
            z.writestr(n, b)
    os.replace(tmp, path)


def main():
    lo_add_button(SRC)
    repair(SRC)
    print("Boton agregado y archivo reparado:", SRC)


if __name__ == "__main__":
    main()
