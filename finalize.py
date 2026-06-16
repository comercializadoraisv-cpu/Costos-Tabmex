# -*- coding: utf-8 -*-
"""
Post-procesa el .xlsm generado por build_nomina.py:
  1) Corrige el content-type a macro-habilitado (evita el aviso de extension).
  2) Incrusta el proyecto VBA (xl/vbaProject.bin) con la macro CargarGasto YA
     dentro del libro, de modo que NO haya que importar nada.
  3) Inserta un boton de formulario en la hoja 'Captura' con la macro
     CargarGasto asignada (VML / legacyDrawing).
"""
import zipfile, os, re, shutil

SRC = "Nomina_por_proyecto.xlsm"
VBA_BIN = "vbaProject.bin"
CT_VBA = 'application/vnd.ms-office.vbaProject'
REL_VBA = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'
CT_OLD = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
CT_NEW = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'

VML = '''<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
 <v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe">
  <v:stroke joinstyle="miter"/>
  <v:path shadowok="f" o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
  <o:lock v:ext="edit" shapetype="t"/>
 </v:shapetype>
 <v:shape id="_x0000_s1025" type="#_x0000_t201" style='position:absolute;margin-left:18pt;margin-top:160pt;width:200pt;height:36pt;z-index:1;mso-wrap-style:tight' o:button="t" fillcolor="#1f4e79" strokecolor="#1f4e79" o:insetmode="auto">
  <v:fill color2="#1f4e79" o:detectmouseclick="t"/>
  <o:lock v:ext="edit" rotation="t"/>
  <v:textbox style='mso-direction-alt:auto' o:singleclick="f">
   <div style='text-align:center'><font face="Calibri" size="240" color="#FFFFFF"><b>CARGAR GASTO</b></font></div>
  </v:textbox>
  <x:ClientData ObjectType="Button">
   <x:Anchor>1, 9, 12, 4, 3, 9, 15, 4</x:Anchor>
   <x:PrintObject>False</x:PrintObject>
   <x:AutoFill>False</x:AutoFill>
   <x:FmlaMacro>CargarGasto</x:FmlaMacro>
   <x:TextHAlign>Center</x:TextHAlign>
   <x:TextVAlign>Center</x:TextVAlign>
  </x:ClientData>
 </v:shape>
</xml>'''


def main():
    with zipfile.ZipFile(SRC) as z:
        names = z.namelist()
        data = {n: z.read(n) for n in names}

    # --- 1) content-type macro-habilitado + Default vml + Override vbaProject ---
    ct = data["[Content_Types].xml"].decode("utf-8")
    ct = ct.replace(CT_OLD, CT_NEW)
    if 'Extension="vml"' not in ct:
        ct = ct.replace(
            "</Types>",
            '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>')
    if "vbaProject.bin" not in ct:
        ct = ct.replace(
            "</Types>",
            '<Override PartName="/xl/vbaProject.bin" ContentType="%s"/></Types>' % CT_VBA)
    data["[Content_Types].xml"] = ct.encode("utf-8")

    # --- 1b) incrustar el binario VBA y su relacion en el workbook ---
    with open(VBA_BIN, "rb") as f:
        data["xl/vbaProject.bin"] = f.read()
    wbrels = data["xl/_rels/workbook.xml.rels"].decode("utf-8")
    if "vbaProject.bin" not in wbrels:
        rel_vba = ('<Relationship Id="rIdVbaProj" Type="%s" '
                   'Target="vbaProject.bin"/>' % REL_VBA)
        wbrels = wbrels.replace("</Relationships>", rel_vba + "</Relationships>")
        data["xl/_rels/workbook.xml.rels"] = wbrels.encode("utf-8")

    # --- localizar el sheetN.xml de 'Captura' ---
    wbxml = data["xl/workbook.xml"].decode("utf-8")
    m = re.search(r'<sheet[^>]*name="Captura"[^>]*r:id="(rId\d+)"', wbxml)
    rid = m.group(1)
    rels = data["xl/_rels/workbook.xml.rels"].decode("utf-8")
    # localizar el <Relationship .../> cuyo Id == rid (atributos en cualquier orden)
    rel = next(r for r in re.findall(r'<Relationship\b[^>]*/>', rels)
               if re.search(r'Id="%s"' % rid, r))
    target = re.search(r'Target="([^"]+)"', rel).group(1)   # p.ej. /xl/worksheets/sheet2.xml
    target = target.lstrip("/")                              # xl/worksheets/sheet2.xml
    sheet_path = target if target.startswith("xl/") else "xl/" + target
    sheet_file = os.path.basename(sheet_path)               # sheet2.xml

    # --- 2a) añadir <legacyDrawing> al worksheet (antes de </worksheet>) ---
    sx = data[sheet_path].decode("utf-8")
    # asegurar que el prefijo r: este declarado en el elemento raiz
    if "xmlns:r=" not in sx.split(">", 1)[0]:
        sx = sx.replace(
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
            1)
    if "<legacyDrawing" not in sx:
        sx = sx.replace("</worksheet>", '<legacyDrawing r:id="rIdVml1"/></worksheet>')
    data[sheet_path] = sx.encode("utf-8")

    # --- 2b) rels del worksheet -> vmlDrawing ---
    rels_path = "xl/worksheets/_rels/%s.rels" % sheet_file
    rel_entry = ('<Relationship Id="rIdVml1" '
                 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" '
                 'Target="../drawings/vmlDrawing1.vml"/>')
    if rels_path in data:
        r = data[rels_path].decode("utf-8").replace("</Relationships>", rel_entry + "</Relationships>")
    else:
        r = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             + rel_entry + "</Relationships>")
    data[rels_path] = r.encode("utf-8")

    # --- 2c) el propio VML ---
    data["xl/drawings/vmlDrawing1.vml"] = VML.encode("utf-8")

    # --- reescribir el zip ---
    tmp = SRC + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in data.items():
            z.writestr(n, b)
    os.replace(tmp, SRC)
    print("Finalizado:", SRC)
    print("  hoja Captura:", sheet_path)


if __name__ == "__main__":
    main()
