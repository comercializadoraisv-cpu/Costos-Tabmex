# -*- coding: utf-8 -*-
"""
Post-procesa el .xlsm generado por build_nomina.py:
  1) Corrige el content-type del workbook a macro-habilitado.
  2) Incrusta el proyecto VBA (xl/vbaProject.bin) con la macro CargarGasto.

El boton CARGAR GASTO lo agrega despues add_button.py (via LibreOffice), que
ademas "bendice" el vbaProject.bin y re-inyecta los nombres definidos.
"""
import zipfile, os

SRC = "Nomina_por_proyecto.xlsm"
VBA_BIN = "vbaProject.bin"
CT_OLD = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
CT_NEW = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
CT_VBA = 'application/vnd.ms-office.vbaProject'
REL_VBA = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'


def main():
    with zipfile.ZipFile(SRC) as z:
        data = {n: z.read(n) for n in z.namelist()}

    # --- content-type macro-habilitado + Override vbaProject ---
    ct = data["[Content_Types].xml"].decode("utf-8")
    ct = ct.replace(CT_OLD, CT_NEW)
    if "vbaProject.bin" not in ct:
        ct = ct.replace(
            "</Types>",
            '<Override PartName="/xl/vbaProject.bin" ContentType="%s"/></Types>' % CT_VBA)
    data["[Content_Types].xml"] = ct.encode("utf-8")

    # --- incrustar el binario VBA y su relacion ---
    with open(VBA_BIN, "rb") as f:
        data["xl/vbaProject.bin"] = f.read()
    wbrels = data["xl/_rels/workbook.xml.rels"].decode("utf-8")
    if "vbaProject.bin" not in wbrels:
        rel_vba = ('<Relationship Id="rIdVbaProj" Type="%s" '
                   'Target="vbaProject.bin"/>' % REL_VBA)
        wbrels = wbrels.replace("</Relationships>", rel_vba + "</Relationships>")
        data["xl/_rels/workbook.xml.rels"] = wbrels.encode("utf-8")

    tmp = SRC + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in data.items():
            z.writestr(n, b)
    os.replace(tmp, SRC)
    print("VBA incrustado en", SRC)


if __name__ == "__main__":
    main()
