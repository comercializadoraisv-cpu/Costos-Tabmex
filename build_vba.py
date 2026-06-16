# -*- coding: utf-8 -*-
"""
Genera un vbaProject.bin valido (MS-OVBA) que contiene el modulo estandar
'CargarGasto', SIN depender de Excel ni de LibreOffice para escribirlo.

Implementa:
  * Compresion/Decompresion MS-OVBA (2.4.1)   -> compress() / decompress()
  * Streams del proyecto VBA (dir, _VBA_PROJECT, PROJECT, PROJECTwm, modulo)
  * Escritor de Compound File Binary (OLE2/CFB) -> write_cfb()

Verificacion: el binario se vuelve a leer con olefile y la fuente se
descomprime para comprobar que coincide byte a byte.
"""
import struct, uuid

# =====================================================================
#  Codigo VBA del modulo (endurecido: tipos Object + constantes literales
#  de Excel, para que compile aunque el host inyecte solo la libreria base)
# =====================================================================
VBA_SOURCE = '''Attribute VB_Name = "modNomina"
'======================================================================
'  CargarGasto  -  Captura de gastos de nomina por proyecto
'----------------------------------------------------------------------
'  Lee el formulario de la hoja "Captura", registra el movimiento en
'  el libro maestro "_Movimientos", lo manda a la hoja del proyecto
'  (creandola desde "_Plantilla" si no existe) y limpia el formulario.
'  Los totales del Resumen y de cada proyecto son formulas: se solos.
'======================================================================
Option Explicit

' --- Constantes de Excel como literales (no dependen de la libreria) ---
Private Const xlUp As Long = -4162
Private Const xlSheetVisible As Long = -1
Private Const xlSheetHidden As Long = 0

' --- Celdas del formulario en la hoja Captura ---
Private Const CEL_FECHA    As String = "C4"
Private Const CEL_QUINCENA As String = "C5"
Private Const CEL_PROYECTO As String = "C6"
Private Const CEL_EMPLEADO As String = "C7"
Private Const CEL_PUESTO   As String = "C8"
Private Const CEL_CONCEPTO As String = "C9"
Private Const CEL_SUELDO   As String = "C11"
Private Const CEL_BONOS    As String = "C12"
Private Const CEL_OTRAS    As String = "C13"

Public Sub CargarGasto()
    Dim wb As Object: Set wb = ThisWorkbook
    Dim cap As Object
    On Error Resume Next
    Set cap = wb.Worksheets("Captura")
    On Error GoTo 0
    If cap Is Nothing Then
        MsgBox "No encuentro la hoja 'Captura'.", vbCritical, "CargarGasto"
        Exit Sub
    End If

    ' --- Leer formulario ---
    Dim fecha As Variant, quincena As String, proyecto As String
    Dim empleado As String, puesto As String, concepto As String
    Dim sueldo As Double, bonos As Double, otras As Double, total As Double

    fecha = cap.Range(CEL_FECHA).Value
    quincena = Trim$(CStr(cap.Range(CEL_QUINCENA).Value))
    proyecto = Trim$(CStr(cap.Range(CEL_PROYECTO).Value))
    empleado = Trim$(CStr(cap.Range(CEL_EMPLEADO).Value))
    puesto = Trim$(CStr(cap.Range(CEL_PUESTO).Value))
    concepto = Trim$(CStr(cap.Range(CEL_CONCEPTO).Value))
    sueldo = ToNum(cap.Range(CEL_SUELDO).Value)
    bonos = ToNum(cap.Range(CEL_BONOS).Value)
    otras = ToNum(cap.Range(CEL_OTRAS).Value)
    total = sueldo + bonos + otras

    ' --- Validaciones ---
    If Len(proyecto) = 0 Then
        MsgBox "Selecciona o escribe un PROYECTO antes de cargar.", vbExclamation, "CargarGasto"
        cap.Range(CEL_PROYECTO).Select
        Exit Sub
    End If
    If total <= 0 Then
        MsgBox "El total del registro es 0. Captura Sueldo, Bonos u Otras prestaciones.", _
               vbExclamation, "CargarGasto"
        cap.Range(CEL_SUELDO).Select
        Exit Sub
    End If
    If IsEmpty(fecha) Or Len(Trim$(CStr(fecha))) = 0 Then fecha = Date

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' --- 1) Registrar en el libro maestro _Movimientos ---
    Dim mov As Object: Set mov = wb.Worksheets("_Movimientos")
    Dim rm As Long
    rm = mov.Cells(mov.Rows.Count, "A").End(xlUp).Row + 1
    If rm < 2 Then rm = 2
    mov.Cells(rm, 1).Value = fecha
    mov.Cells(rm, 1).NumberFormat = "dd/mm/yyyy"
    mov.Cells(rm, 2).Value = quincena
    mov.Cells(rm, 3).Value = proyecto
    mov.Cells(rm, 4).Value = empleado
    mov.Cells(rm, 5).Value = puesto
    mov.Cells(rm, 6).Value = concepto
    mov.Cells(rm, 7).Value = sueldo
    mov.Cells(rm, 8).Value = bonos
    mov.Cells(rm, 9).Value = otras
    mov.Cells(rm, 10).Value = total

    ' --- 2) Hoja del proyecto (crear si no existe) ---
    Dim ws As Object: Set ws = ObtenerHojaProyecto(wb, proyecto)
    If ws Is Nothing Then
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "No pude crear la hoja del proyecto '" & proyecto & "'.", vbCritical, "CargarGasto"
        Exit Sub
    End If

    ' --- 3) Agregar fila en la hoja del proyecto (datos desde fila 6) ---
    Dim rp As Long
    rp = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If rp < 5 Then rp = 5
    rp = rp + 1
    ws.Cells(rp, 1).Value = fecha
    ws.Cells(rp, 1).NumberFormat = "dd/mm/yyyy"
    ws.Cells(rp, 2).Value = quincena
    ws.Cells(rp, 3).Value = empleado
    ws.Cells(rp, 4).Value = puesto
    ws.Cells(rp, 5).Value = concepto
    ws.Cells(rp, 6).Value = sueldo
    ws.Cells(rp, 7).Value = bonos
    ws.Cells(rp, 8).Value = otras
    ws.Cells(rp, 9).Value = total
    ws.Cells(rp, 9).NumberFormat = "#,##0.00"
    ws.Cells(rp, 10).Formula = "=I" & rp & "*(1+FactorCargaSocial)"
    ws.Cells(rp, 10).NumberFormat = "#,##0.00"

    ' --- 4) Asegurar que el proyecto este en el catalogo ---
    AgregarProyectoCatalogo wb, proyecto

    ' --- 5) Limpiar formulario ---
    cap.Range(CEL_QUINCENA & "," & CEL_PROYECTO & "," & CEL_EMPLEADO & "," & _
              CEL_PUESTO & "," & CEL_CONCEPTO & "," & CEL_SUELDO & "," & _
              CEL_BONOS & "," & CEL_OTRAS).ClearContents
    cap.Range(CEL_FECHA).ClearContents

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Registro cargado en el proyecto '" & proyecto & "'." & vbCrLf & _
           "Total percibido: " & Format$(total, "#,##0.00") & " MXN", _
           vbInformation, "CargarGasto"
    cap.Range(CEL_FECHA).Select
End Sub

'----------------------------------------------------------------------
' Devuelve la hoja del proyecto; si no existe la crea desde _Plantilla.
'----------------------------------------------------------------------
Private Function ObtenerHojaProyecto(wb As Object, ByVal proyecto As String) As Object
    Dim nombre As String: nombre = NombreHojaValido(proyecto)
    Dim ws As Object
    On Error Resume Next
    Set ws = wb.Worksheets(nombre)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Set ObtenerHojaProyecto = ws
        Exit Function
    End If

    ' Crear copiando la plantilla oculta
    Dim plant As Object
    On Error Resume Next
    Set plant = wb.Worksheets("_Plantilla")
    On Error GoTo 0
    If plant Is Nothing Then Exit Function

    plant.Visible = xlSheetVisible
    plant.Copy After:=wb.Worksheets(wb.Worksheets.Count)
    Dim nueva As Object: Set nueva = wb.Worksheets(wb.Worksheets.Count)
    nueva.Name = nombre
    nueva.Visible = xlSheetVisible
    nueva.Range("A1").Value = proyecto
    plant.Visible = xlSheetHidden
    Set ObtenerHojaProyecto = nueva
End Function

'----------------------------------------------------------------------
' Agrega el proyecto a Catalogos!A4:A... si aun no esta en la lista.
'----------------------------------------------------------------------
Private Sub AgregarProyectoCatalogo(wb As Object, ByVal proyecto As String)
    Dim cat As Object: Set cat = wb.Worksheets("Catalogos")
    Dim ult As Long, r As Long
    ult = cat.Cells(cat.Rows.Count, "A").End(xlUp).Row
    If ult < 4 Then ult = 3
    For r = 4 To ult
        If StrComp(Trim$(CStr(cat.Cells(r, 1).Value)), proyecto, vbTextCompare) = 0 Then
            Exit Sub  ' ya existe
        End If
    Next r
    cat.Cells(ult + 1, 1).Value = proyecto
End Sub

'----------------------------------------------------------------------
' Limpia un texto para usarlo como nombre de hoja valido en Excel.
'----------------------------------------------------------------------
Private Function NombreHojaValido(ByVal s As String) As String
    Dim malo As Variant, ch As Variant
    malo = Array(":", "\\", "/", "?", "*", "[", "]")
    For Each ch In malo
        s = Replace(s, CStr(ch), " ")
    Next ch
    s = Trim$(s)
    If Len(s) > 31 Then s = Left$(s, 31)
    If Len(s) = 0 Then s = "Proyecto"
    NombreHojaValido = s
End Function

'----------------------------------------------------------------------
' Convierte a numero de forma segura (celdas vacias o texto -> 0).
'----------------------------------------------------------------------
Private Function ToNum(ByVal v As Variant) As Double
    If IsNumeric(v) Then
        ToNum = CDbl(v)
    Else
        ToNum = 0
    End If
End Function
'''

MODULE_NAME = "modNomina"     # nombre del MODULO (distinto del Sub CargarGasto)
SUB_NAME = "CargarGasto"      # nombre del procedimiento que ejecuta el boton
PROJECT_NAME = "VBAProject"

# =====================================================================
#  MS-OVBA  compression / decompression  (2.4.1)
# =====================================================================
def _compress_chunk(chunk):
    out = bytearray()
    pos = 0
    n = len(chunk)
    while pos < n:
        flag_idx = len(out)
        out.append(0)
        flag = 0
        for bit in range(8):
            if pos >= n:
                break
            diff = pos
            bit_count = max((diff - 1).bit_length(), 4) if diff > 0 else 4
            max_len = (0xFFFF >> bit_count) + 3
            best_len, best_off = 0, 0
            limit = min(max_len, n - pos)
            # buscar la coincidencia mas larga en la ventana [0, pos)
            cand = pos - 1
            while cand >= 0:
                off = pos - cand
                l = 0
                while l < limit and chunk[cand + l] == chunk[pos + l]:
                    l += 1
                if l > best_len:
                    best_len, best_off = l, off
                    if l == limit:
                        break
                cand -= 1
            if best_len >= 3:
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                out += struct.pack('<H', token)
                flag |= (1 << bit)
                pos += best_len
            else:
                out.append(chunk[pos])
                pos += 1
        out[flag_idx] = flag
    return bytes(out)


def compress(data):
    out = bytearray([0x01])  # SignatureByte
    i = 0
    while i < len(data):
        chunk = data[i:i + 4096]
        i += 4096
        comp = _compress_chunk(chunk)
        if len(comp) < len(chunk) and len(comp) <= 4096:
            size12 = (2 + len(comp)) - 3
            header = 0xB000 | (size12 & 0x0FFF)  # flag=1, sig=0b011
            out += struct.pack('<H', header) + comp
        else:
            # chunk sin comprimir: exactamente 4096 bytes
            raw = chunk + b'\x00' * (4096 - len(chunk))
            header = 0x3000 | 0x0FFF  # flag=0, sig=0b011, size=4098-3
            out += struct.pack('<H', header) + raw
    return bytes(out)


def decompress(data):
    assert data[0] == 0x01, "firma invalida"
    out = bytearray()
    i = 1
    while i < len(data):
        header = struct.unpack('<H', data[i:i + 2])[0]
        i += 2
        size = (header & 0x0FFF) + 3
        flag = (header >> 15) & 1
        end = i + size - 2
        if flag == 0:
            out += data[i:i + 4096]
            i += 4096
        else:
            chunk_start = len(out)
            while i < end:
                fb = data[i]; i += 1
                for bit in range(8):
                    if i >= end:
                        break
                    if (fb >> bit) & 1:
                        token = struct.unpack('<H', data[i:i + 2])[0]; i += 2
                        diff = len(out) - chunk_start
                        bit_count = max((diff - 1).bit_length(), 4) if diff > 0 else 4
                        len_mask = 0xFFFF >> bit_count
                        off_mask = (~len_mask) & 0xFFFF
                        length = (token & len_mask) + 3
                        off = ((token & off_mask) >> (16 - bit_count)) + 1
                        src = len(out) - off
                        for k in range(length):
                            out.append(out[src + k])
                    else:
                        out.append(data[i]); i += 1
    return bytes(out)


# =====================================================================
#  Streams del proyecto VBA
# =====================================================================
def _rec(idv, payload):
    return struct.pack('<HI', idv, len(payload)) + payload


def build_dir_stream():
    p = bytearray()
    # ---- PROJECTINFORMATION ----
    p += _rec(0x0001, struct.pack('<I', 0x00000001))   # SYSKIND = Win32
    p += _rec(0x0002, struct.pack('<I', 0x00000409))   # LCID
    p += _rec(0x0014, struct.pack('<I', 0x00000409))   # LCIDINVOKE
    p += _rec(0x0003, struct.pack('<H', 0x04E4))       # CODEPAGE = 1252
    p += _rec(0x0004, PROJECT_NAME.encode('ascii'))    # NAME
    # DOCSTRING
    p += struct.pack('<H', 0x0005) + struct.pack('<I', 0) + struct.pack('<H', 0x0040) + struct.pack('<I', 0)
    # HELPFILEPATH
    p += struct.pack('<H', 0x0006) + struct.pack('<I', 0) + struct.pack('<H', 0x003D) + struct.pack('<I', 0)
    p += _rec(0x0007, struct.pack('<I', 0))            # HELPCONTEXT
    p += _rec(0x0008, struct.pack('<I', 0))            # LIBFLAGS
    # VERSION: Reserved(=4) + Major(4) + Minor(2)
    p += struct.pack('<H', 0x0009) + struct.pack('<I', 4) + struct.pack('<I', 1) + struct.pack('<H', 0)
    # CONSTANTS
    p += struct.pack('<H', 0x000C) + struct.pack('<I', 0) + struct.pack('<H', 0x003C) + struct.pack('<I', 0)

    # ---- PROJECTREFERENCES: ninguna ----

    # ---- PROJECTMODULES ----
    p += struct.pack('<H', 0x000F) + struct.pack('<I', 2) + struct.pack('<H', 1)        # modules count
    p += struct.pack('<H', 0x0013) + struct.pack('<I', 2) + struct.pack('<H', 0xFFFF)   # cookie

    mname = MODULE_NAME.encode('ascii')
    mnameu = MODULE_NAME.encode('utf-16-le')
    p += _rec(0x0019, mname)                            # MODULENAME
    p += _rec(0x0047, mnameu)                           # MODULENAMEUNICODE
    # MODULESTREAMNAME
    p += struct.pack('<H', 0x001A) + struct.pack('<I', len(mname)) + mname
    p += struct.pack('<H', 0x0032) + struct.pack('<I', len(mnameu)) + mnameu
    # MODULEDOCSTRING
    p += struct.pack('<H', 0x001C) + struct.pack('<I', 0) + struct.pack('<H', 0x0048) + struct.pack('<I', 0)
    p += _rec(0x0031, struct.pack('<I', 0))            # MODULEOFFSET = 0
    p += _rec(0x001E, struct.pack('<I', 0))            # MODULEHELPCONTEXT
    p += _rec(0x002C, struct.pack('<H', 0xFFFF))       # MODULECOOKIE
    p += struct.pack('<H', 0x0021) + struct.pack('<I', 0)   # MODULETYPE procedural
    p += struct.pack('<H', 0x002B) + struct.pack('<I', 0)   # MODULE terminator

    # ---- Terminator + Reserved ----
    p += struct.pack('<H', 0x0010) + struct.pack('<I', 0) + struct.pack('<I', 0)
    return bytes(p)


def build_project_stream():
    guid = "{%s}" % str(uuid.uuid4()).upper()
    lines = [
        'ID="%s"' % guid,
        'Module=%s' % MODULE_NAME,
        'Name="%s"' % PROJECT_NAME,
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        '',
        '[Host Extender Info]',
        '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000',
        '',
    ]
    return ("\r\n".join(lines)).encode('ascii')


def build_projectwm_stream():
    out = bytearray()
    out += MODULE_NAME.encode('ascii') + b'\x00'
    out += MODULE_NAME.encode('utf-16-le') + b'\x00\x00'
    out += b'\x00\x00'
    return bytes(out)


def build_streams():
    src = VBA_SOURCE.replace('\n', '\r\n').encode('ascii')
    module_stream = compress(src)
    # verificacion de round-trip de la fuente
    assert decompress(module_stream) == src, "round-trip de la fuente fallo"

    dir_raw = build_dir_stream()
    dir_stream = compress(dir_raw)
    assert decompress(dir_stream) == dir_raw, "round-trip de dir fallo"

    vba_project = b'\xCC\x61\xFF\xFF\x00\x00\x00'

    return {
        'PROJECT': build_project_stream(),
        'PROJECTwm': build_projectwm_stream(),
        'VBA/_VBA_PROJECT': vba_project,
        'VBA/dir': dir_stream,
        'VBA/' + MODULE_NAME: module_stream,
    }, src


# =====================================================================
#  Escritor de Compound File Binary (OLE2 / CFB)   (MS-CFB)
# =====================================================================
SECTOR = 512
MINISECTOR = 64
FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF
MINI_CUTOFF = 4096


def _dir_entry(name, etype, color, left, right, child, start, size):
    nb = name.encode('utf-16-le') + b'\x00\x00'
    if len(nb) > 64:
        raise ValueError("nombre demasiado largo: " + name)
    nb = nb + b'\x00' * (64 - len(nb))
    e = nb
    e += struct.pack('<H', (len(name) + 1) * 2)   # name length incl. null
    e += struct.pack('<B', etype)
    e += struct.pack('<B', color)
    e += struct.pack('<I', left)
    e += struct.pack('<I', right)
    e += struct.pack('<I', child)
    e += b'\x00' * 16            # CLSID
    e += struct.pack('<I', 0)    # state bits
    e += struct.pack('<Q', 0)    # create time
    e += struct.pack('<Q', 0)    # modify time
    e += struct.pack('<I', start)
    e += struct.pack('<Q', size)
    assert len(e) == 128
    return e


def write_cfb(streams, path):
    """streams: dict ruta->bytes.  Crea el arbol fijo del proyecto VBA."""
    PROJECT = streams['PROJECT']
    PROJECTwm = streams['PROJECTwm']
    VBA_PROJECT = streams['VBA/_VBA_PROJECT']
    DIRS = streams['VBA/dir']
    MODULE = streams['VBA/' + MODULE_NAME]

    # Todos nuestros streams son < 4096 -> mini stream
    for nm, b in streams.items():
        assert len(b) < MINI_CUTOFF, "stream grande no soportado: " + nm

    # --- 1) Construir el mini stream y la miniFAT ---
    # orden en el mini stream: PROJECT, PROJECTwm, _VBA_PROJECT, dir, modulo
    mini_layout = [
        ('PROJECT', PROJECT),
        ('PROJECTwm', PROJECTwm),
        ('_VBA_PROJECT', VBA_PROJECT),
        ('dir', DIRS),
        (MODULE_NAME, MODULE),
    ]
    ministream = bytearray()
    minifat = []
    mini_start = {}
    mini_size = {}
    for nm, b in mini_layout:
        start = len(minifat)
        nsec = max(1, (len(b) + MINISECTOR - 1) // MINISECTOR)
        for k in range(nsec):
            minifat.append(len(minifat) + 1 if k < nsec - 1 else ENDOFCHAIN)
        chunk = b + b'\x00' * (nsec * MINISECTOR - len(b))
        ministream += chunk
        mini_start[nm] = start
        mini_size[nm] = len(b)
    # pad ministream a multiplo de sector
    if len(ministream) % SECTOR:
        ministream += b'\x00' * (SECTOR - len(ministream) % SECTOR)

    # --- 2) Directorio ---
    # indices: 0 Root, 1 PROJECT, 2 PROJECTwm, 3 VBA, 4 _VBA_PROJECT, 5 dir, 6 modulo
    entries = []
    # Root: child -> PROJECT(1).  start = primer sector del ministream (se fija luego)
    entries.append(['Root Entry', 5, 1, NOSTREAM, NOSTREAM, 1, None, len(ministream)])
    entries.append([ 'PROJECT', 2, 1, 3, 2, NOSTREAM, mini_start['PROJECT'], mini_size['PROJECT']])
    entries.append([ 'PROJECTwm', 2, 1, NOSTREAM, NOSTREAM, NOSTREAM, mini_start['PROJECTwm'], mini_size['PROJECTwm']])
    entries.append([ 'VBA', 1, 1, NOSTREAM, NOSTREAM, 6, 0, 0])
    entries.append([ '_VBA_PROJECT', 2, 1, NOSTREAM, NOSTREAM, NOSTREAM, mini_start['_VBA_PROJECT'], mini_size['_VBA_PROJECT']])
    entries.append([ 'dir', 2, 1, NOSTREAM, NOSTREAM, NOSTREAM, mini_start['dir'], mini_size['dir']])
    entries.append([ MODULE_NAME, 2, 1, 5, 4, NOSTREAM, mini_start[MODULE_NAME], mini_size[MODULE_NAME]])

    # --- 3) Disposicion de sectores grandes ---
    # [ dir-stream ][ miniFAT ][ ministream ][ FAT ]
    dir_bytes_count = len(entries) * 128
    # pad directorio a sector
    dir_sectors = (dir_bytes_count + SECTOR - 1) // SECTOR

    minifat_bytes = b''.join(struct.pack('<I', x) for x in minifat)
    if len(minifat_bytes) % SECTOR:
        minifat_bytes += b'\xFF' * (SECTOR - len(minifat_bytes) % SECTOR)
    minifat_sectors = len(minifat_bytes) // SECTOR

    ministream_sectors = len(ministream) // SECTOR

    non_fat_sectors = dir_sectors + minifat_sectors + ministream_sectors

    # numero de sectores FAT: iterar hasta punto fijo
    fat_sectors = 1
    while True:
        total = non_fat_sectors + fat_sectors
        needed = (total * 4 + SECTOR - 1) // SECTOR
        if needed <= fat_sectors:
            break
        fat_sectors = needed

    total_sectors = non_fat_sectors + fat_sectors

    # asignacion de indices de sector
    s = 0
    dir_first = s; s += dir_sectors
    minifat_first = s; s += minifat_sectors
    ministream_first = s; s += ministream_sectors
    fat_first = s; s += fat_sectors
    assert s == total_sectors

    # Root entry apunta al primer sector del ministream
    entries[0][6] = ministream_first

    # --- 4) Construir la FAT ---
    fat = [FREESECT] * total_sectors

    def chain(first, count):
        for k in range(count):
            sec = first + k
            fat[sec] = (first + k + 1) if k < count - 1 else ENDOFCHAIN

    chain(dir_first, dir_sectors)
    chain(minifat_first, minifat_sectors)
    chain(ministream_first, ministream_sectors)
    for k in range(fat_sectors):
        fat[fat_first + k] = FATSECT

    fat_bytes = b''.join(struct.pack('<I', x) for x in fat)
    if len(fat_bytes) % SECTOR:
        fat_bytes += b'\xFF' * (SECTOR - len(fat_bytes) % SECTOR)

    # --- 5) Encabezado ---
    header = bytearray()
    header += b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'   # signature
    header += b'\x00' * 16                           # CLSID
    header += struct.pack('<H', 0x003E)              # minor version
    header += struct.pack('<H', 0x0003)              # major version (512)
    header += struct.pack('<H', 0xFFFE)              # byte order
    header += struct.pack('<H', 0x0009)              # sector shift (512)
    header += struct.pack('<H', 0x0006)              # mini sector shift (64)
    header += b'\x00' * 6                            # reserved
    header += struct.pack('<I', 0)                   # num dir sectors (v3 = 0)
    header += struct.pack('<I', fat_sectors)         # num FAT sectors
    header += struct.pack('<I', dir_first)           # first dir sector
    header += struct.pack('<I', 0)                   # transaction sig
    header += struct.pack('<I', MINI_CUTOFF)         # mini stream cutoff
    header += struct.pack('<I', minifat_first if minifat_sectors else ENDOFCHAIN)
    header += struct.pack('<I', minifat_sectors)     # num minifat sectors
    header += struct.pack('<I', ENDOFCHAIN)          # first DIFAT sector
    header += struct.pack('<I', 0)                   # num DIFAT sectors
    difat = [FREESECT] * 109
    for k in range(fat_sectors):
        if k < 109:
            difat[k] = fat_first + k
    for v in difat:
        header += struct.pack('<I', v)
    assert len(header) == 512

    # --- 6) Cuerpo ---
    dir_bytes = b''.join(
        _dir_entry(e[0], e[1], e[2], e[3], e[4], e[5], e[6], e[7]) for e in entries)
    if len(dir_bytes) % SECTOR:
        dir_bytes += b'\x00' * (SECTOR - len(dir_bytes) % SECTOR)

    body = dir_bytes + minifat_bytes + bytes(ministream) + fat_bytes

    with open(path, 'wb') as f:
        f.write(header)
        f.write(body)
    return path


def main(out='vbaProject.bin'):
    streams, src = build_streams()
    write_cfb(streams, out)
    _verify(out, src)
    print("OK ->", out)


def _verify(path, src):
    import olefile, io
    with open(path, 'rb') as f:
        data = f.read()
    o = olefile.OleFileIO(io.BytesIO(data))
    listing = ['/'.join(p) for p in o.listdir()]
    need = ['PROJECT', 'PROJECTwm', 'VBA/_VBA_PROJECT', 'VBA/dir', 'VBA/' + MODULE_NAME]
    for n in need:
        assert n in listing, "falta stream %s en %s" % (n, listing)
    mod = o.openstream('VBA/' + MODULE_NAME).read()
    assert decompress(mod) == src, "la fuente leida no coincide"
    draw = decompress(o.openstream('VBA/dir').read())
    assert draw == build_dir_stream(), "dir leido no coincide"
    o.close()
    print("  verificado con olefile:", listing)


if __name__ == '__main__':
    main()
