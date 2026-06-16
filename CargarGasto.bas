Attribute VB_Name = "modNomina"
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
    malo = Array(":", "\", "/", "?", "*", "[", "]")
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
