Option Explicit

'==============================================================
' ?? GENERAR PDFS DE DICTÁMENES - NORMAL
'==============================================================
Sub pdf()
    Dim ws As Worksheet
    Dim numeroDictamen As String
    Dim ruta As String
    Dim archivoOriginal As String, archivoCopia As String

    ruta = "C:\Dictamenes2025"
    If Dir(ruta, vbDirectory) = "" Then MkDir ruta
    If Dir(ruta & "\Original", vbDirectory) = "" Then MkDir ruta & "\Original"
    If Dir(ruta & "\Copia", vbDirectory) = "" Then MkDir ruta & "\Copia"

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            numeroDictamen = ""
            On Error Resume Next
            If ws.Range("Q4").value <> "" Then numeroDictamen = ws.Range("Q4").value
            If ws.Range("R4").value <> "" Then numeroDictamen = ws.Range("R4").value
            If ws.Range("O7").value <> "" Then numeroDictamen = ws.Range("O7").value
            If ws.Range("L6").value <> "" Then numeroDictamen = ws.Range("L6").value
            On Error GoTo 0

            If Trim(numeroDictamen) = "" Then GoTo Siguiente

            numeroDictamen = CleanFileName(numeroDictamen)
            archivoOriginal = ruta & "\Original\" & numeroDictamen & ".pdf"
            archivoCopia = ruta & "\Copia\" & numeroDictamen & "_DUP.pdf"

            ConfigurarCabecera ws, ruta
            configpag

            ' Exportar ORIGINAL
            ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=archivoOriginal, _
                Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=True, OpenAfterPublish:=False

            ' Exportar DUPLICADO
            ConfigurarCabecera ws, ruta, True
            ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=archivoCopia, _
                Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=True, OpenAfterPublish:=False
        End If
Siguiente:
    Next ws

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "? PDFs generados correctamente en C:\Dictamenes2025", vbInformation
End Sub
'==============================================================
' ?? GENERAR PDF DE ILUMINACIÓN (varias hojas ? un solo PDF)
'==============================================================
Sub pdf_iluminacion()
    Dim ws As Worksheet
    Dim hojasSeleccionadas As Sheets
    Dim ruta As String, numeroDictamen As String
    Dim archivoOriginal As String, archivoCopia As String
    Dim i As Long
    Dim tempPDF As String, archivoTemporal() As String
    Dim carpetaTemp As String
    Dim unionPDF As String

    ' --- Validar selección ---
    If TypeName(ActiveWindow) <> "Window" Or ActiveWindow.SelectedSheets.count = 0 Then
        MsgBox "Selecciona (Ctrl+clic) las hojas del dictamen y vuelve a ejecutar.", vbExclamation
        Exit Sub
    End If
    Set hojasSeleccionadas = ActiveWindow.SelectedSheets

    ruta = "C:\Dictamenes2025"
    If Dir(ruta, vbDirectory) = "" Then MkDir ruta
    If Dir(ruta & "\Original", vbDirectory) = "" Then MkDir ruta & "\Original"
    If Dir(ruta & "\Copia", vbDirectory) = "" Then MkDir ruta & "\Copia"
    carpetaTemp = ruta & "\Temp"
    If Dir(carpetaTemp, vbDirectory) = "" Then MkDir carpetaTemp

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ' --- Obtener número de dictamen de la primera hoja ---
    Set ws = hojasSeleccionadas(1)
    On Error Resume Next
    numeroDictamen = ws.Range("Q4").value
    If numeroDictamen = "" Then numeroDictamen = ws.Range("R4").value
    If numeroDictamen = "" Then numeroDictamen = ws.Range("O7").value
    If numeroDictamen = "" Then numeroDictamen = ws.Range("L6").value
    On Error GoTo 0

    If Trim(numeroDictamen) = "" Then
        MsgBox "?? No se encontró número de dictamen en la primera hoja.", vbExclamation
        GoTo Salir
    End If
    numeroDictamen = CleanFileName(numeroDictamen)

    archivoOriginal = ruta & "\Original\" & numeroDictamen & ".pdf"
    archivoCopia = ruta & "\Copia\" & numeroDictamen & "_DUP.pdf"

    ' --- Exportar todas las hojas como un solo PDF (ORIGINAL) ---
    For Each ws In hojasSeleccionadas
        ConfigurarCabecera_IluminacionFondoEspecial ws, ruta, False
        ConfigurarMargenesIluminacion ws
    Next ws

    hojasSeleccionadas.Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=archivoOriginal, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

    ' --- Exportar todas las hojas como un solo PDF (DUPLICADO) ---
    For Each ws In hojasSeleccionadas
        ConfigurarCabecera_IluminacionFondoEspecial ws, ruta, True
        ConfigurarMargenesIluminacion ws
    Next ws

    hojasSeleccionadas.Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=archivoCopia, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

    MsgBox "? PDF generado correctamente:" & vbCrLf & _
           archivoOriginal & vbCrLf & archivoCopia, vbInformation

Salir:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

'==============================================================
' ?? CONFIGURAR CABECERA GENERAL (para dictámenes normales)
'==============================================================
Private Sub ConfigurarCabecera(ws As Worksheet, ruta As String, Optional esDuplicado As Boolean = False)
    Dim img As String

    If esDuplicado Then
        img = ruta & "\DUPLICADO.png"
    Else
        img = ruta & "\CERTIFICADO.png"
    End If

    If Dir(img) = "" Then Exit Sub

    With ws.PageSetup
        .CenterHeader = "&G"
        .CenterHeaderPicture.Filename = img
        On Error Resume Next
        .CenterHeaderPicture.AlignWithMargins = False
        On Error GoTo 0

        .CenterHeaderPicture.Height = 1008
        .CenterHeaderPicture.Width = 632.25
        .TopMargin = Application.CentimetersToPoints(2.8)
        .BottomMargin = Application.CentimetersToPoints(3)
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = 0
        .FooterMargin = 0
        .CenterHorizontally = True
        .CenterVertically = True
        .Orientation = xlPortrait
        .PaperSize = xlPaperLegal
    End With
End Sub


'==============================================================
' ?? CONFIGURAR CABECERA - ILUMINACIÓN (fondo especial)
'==============================================================
Private Sub ConfigurarCabecera_IluminacionFondoEspecial(ws As Worksheet, ruta As String, Optional esDuplicado As Boolean = False)
    Dim img As String

    If esDuplicado Then
        img = ruta & "\DUPLICADO_ILUMINACION.png"
    Else
        img = ruta & "\CERTIFICADO_ILUMINACION.png"
    End If

    If Dir(img) = "" Then Exit Sub

    With ws.PageSetup
        .CenterHeader = "&G"
        .CenterHeaderPicture.Filename = img
        On Error Resume Next
        .CenterHeaderPicture.AlignWithMargins = False
        On Error GoTo 0

        .CenterHeaderPicture.Height = 1008
        .CenterHeaderPicture.Width = 632.25
        .TopMargin = Application.CentimetersToPoints(2.8)
        .BottomMargin = Application.CentimetersToPoints(3)
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = 0
        .FooterMargin = 0
        .CenterHorizontally = True
        .CenterVertically = True
        .Orientation = xlPortrait
        .PaperSize = xlPaperLegal
    End With
End Sub


'==============================================================
' ?? CONFIGURAR MÁRGENES ESPECIALES - ILUMINACIÓN
'==============================================================
Private Sub ConfigurarMargenesIluminacion(ws As Worksheet)
    With ws.PageSetup
        .PaperSize = xlPaperLegal
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .TopMargin = Application.CentimetersToPoints(-0.3)
        .BottomMargin = Application.CentimetersToPoints(0.8)
        .LeftMargin = Application.CentimetersToPoints(0.2)
        .RightMargin = Application.CentimetersToPoints(0.2)
        .HeaderMargin = 0
        .FooterMargin = 0
        .CenterHorizontally = True
        .CenterVertically = True
    End With
End Sub


'==============================================================
' ?? CONFIGURAR PÁGINA GENERAL
'==============================================================
Private Sub configpag()
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .AlignMarginsHeaderFooter = False
        .PrintQuality = 600
    End With
    Application.PrintCommunication = True
End Sub


'==============================================================
' ?? LIMPIAR NOMBRE DE ARCHIVO
'==============================================================
Private Function CleanFileName(ByVal s As String) As String
    Dim i As Long, bad As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, CStr(bad(i)), "-")
    Next i
    s = Trim(s)
    If Right(s, 1) = "." Then s = Left(s, Len(s) - 1)
    If s = "" Then s = "DICTAMEN"
    CleanFileName = s
End Function


'==============================================================
' ?? CONSECUTIVOS AUTOMÁTICOS
'==============================================================
Public Sub Consecutivos()
    Dim ws As Worksheet
    Dim celdaDestino As String
    Dim consecutivo As Long
    Dim prefijo As String

    celdaDestino = "Q4"
    prefijo = "SM-"
    consecutivo = InputBox("Ingrese número inicial de consecutivo:", "Consecutivos", 1)
    If Not IsNumeric(consecutivo) Then Exit Sub

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Range(celdaDestino).value = prefijo & Format(consecutivo, "000")
            consecutivo = consecutivo + 1
        End If
    Next ws

    MsgBox "? Consecutivos generados correctamente.", vbInformation
End Sub


'==============================================================
' ??? FIRMAS - COPIA AUTOMÁTICA ENTRE HOJAS
'==============================================================
Public Sub Firmas()
    Dim ws As Worksheet
    Dim hojaOrigen As Worksheet
    Dim shp As Shape, imgTemp As Shape
    Dim listaPos As Collection
    Dim i As Long
    Dim seleccionEncontrada As Boolean

    Set hojaOrigen = ActiveSheet
    Set listaPos = New Collection
    seleccionEncontrada = False

    For Each shp In hojaOrigen.Shapes
        If shp.Type = msoPicture Then
            On Error Resume Next
            If shp.Selected Then
                On Error GoTo 0
                listaPos.Add Array(shp.Left, shp.Top, shp.Width, shp.Height)
                seleccionEncontrada = True
            End If
            On Error GoTo 0
        End If
    Next shp

    If Not seleccionEncontrada Then
        MsgBox "?? No hay imágenes seleccionadas." & vbCrLf & _
               "Selecciona una o más firmas (Ctrl + clic) y vuelve a ejecutar la macro.", vbExclamation
        Exit Sub
    End If

    Selection.Copy
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible And ws.name <> hojaOrigen.name Then
            ws.Paste
            For i = 1 To listaPos.count
                If ws.Shapes.count >= i Then
                    Set imgTemp = ws.Shapes(ws.Shapes.count - listaPos.count + i)
                    imgTemp.Left = listaPos(i)(0)
                    imgTemp.Top = listaPos(i)(1)
                    imgTemp.Width = listaPos(i)(2)
                    imgTemp.Height = listaPos(i)(3)
                End If
            Next i
        End If
    Next ws

    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "? Firmas copiadas correctamente en todas las hojas visibles.", vbInformation
End Sub

