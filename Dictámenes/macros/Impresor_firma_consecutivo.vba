'==============================================================
' 🔧 GENERAR PDFS DE DICTÁMENES - ORIGINAL Y DUPLICADO
'     (Iluminación Exterior: HOJA 1 + HOJA 2 unificadas con Select)
'==============================================================
Sub pdf()
    Dim ws As Worksheet
    Dim wsHoja1 As Worksheet, wsHoja2 As Worksheet
    Dim numeroDictamen As String
    Dim ruta As String
    Dim archivoOriginal As String, archivoCopia As String
    Dim esIluminacion As Boolean

    ruta = "C:\Dictamenes2025"
    If Dir(ruta, vbDirectory) = "" Then MkDir ruta
    If Dir(ruta & "\Original", vbDirectory) = "" Then MkDir ruta & "\Original"
    If Dir(ruta & "\Copia", vbDirectory) = "" Then MkDir ruta & "\Copia"

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error Resume Next
    Set wsHoja1 = Worksheets("HOJA 1")
    Set wsHoja2 = Worksheets("HOJA 2")
    On Error GoTo 0

    '--------------------------------------------------------------
    ' 🔹 CASO ESPECIAL: ILUMINACIÓN EXTERIOR / ALUMBRADO PÚBLICO
    '--------------------------------------------------------------
    If Not wsHoja1 Is Nothing Then
        esIluminacion = (InStr(1, wsHoja1.Name, "ILUMINACIÓN EXTERIOR", vbTextCompare) > 0 _
                      Or InStr(1, wsHoja1.Name, "ALUMBRADO PÚBLICO", vbTextCompare) > 0)
    End If

    If esIluminacion Then
        ' === Obtener número de dictamen desde HOJA 1 ===
        On Error Resume Next
        If wsHoja1.Range("Q4").Value <> "" Then numeroDictamen = wsHoja1.Range("Q4").Value
        If wsHoja1.Range("R4").Value <> "" Then numeroDictamen = wsHoja1.Range("R4").Value
        If wsHoja1.Range("M7").Value <> "" Then numeroDictamen = wsHoja1.Range("M7").Value
        If wsHoja1.Range("L6").Value <> "" Then numeroDictamen = wsHoja1.Range("L6").Value
        If wsHoja1.Range("M6").Value <> "" Then numeroDictamen = wsHoja1.Range("M6").Value
        On Error GoTo 0

        If Trim(numeroDictamen) = "" Then
            MsgBox "⚠️ No se encontró el número de dictamen en HOJA 1.", vbExclamation
            GoTo Fin
        End If

        numeroDictamen = CleanFileName(numeroDictamen)
        archivoOriginal = ruta & "\Original\" & numeroDictamen & ".pdf"
        archivoCopia = ruta & "\Copia\" & numeroDictamen & "_DUP.pdf"

        ' === Configurar ambas hojas ===
        ConfigurarCabecera wsHoja1, ruta
        configpag
        ConfigurarIluminacionExterior wsHoja1
        If Not wsHoja2 Is Nothing Then ConfigurarIluminacionExterior wsHoja2

        ' =======================================================
        ' === Exportar ORIGINAL (fusionado) ===
        ' =======================================================
        Debug.Print "Creando ORIGINAL → " & archivoOriginal
        If Not wsHoja2 Is Nothing Then
            ThisWorkbook.Sheets(Array(wsHoja1.Name, wsHoja2.Name)).Select
        Else
            wsHoja1.Select
        End If
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=archivoOriginal, Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=True, OpenAfterPublish:=False

        ' =======================================================
        ' === Exportar DUPLICADO (fusionado) ===
        ' =======================================================
        With wsHoja1.PageSetup
            .CenterHeaderPicture.Filename = ruta & "\DUPLICADO.png"
            .CenterHeader = "&G"
        End With
        If Not wsHoja2 Is Nothing Then
            With wsHoja2.PageSetup
                .CenterHeaderPicture.Filename = ruta & "\DUPLICADO.png"
                .CenterHeader = "&G"
            End With
        End If

        Debug.Print "Creando DUPLICADO → " & archivoCopia
        If Not wsHoja2 Is Nothing Then
            ThisWorkbook.Sheets(Array(wsHoja1.Name, wsHoja2.Name)).Select
        Else
            wsHoja1.Select
        End If
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=archivoCopia, Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=True, OpenAfterPublish:=False

        wsHoja1.Select
        Debug.Print "✅ PDF unificado generado (HOJA 1 + HOJA 2)"
    End If

    '--------------------------------------------------------------
    ' 🔹 CASO GENERAL (otros dictámenes)
    '--------------------------------------------------------------
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            If ws.Name <> "HOJA 1" And ws.Name <> "HOJA 2" Then
                numeroDictamen = ""
                On Error Resume Next
                If ws.Range("Q4").Value <> "" Then numeroDictamen = ws.Range("Q4").Value
                If ws.Range("R4").Value <> "" Then numeroDictamen = ws.Range("R4").Value
                If ws.Range("M7").Value <> "" Then numeroDictamen = ws.Range("M7").Value
                If ws.Range("L6").Value <> "" Then numeroDictamen = ws.Range("L6").Value
                If ws.Range("M6").Value <> "" Then numeroDictamen = ws.Range("M6").Value
                On Error GoTo 0

                If Trim(numeroDictamen) <> "" Then
                    numeroDictamen = CleanFileName(numeroDictamen)
                    archivoOriginal = ruta & "\Original\" & numeroDictamen & ".pdf"
                    archivoCopia = ruta & "\Copia\" & numeroDictamen & "_DUP.pdf"

                    ConfigurarCabecera ws, ruta
                    configpag

                    ws.ExportAsFixedFormat Type:=xlTypePDF, _
                        Filename:=archivoOriginal, Quality:=xlQualityStandard, _
                        IncludeDocProperties:=True, IgnorePrintAreas:=True, OpenAfterPublish:=False

                    With ws.PageSetup
                        .CenterHeaderPicture.Filename = ruta & "\DUPLICADO.png"
                        .CenterHeader = "&G"
                    End With

                    ws.ExportAsFixedFormat Type:=xlTypePDF, _
                        Filename:=archivoCopia, Quality:=xlQualityStandard, _
                        IncludeDocProperties:=True, IgnorePrintAreas:=True, OpenAfterPublish:=False

                    Debug.Print "✅ " & ws.Name & " OK"
                End If
            End If
        End If
    Next ws

Fin:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "✅ PDFs generados correctamente en C:\Dictamenes2025", vbInformation
End Sub
