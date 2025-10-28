Sub Consecutivos()
         Dim WS_Count As Integer
         Dim I As Double
         Dim CONSECUTIVO As Double
         Dim CELDA As String
         Dim pref As String
              
         CELDA = InputBox("Ingrese la celda donde quiere que se copie el consecutivo", "Celda")
         pref = InputBox("Ingrese prefijo del consecutivo", "Prefijo")
         CONSECUTIVO = InputBox("Ingrese el inicio de numero del consecutivo", "Consecutivo")
         
         ' Set WS_Count equal to the number of worksheets in the active
         ' workbook.
         WS_Count = ActiveWorkbook.Worksheets.Count
 
         ' Begin the loop.
         For I = 1 To WS_Count
            
            
            Worksheets(I).Range(CELDA) = pref & CONSECUTIVO
            
            ' Insert your code here.
            ' The following line shows how to reference a sheet within
            ' the loop by displaying the worksheet name in a dialog box.
            CONSECUTIVO = CONSECUTIVO + 1
            'MsgBox ActiveWorkbook.Worksheets(I).Name
 
         Next I
'    ActiveWorkbook.Save
'    ActiveWorkbook.SaveAs Filename:= _
'        "C:\Users\jennifer.campos\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.XLSB" _
'        , FileFormat:=xlExcel12, CreateBackup:=False
'    Range("C2642").Select
'    ActiveWindow.SmallScroll Down:=-18
'    Range("E2626").Select
End Sub

Sub Firmas()

    MiHoja = ActiveSheet.Name
    WS = ActiveWorkbook.Worksheets.Count
    
    For I = 1 To WS
        nombre = ActiveWorkbook.Worksheets(I).Name
        Sheets(nombre).Select
        Range("M71").Select
        ActiveSheet.Paste
    Next I
    
    
'    For Each Hoja In Application.Sheets
    
'    If Hoja.Name <> MiHoja Then
'    Worksheets(Hoja.Name).Select
'    Hoja.Paste Destination:=[M71]
    'Application.Goto Reference:=Worksheets(Hoja.Name).Range("M71"), Scroll:=True
    'Application.Goto Hoja.Cells(1, 1)
    'Application.Goto Sheets(MiHoja).Cells(1, 1)
'    End If
'    Next Hoja
'    Application.Goto Sheets(MiHoja).Cells(1, 1)
    
End Sub

Sub pdf()
    
    Dim numeroHojas As Integer
    Dim cont As Integer
    Dim nombreHoja As String
    Dim dictamen As String
    Dim ruta, pdfi, pdfo As String
    Dim wsh As Object
    
    'rendimiento
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ruta = "C:\DictamenesRetie"
    pdfi = ""
    pdfo = ""
    numeroDictamen = ""
    
    'Set wsh = VBA.CreateObject("WScript.Shell")
    'Set wsh = CreateObject("Wscript.Shell")
    numeroHojas = Sheets.Count
    
    'configurar impresora windows pdf
    Application.ActivePrinter = GetPrinterFullName("Microsoft Print to PDF")
    
    For cont = 1 To numeroHojas
        If Sheets(cont).Visible = True Then
            Sheets(cont).Activate
            'identificar tipo de dictamen
            numeroDictamen = ""
            If Range("A11").Value = "B. IDENTIFICACIÓN DE LA INSTALACIÓN ELÉCTRICA DE USO FINAL OBJETO DEL DICTAMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$P$78"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A11").Value = "B. IDENTIFICACIÓN DE LA LÍNEA OBJETO DEL DICTÁMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$O$76"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A11").Value = "B. IDENTIFICACIÓN DE LA INSTALACIÓN DE DISTRIBUCIÓN OBJETO DEL DICTAMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$R$77"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A11").Value = "B. IDENTIFICACIÓN SUBESTACIÓN OBJETO DEL DICTAMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$O$88"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A2").Value = "DICTAMEN DE INSPECCIÓN Y VERIFICACIÓN DE ILUMINACIÓN EXTERIOR" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$N$70"
                numeroDictamen = Range("N4").Value
            ElseIf Range("A5").Value = "DICTAMEN DE INSPECCIÓN Y VERIFICACIÓN DE ILUMINACIÓN INTERIOR SEGÚN RETILAP" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$N$65"
                numeroDictamen = Range("N4").Value
            End If
            'MsgBox numeroDictamen
                        
            If numeroDictamen <> "" Then
                'configurar imagen encabezado
                ActiveSheet.PageSetup.CenterHeaderPicture.Filename = ruta & "\CERTIFICADO.png"
                'ajustar tamaño imagen encabezado
                With ActiveSheet.PageSetup.CenterHeaderPicture
                    .Height = 825.75
                    .Width = 632.25
                End With
                'configurar pagina
                Call configpag
                'generar pdf original
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=ruta & "\Original\" & numeroDictamen & ".pdf", Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
                'Filename:=ruta & "\Original\" & numeroDictamen & "-1.pdf", Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
        
                'configurar imagen encabezado
                ActiveSheet.PageSetup.CenterHeaderPicture.Filename = ruta & "\DUPLICADO.png"
                'ajustar tamaño imagen encabezado
                With ActiveSheet.PageSetup.CenterHeaderPicture
                    .Height = 825.75
                    .Width = 632.25
                End With
                'configurar pagina
                Call configpag
                'generar pdf copia
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=ruta & "\Copia\" & numeroDictamen & ".pdf", Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
                'Filename:=ruta & "\Copia\" & numeroDictamen & "-1.pdf", Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
       
                'proteger pdf original
                'pdfi = ruta & "\Original\" & numeroDictamen & "-1.pdf"
                'pdfo = ruta & "\Original\" & numeroDictamen & ".pdf"
                'MsgBox "cmd.exe /c pdftk.exe" & " " & """" & pdfi & """" & " output " & """" & pdfo & """" & " owner_pw foopass allow printing"
                'wsh.Run "cmd.exe /c pdftk.exe" & " " & """" & pdfi & """" & " output " & """" & pdfo & """" & " owner_pw foopass allow printing > C:\DictamenesRetie\Error.txt", 1, True
                'Shell "cmd.exe /c pdftk.exe" & " " & """" & pdfi & """" & " output " & """" & pdfo & """" & " owner_pw foopass allow printing > C:\DictamenesRetie\Error.txt", 1
                'Kill pdfi
                'proteger pdf copia
                'pdfi = ruta & "\Copia\" & numeroDictamen & "-1.pdf"
                'pdfo = ruta & "\Copia\" & numeroDictamen & ".pdf"
                'wsh.Run "cmd.exe /c pdftk.exe" & " " & """" & pdfi & """" & " output " & """" & pdfo & """" & " owner_pw foopass allow printing", 1, True
                'Kill pdfi
            End If
        End If
    Next cont
    
    'Set wsh = Nothing
    
    'rendimiento
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox ("Archivos PDF generados correctamente.")
    
End Sub

Sub pdf_2024()
    
    Dim numeroHojas As Integer
    Dim cont As Integer
    Dim nombreHoja As String
    Dim dictamen As String
    Dim ruta, pdfi, pdfo As String
    Dim wsh As Object
    
    'rendimiento
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ruta = "C:\DictamenesRetie"
    pdfi = ""
    pdfo = ""
    numeroDictamen = ""
    
    'Set wsh = VBA.CreateObject("WScript.Shell")
    'Set wsh = CreateObject("Wscript.Shell")
    numeroHojas = Sheets.Count
    
    'configurar impresora windows pdf
    Application.ActivePrinter = GetPrinterFullName("Microsoft Print to PDF")
    
    For cont = 1 To numeroHojas
        If Sheets(cont).Visible = True Then
            Sheets(cont).Activate
            'identificar tipo de dictamen
            numeroDictamen = ""
            If Range("A11").Value = "B. IDENTIFICACIÓN DE LA INSTALACIÓN ELÉCTRICA DE USO FINAL OBJETO DEL DICTAMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$P$78"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A11").Value = "B. IDENTIFICACIÓN DE LA LÍNEA OBJETO DEL DICTÁMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$O$76"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A11").Value = "B. IDENTIFICACIÓN DE LA INSTALACIÓN DE DISTRIBUCIÓN OBJETO DEL DICTAMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$R$77"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A11").Value = "B. IDENTIFICACIÓN SUBESTACIÓN OBJETO DEL DICTAMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$O$88"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A2").Value = "DICTAMEN DE INSPECCIÓN Y VERIFICACIÓN DE ILUMINACIÓN EXTERIOR" Or Range("G2").Value = "DICTAMEN DE INSPECCIÓN Y VERIFICACIÓN DE ILUMINACIÓN EXTERIOR" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$N$74"
                numeroDictamen = Range("N5").Value
            ElseIf Range("A2").Value = "DICTAMEN DE INSPECCIÓN Y VERIFICACIÓN DE ILUMINACIÓN INTERIOR SEGÚN RETILAP" Or Range("F2").Value = "DICTAMEN DE INSPECCIÓN Y VERIFICACIÓN DE ILUMINACIÓN INTERIOR SEGÚN RETILAP" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$N$65"
                numeroDictamen = Range("N4").Value
            End If
            'MsgBox numeroDictamen
                        
            If numeroDictamen <> "" Then
                'configurar imagen encabezado
                ActiveSheet.PageSetup.CenterHeaderPicture.Filename = ruta & "\CERTIFICADO_2024.png"
                'ajustar tamaño imagen encabezado
                With ActiveSheet.PageSetup.CenterHeaderPicture
                    .Height = 819.75
                    .Width = 633.75
                End With
                'configurar pagina
                Call configpag_2024
                'generar pdf original
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=ruta & "\Original\" & numeroDictamen & ".pdf", Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
                        
                'configurar imagen encabezado
                ActiveSheet.PageSetup.CenterHeaderPicture.Filename = ruta & "\DUPLICADO_2024.png"
                'ajustar tamaño imagen encabezado
                With ActiveSheet.PageSetup.CenterHeaderPicture
                    .Height = 819.75
                    .Width = 633.75
                End With
                'configurar pagina
                Call configpag_2024
                'generar pdf copia
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=ruta & "\Copia\" & numeroDictamen & ".pdf", Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
                
            End If
        End If
    Next cont
     
    'rendimiento
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox ("Archivos PDF generados correctamente.")
    
End Sub

Sub pdf_R()
    
    Dim numeroHojas As Integer
    Dim cont As Integer
    Dim nombreHoja As String
    Dim dictamen As String
    Dim ruta, pdfi, pdfo As String
    Dim wsh As Object
    
    'rendimiento
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ruta = "C:\DictamenesRetie"
    pdfi = ""
    pdfo = ""
    numeroDictamen = ""
    
    'Set wsh = VBA.CreateObject("WScript.Shell")
    'Set wsh = CreateObject("Wscript.Shell")
    numeroHojas = Sheets.Count
    
    'configurar impresora windows pdf
    Application.ActivePrinter = GetPrinterFullName("Microsoft Print to PDF")
    
    For cont = 1 To numeroHojas
        If Sheets(cont).Visible = True Then
            Sheets(cont).Activate
            'identificar tipo de dictamen
            numeroDictamen = ""
            If Range("A11").Value = "B. IDENTIFICACIÓN DE LA INSTALACIÓN ELÉCTRICA DE USO FINAL OBJETO DEL DICTAMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$P$78"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A11").Value = "B. IDENTIFICACIÓN DE LA LÍNEA OBJETO DEL DICTÁMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$O$76"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A11").Value = "B. IDENTIFICACIÓN DE LA INSTALACIÓN DE DISTRIBUCIÓN OBJETO DEL DICTAMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$R$77"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A11").Value = "B. IDENTIFICACIÓN SUBESTACIÓN OBJETO DEL DICTAMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$O$88"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A2").Value = "DICTAMEN DE INSPECCIÓN Y VERIFICACIÓN DE ILUMINACIÓN EXTERIOR" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$N$70"
                numeroDictamen = Range("N5").Value
            ElseIf Range("A5").Value = "DICTAMEN DE INSPECCIÓN Y VERIFICACIÓN DE ILUMINACIÓN INTERIOR SEGÚN RETILAP" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$N$66"
                numeroDictamen = Range("N7").Value
            End If
            'MsgBox numeroDictamen
                        
            If numeroDictamen <> "" Then
                'configurar imagen encabezado
                ActiveSheet.PageSetup.CenterHeaderPicture.Filename = ruta & "\CERTIFICADO.png"
                'ajustar tamaño imagen encabezado
                With ActiveSheet.PageSetup.CenterHeaderPicture
                    '.Height = 759.6 '26.8 cm
                    '.Width = 581.4 '20.51 cm
                    .Height = 793.8 '28cm
                    .Width = 608.4
                End With
                'configurar pagina
                Call configpag
                'generar pdf original
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=ruta & "\Original\" & numeroDictamen & ".pdf", Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
                'Filename:=ruta & "\Original\" & numeroDictamen & "-1.pdf", Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
        
                'configurar imagen encabezado
                ActiveSheet.PageSetup.CenterHeaderPicture.Filename = ruta & "\DUPLICADO.png"
                'ajustar tamaño imagen encabezado
                With ActiveSheet.PageSetup.CenterHeaderPicture
                    '.Height = 759.6 '26.8 cm
                    '.Width = 581.4 '20.51 cm
                    .Height = 793.8 '28cm
                    .Width = 608.4
                End With
                'configurar pagina
                Call configpag
                'generar pdf copia
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=ruta & "\Copia\" & numeroDictamen & ".pdf", Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
                'Filename:=ruta & "\Copia\" & numeroDictamen & "-1.pdf", Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
       
                'proteger pdf original
                'pdfi = ruta & "\Original\" & numeroDictamen & "-1.pdf"
                'pdfo = ruta & "\Original\" & numeroDictamen & ".pdf"
                'MsgBox "cmd.exe /c pdftk.exe" & " " & """" & pdfi & """" & " output " & """" & pdfo & """" & " owner_pw foopass allow printing"
                'wsh.Run "cmd.exe /c pdftk.exe" & " " & """" & pdfi & """" & " output " & """" & pdfo & """" & " owner_pw foopass allow printing > C:\DictamenesRetie\Error.txt", 1, True
                'Shell "cmd.exe /c pdftk.exe" & " " & """" & pdfi & """" & " output " & """" & pdfo & """" & " owner_pw foopass allow printing > C:\DictamenesRetie\Error.txt", 1
                'Kill pdfi
                'proteger pdf copia
                'pdfi = ruta & "\Copia\" & numeroDictamen & "-1.pdf"
                'pdfo = ruta & "\Copia\" & numeroDictamen & ".pdf"
                'wsh.Run "cmd.exe /c pdftk.exe" & " " & """" & pdfi & """" & " output " & """" & pdfo & """" & " owner_pw foopass allow printing", 1, True
                'Kill pdfi
            End If
        End If
    Next cont
    
    'Set wsh = Nothing
    
    'rendimiento
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox ("Archivos PDF generados correctamente.")
    
End Sub


Sub pdf_esp()
    
    Dim numeroHojas As Integer
    Dim cont As Integer
    Dim nombreHoja As String
    Dim dictamen As String
    Dim ruta, pdfi, pdfo As String
    Dim wsh As Object
    
    'rendimiento
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ruta = "C:\DictamenesRetie"
    pdfi = ""
    pdfo = ""
    numeroDictamen = ""
    
    'Set wsh = VBA.CreateObject("WScript.Shell")
    Set wsh = CreateObject("Wscript.Shell")
    numeroHojas = Sheets.Count
    
    'configurar impresora windows pdf
    Application.ActivePrinter = GetPrinterFullName("Microsoft Print to PDF")
    
    For cont = 1 To numeroHojas
        If Sheets(cont).Visible = True Then
            Sheets(cont).Activate
            'identificar tipo de dictamen
            
            numeroDictamen = ""
            If Range("A11").Value = "B. IDENTIFICACIÓN DE LA INSTALACIÓN ELÉCTRICA DE USO FINAL OBJETO DEL DICTAMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$O$79"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A11").Value = "B. IDENTIFICACIÓN DE LA LÍNEA OBJETO DEL DICTÁMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$O$76"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A11").Value = "B. IDENTIFICACIÓN DE LA INSTALACIÓN DE DISTRIBUCIÓN OBJETO DEL DICTAMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$O$79"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A11").Value = "B. IDENTIFICACIÓN DE LA SUBESTACIÓN OBJETO DEL DICTAMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$O$89"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A2").Value = "DICTAMEN DE INSPECCIÓN Y VERIFICACIÓN DE ILUMINACIÓN EXTERIOR" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$N$70"
                numeroDictamen = Range("N5").Value
            ElseIf Range("A5").Value = "DICTAMEN DE INSPECCIÓN Y VERIFICACIÓN DE ILUMINACIÓN INTERIOR SEGÚN RETILAP" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$N$65"
                numeroDictamen = Range("N7").Value
            End If
            
            
            'MsgBox numeroDictamen
                        
            If numeroDictamen <> "" Then
                'configurar imagen encabezado
                ActiveSheet.PageSetup.CenterHeaderPicture.Filename = ruta & "\CERTIFICADO.png"
                'ajustar tamaño imagen encabezado
                With ActiveSheet.PageSetup.CenterHeaderPicture
                    .Height = 790.8
                    .Width = 606
                End With
                'configurar pagina
                Call configpag
                'generar pdf original
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=ruta & "\Original\" & numeroDictamen & ".pdf", Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
                'Filename:=ruta & "\Original\" & numeroDictamen & "-1.pdf", Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
        
                'configurar imagen encabezado
                ActiveSheet.PageSetup.CenterHeaderPicture.Filename = ruta & "\DUPLICADO.png"
                'ajustar tamaño imagen encabezado
                With ActiveSheet.PageSetup.CenterHeaderPicture
                    .Height = 790.8
                    .Width = 606
                End With
                'configurar pagina
                Call configpag
                'generar pdf copia
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=ruta & "\Copia\" & numeroDictamen & ".pdf", Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
                'Filename:=ruta & "\Copia\" & numeroDictamen & "-1.pdf", Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
       
                'proteger pdf original
                'pdfi = ruta & "\Original\" & numeroDictamen & "-1.pdf"
                'pdfo = ruta & "\Original\" & numeroDictamen & ".pdf"
                'MsgBox "cmd.exe /c pdftk.exe" & " " & """" & pdfi & """" & " output " & """" & pdfo & """" & " owner_pw foopass allow printing"
                'wsh.Run "cmd.exe /c pdftk.exe" & " " & """" & pdfi & """" & " output " & """" & pdfo & """" & " owner_pw foopass allow printing > C:\DictamenesRetie\Error.txt", 1, True
                'Shell "cmd.exe /c pdftk.exe" & " " & """" & pdfi & """" & " output " & """" & pdfo & """" & " owner_pw foopass allow printing > C:\DictamenesRetie\Error.txt", 1
                'Kill pdfi
                'proteger pdf copia
                'pdfi = ruta & "\Copia\" & numeroDictamen & "-1.pdf"
                'pdfo = ruta & "\Copia\" & numeroDictamen & ".pdf"
                'wsh.Run "cmd.exe /c pdftk.exe" & " " & """" & pdfi & """" & " output " & """" & pdfo & """" & " owner_pw foopass allow printing", 1, True
                'Kill pdfi
            End If
        End If
    Next cont
    
    Set wsh = Nothing
    
    'rendimiento
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox ("Archivos PDF generados correctamente.")
    
End Sub

Sub configpag()

    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = "&G"
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(1.37795275590551)
        .BottomMargin = Application.InchesToPoints(0.984251968503937)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintSheetEnd
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = False
        .AlignMarginsHeaderFooter = False
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    
End Sub

Sub configpag_2024()

    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = "&G"
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(1.18110236220472)
        .BottomMargin = Application.InchesToPoints(1.37795275590551)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintSheetEnd
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = False
        .AlignMarginsHeaderFooter = False
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    
End Sub


Public Function GetPrinterFullName(Printer As String) As String
 
    ' This function returns the full name of the first printerdevice that matches Printer.
    ' Full name is like "PDFCreator on Ne01:" for a English Windows and like
    ' "PDFCreator sur Ne01:" for French.
    ' Created: Frans Bus, 2015. See http://pixcels.nl/set-activeprinter-excel
    ' see http://blogs.msdn.com/b/alejacma/archive/2008/04/11/how-to-read-a-registry-key-and-its-values.aspx
    ' see http://www.experts-exchange.com/Software/Microsoft_Applications/Q_27566782.html
 
    Const HKEY_CURRENT_USER = &H80000001
    Dim regobj As Object
    Dim aTypes As Variant
    Dim aDevices As Variant
    Dim vDevice As Variant
    Dim sValue As String
    Dim v As Variant
    Dim sLocaleOn As String
     
    ' get locale "on" from current activeprinter
    v = Split(Application.ActivePrinter, Space(1))
    sLocaleOn = Space(1) & CStr(v(UBound(v) - 1)) & Space(1)
     
    ' connect to WMI registry provider on current machine with current user
    Set regobj = GetObject("WINMGMTS:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
     
    ' get the Devices from the registry
    regobj.EnumValues HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Devices", aDevices, aTypes
     
    ' find Printer and create full name
    For Each vDevice In aDevices
        ' get port of device
        regobj.GetStringValue HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Devices", vDevice, sValue
        ' select device
        If Left(vDevice, Len(Printer)) = Printer Then ' match!
            ' create localized printername
            GetPrinterFullName = vDevice & sLocaleOn & Split(sValue, ",")(1)
            Exit Function
        End If
    Next
     
    ' at this point no match found
    GetPrinterFullName = vbNullString
 
End Function
Sub pdf_2024_2()

Dim numeroHojas As Integer
    Dim cont As Integer
    Dim nombreHoja As String
    Dim dictamen As String
    Dim ruta, pdfi, pdfo As String
    Dim wsh As Object
    
    'rendimiento
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ruta = "C:\DictamenesRetie"
    pdfi = ""
    pdfo = ""
    numeroDictamen = ""
    
    'Set wsh = VBA.CreateObject("WScript.Shell")
    'Set wsh = CreateObject("Wscript.Shell")
    numeroHojas = Sheets.Count
    
    'configurar impresora windows pdf
    Application.ActivePrinter = GetPrinterFullName("Microsoft Print to PDF")
    
    For cont = 1 To numeroHojas
        If Sheets(cont).Visible = True Then
            Sheets(cont).Activate
            'identificar tipo de dictamen
            numeroDictamen = ""
            If Range("A11").Value = "B. IDENTIFICACIÓN DE LA INSTALACIÓN ELÉCTRICA DE USO FINAL OBJETO DEL DICTAMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$P$78"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A11").Value = "B. IDENTIFICACIÓN DE LA LÍNEA OBJETO DEL DICTÁMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$O$76"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A11").Value = "B. IDENTIFICACIÓN DE LA INSTALACIÓN DE DISTRIBUCIÓN OBJETO DEL DICTAMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$R$77"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A11").Value = "B. IDENTIFICACIÓN SUBESTACIÓN OBJETO DEL DICTAMEN" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$O$88"
                If Range("M6").Value = "" Then
                    numeroDictamen = Range("N6").Value
                Else
                    numeroDictamen = Range("M6").Value
                End If
            ElseIf Range("A2").Value = "DICTAMEN DE INSPECCIÓN Y VERIFICACIÓN DE ILUMINACIÓN EXTERIOR" Or Range("G2").Value = "DICTAMEN DE INSPECCIÓN Y VERIFICACIÓN DE ILUMINACIÓN EXTERIOR" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$N$74"
                numeroDictamen = Range("N5").Value
            ElseIf Range("A2").Value = "DICTAMEN DE INSPECCIÓN Y VERIFICACIÓN DE ILUMINACIÓN INTERIOR SEGÚN RETILAP" Or Range("F2").Value = "DICTAMEN DE INSPECCIÓN Y VERIFICACIÓN DE ILUMINACIÓN INTERIOR SEGÚN RETILAP" Then
                ActiveSheet.PageSetup.PrintArea = "$A$1:$N$65"
                numeroDictamen = Range("N4").Value
            End If
            'MsgBox numeroDictamen
                        
            If numeroDictamen <> "" Then
                'configurar imagen encabezado
                ActiveSheet.PageSetup.CenterHeaderPicture.Filename = ruta & "\CERTIFICADO_2024_2.png"
                'ajustar tamaño imagen encabezado
                With ActiveSheet.PageSetup.CenterHeaderPicture
                    .Height = 819.75
                    .Width = 633.75
                End With
                'configurar pagina
                Call configpag_2024
                'generar pdf original
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=ruta & "\Original\" & numeroDictamen & ".pdf", Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
                        
                'configurar imagen encabezado
                ActiveSheet.PageSetup.CenterHeaderPicture.Filename = ruta & "\DUPLICADO_2024_2.png"
                'ajustar tamaño imagen encabezado
                With ActiveSheet.PageSetup.CenterHeaderPicture
                    .Height = 819.75
                    .Width = 633.75
                End With
                'configurar pagina
                Call configpag_2024
                'generar pdf copia
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=ruta & "\Copia\" & numeroDictamen & ".pdf", Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
                
            End If
        End If
    Next cont
     
    'rendimiento
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox ("Archivos PDF generados correctamente.")
End Sub
