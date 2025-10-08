Option Explicit

Sub Extraer_informacion_de_Dictamenes()

    On Error GoTo ManejarError
    
    ' 🔹 Mejora de rendimiento
    ConfigurarRendimiento False
    
    Dim dlg As FileDialog
    Dim lngCount As Long, i As Integer
    Dim Contador_filas_TOTAL As Long, Contador_filas As Long
    Dim nombre As String, nombre_libro_dictamenes As String
    Dim identificacion_de_dictamen As String, clase_de_dictamen As String
    Dim Nombre_Inspector As String, numero_dictamen As String
    Dim fecha_emision As Variant, fecha_de_inspeccion As Variant
    Dim direccion_proyecto As String, matricula_inspector As String
    Dim nombre_disenador As String, matricula_disenador As String
    Dim nombre_declarante As String, matricula_declarante As String
    Dim desc_alcance As String
    
    Dim wbDestino As Workbook
    Dim wsDestino As Worksheet
    Dim ws As Worksheet
    Dim wbOrigen As Workbook
    
    Set wbDestino = ThisWorkbook
    Set wsDestino = wbDestino.Worksheets(1)
    
    Contador_filas_TOTAL = wsDestino.Range("B2").Value
    
    ' 🔹 Seleccionar archivos
    Set dlg = Application.FileDialog(msoFileDialogOpen)
    With dlg
        .AllowMultiSelect = True
        .Title = "Seleccione los dictámenes a procesar"
        .Filters.Clear
        .Filters.Add "Archivos Excel", "*.xls; *.xlsx; *.xlsm"
        If .Show = -1 Then
            
            For lngCount = 1 To .SelectedItems.Count
                
                nombre = .SelectedItems(lngCount)
                Set wbOrigen = Workbooks.Open(nombre)
                
                For i = 1 To wbOrigen.Worksheets.Count
                    
                    Set ws = wbOrigen.Worksheets(i)
                    identificacion_de_dictamen = Trim(ws.Range("A16").Value)
                    
                    ' Determinar tipo de dictamen
                    Select Case identificacion_de_dictamen
                        Case "C. IDENTIFICACIÓN DE LA INSTALACIÓN DE DISTRIBUCIÓN OBJETO DEL DICTAMEN"
                            clase_de_dictamen = "Distribución"
                        Case "C. IDENTIFICACIÓN DE LA INSTALACIÓN DE TRANSFORMACIÓN OBJETO DEL DICTAMEN"
                            clase_de_dictamen = "Subestación"
                        Case "C. IDENTIFICACIÓN DE LA INSTALACIÓN DE USO FINAL OBJETO DEL DICTAMEN"
                            clase_de_dictamen = "Uso Final"
                        Case Else
                            clase_de_dictamen = "SIN IDENTIFICADOR"
                    End Select
                    
                    ' Leer valores según tipo
                    If clase_de_dictamen <> "SIN IDENTIFICADOR" Then
                        LeerDatosPorTipo ws, clase_de_dictamen, numero_dictamen, Nombre_Inspector, _
                            direccion_proyecto, matricula_inspector, nombre_disenador, matricula_disenador, _
                            nombre_declarante, matricula_declarante, desc_alcance, fecha_emision
                    Else
                        numero_dictamen = "SIN NUMERO"
                        fecha_emision = "SIN FECHA"
                        Nombre_Inspector = "ERROR"
                        direccion_proyecto = ""
                        matricula_inspector = ""
                        nombre_disenador = ""
                        matricula_disenador = ""
                        nombre_declarante = ""
                        matricula_declarante = ""
                        desc_alcance = ""
                    End If
                    
                    ' Control de vacíos
                    If IsEmpty(numero_dictamen) Or numero_dictamen = "" Then numero_dictamen = "SIN NUMERO"
                    If IsEmpty(fecha_emision) Or fecha_emision = "" Then fecha_emision = "SIN FECHA"

                    ' 🔎 Imprimir valores principales
                    Debug.Print "Clase: " & clase_de_dictamen
                    Debug.Print "Número Dictamen: " & numero_dictamen
                    Debug.Print "Fecha Emisión: " & fecha_emision
                    Debug.Print "Inspector: " & Nombre_Inspector & " | Matrícula: " & matricula_inspector
                    Debug.Print "Diseñador: " & nombre_disenador & " | Matrícula: " & matricula_disenador
                    Debug.Print "Declarante: " & nombre_declarante & " | Matrícula: " & matricula_declarante
                    Debug.Print "Dirección: " & direccion_proyecto
                    Debug.Print "Descripción del alcance: " & desc_alcance
                    Debug.Print "--------------------------------------"
                    
                    ' Escribir resultados
                    Contador_filas_TOTAL = Contador_filas_TOTAL + 1
                    With wsDestino
                        .Range("B5").Offset(Contador_filas_TOTAL, 0).Value = numero_dictamen
                        .Range("E5").Offset(Contador_filas_TOTAL, 0).Value = fecha_emision
                        .Range("F5").Offset(Contador_filas_TOTAL, 0).Value = fecha_de_inspeccion
                        .Range("J5").Offset(Contador_filas_TOTAL, 0).Value = clase_de_dictamen
                        .Range("K5").Offset(Contador_filas_TOTAL, 0).Value = direccion_proyecto
                        .Range("O5").Offset(Contador_filas_TOTAL, 0).Value = Nombre_Inspector
                        .Range("AC5").Offset(Contador_filas_TOTAL, 0).Value = matricula_inspector
                        .Range("AL5").Offset(Contador_filas_TOTAL, 0).Value = nombre_disenador
                        .Range("AM5").Offset(Contador_filas_TOTAL, 0).Value = matricula_disenador
                        .Range("AN5").Offset(Contador_filas_TOTAL, 0).Value = nombre_declarante
                        .Range("AO5").Offset(Contador_filas_TOTAL, 0).Value = matricula_declarante
                        .Range("AH5").Offset(Contador_filas_TOTAL, 0).Value = desc_alcance
                    End With
                
                    Next i
                    
                    wbOrigen.Close SaveChanges:=False
    SiguienteArchivo:
                Next lngCount
                
                wsDestino.Range("B2").Value = Contador_filas_TOTAL
                'MsgBox "✅ Extracción completada correctamente.", vbInformation
                
            Else
                MsgBox "No se seleccionaron archivos.", vbExclamation
            End If
        End With
        
      Finalizar:
          ConfigurarRendimiento True
          Application.CutCopyMode = False
          Exit Sub

      ManejarError:
          MsgBox "⚠️ Error en archivo: " & nombre & vbCrLf & _
                "Detalle: " & Err.Description, vbExclamation
          Resume SiguienteArchivo
End Sub

Private Sub ConfigurarRendimiento(ByVal activar As Boolean)
      With Application
          .ScreenUpdating = activar
          .Calculation = IIf(activar, xlCalculationAutomatic, xlCalculationManual)
          .EnableEvents = activar
          .DisplayAlerts = activar
          .AskToUpdateLinks = activar
      End With
      On Error Resume Next
      ActiveSheet.DisplayPageBreaks = activar
      On Error GoTo 0
End Sub


Private Sub LeerDatosPorTipo(ws As Worksheet, tipo As String, _
    ByRef numero As String, ByRef inspector As String, ByRef direccion As String, _
    ByRef matricula As String, ByRef disenador As String, ByRef matriculaDis As String, _
    ByRef declarante As String, ByRef matriculaDec As String, ByRef alcance As String, _
    ByRef fechaEmision As Variant)
    
    Select Case tipo
    
        Case "Distribución"
            inspector = ws.Range("O79").Value
            numero = ws.Range("Q4").Value
            fechaEmision = ws.Range("D10").Value
            direccion = ws.Range("O14").Value
            matricula = ws.Range("O83").Value
            disenador = ws.Range("D24").Value
            matriculaDis = ws.Range("R24").Value
            declarante = ws.Range("D25").Value
            matriculaDec = ws.Range("R27").Value
            alcance = ws.Range("A64").Value
            
        Case "Subestación"
            inspector = ws.Range("O86").Value
            numero = ws.Range("R4").Value
            fechaEmision = ws.Range("E4").Value
            direccion = "" 'no aplica
            matricula = ws.Range("O90").Value
            disenador = ws.Range("D23").Value
            matriculaDis = ws.Range("S23").Value
            declarante = ws.Range("D24").Value
            matriculaDec = ws.Range("S24").Value
            alcance = ws.Range("A77").Value
            
        Case "Uso Final"
            inspector = ws.Range("O78").Value
            numero = ws.Range("Q4").Value
            fechaEmision = ws.Range("E4").Value
            direccion = ws.Range("O14").Value
            matricula = ws.Range("O82").Value
            disenador = ws.Range("D23").Value
            matriculaDis = ws.Range("R23").Value
            declarante = ws.Range("D24").Value
            matriculaDec = ws.Range("R24").Value
            alcance = ws.Range("A69").Value
            
    End Select
End Sub

Sub extraer_de_control()

    On Error GoTo ManejarError
    ConfigurarRendimiento False
    
    Dim dlg As FileDialog
    Dim nombre As String
    Dim wbFuente As Workbook, wsFuente As Worksheet
    Dim wbDestino As Workbook, wsDestino As Worksheet
    Dim lngCount As Long, F As Long
    Dim Contador_filas_TOTAL As Long, Contador_filas_TOTAL_2 As Long
    
    Set wbDestino = ThisWorkbook
    Set wsDestino = wbDestino.Worksheets(1)
    
    '🧮 Contadores iniciales
    Contador_filas_TOTAL = wsDestino.Range("B2").Value
    Contador_filas_TOTAL_2 = wsDestino.Range("D2").Value
    
    '📂 Seleccionar archivo fuente
    Set dlg = Application.FileDialog(msoFileDialogOpen)
    With dlg
        .AllowMultiSelect = False
        .Title = "Seleccionar archivo de control"
        .Filters.Clear
        .Filters.Add "Archivos Excel", "*.xls; *.xlsx; *.xlsm"
        
        If .Show = -1 Then
            nombre = .SelectedItems(1)
            Set wbFuente = Workbooks.Open(nombre)
            Set wsFuente = wbFuente.Sheets(1)
            
            Debug.Print "📂 Archivo abierto: " & wbFuente.Name
            
            '🔁 Recorrer las filas del rango indicado
            For F = Contador_filas_TOTAL_2 To Contador_filas_TOTAL
                
                Contador_filas_TOTAL_2 = Contador_filas_TOTAL_2 + 1
                
                '💾 Cargar datos desde el archivo fuente
                Dim numero_inspeccion As Variant, nombre_proyecto As Variant, propietario As Variant
                Dim contacto As Variant, numero_cotizacion As Variant, numero_municipio As Variant
                Dim numero_departamento As Variant, Regional As Variant, direccion_proyecto As Variant
                Dim instalacion As Variant, cedula_inspector As Variant, nombre_comercial As Variant
                Dim reglamento As Variant, cedula_constructor As Variant, nombre_constructor As Variant
                
                numero_inspeccion = wsFuente.Range("G21").Value
                nombre_proyecto = wsFuente.Range("K7").Value
                propietario = wsFuente.Range("B11").Value
                contacto = wsFuente.Range("B9").Value
                numero_cotizacion = wsFuente.Range("B21").Value
                numero_municipio = wsFuente.Range("M19").Value
                numero_departamento = wsFuente.Range("B19").Value
                Regional = wsFuente.Range("B17").Value
                direccion_proyecto = wsFuente.Range("B15").Value
                instalacion = wsFuente.Range("N9").Value
                cedula_inspector = wsFuente.Range("K13").Value
                nombre_comercial = wsFuente.Range("M21").Value
                reglamento = wsFuente.Range("L15").Value
                cedula_constructor = wsFuente.Range("K26").Value
                nombre_constructor = wsFuente.Range("F26").Value
                
                '⚙️ Normalizar datos
                Dim Lugar_emision As String, Estado_dictamen As String
                Lugar_emision = "BOGOTA"
                Estado_dictamen = "APROBADO"
                
                If LCase(instalacion) = "nuevo" Then
                    instalacion = "Nueva"
                Else
                    instalacion = "En funcionamiento"
                End If

                '🔎 Imprimir variables principales en la ventana inmediata
                Debug.Print "Fila " & F + 5 & ":"
                Debug.Print "  • Nº Inspección: " & numero_inspeccion
                Debug.Print "  • Proyecto: " & nombre_proyecto
                Debug.Print "  • Propietario: " & propietario
                Debug.Print "  • Contacto: " & contacto
                Debug.Print "  • Cotización: " & numero_cotizacion
                Debug.Print "  • Municipio: " & numero_municipio & " | Departamento: " & numero_departamento
                Debug.Print "  • Regional: " & Regional
                Debug.Print "  • Dirección: " & direccion_proyecto
                Debug.Print "  • Instalación: " & instalacion
                Debug.Print "  • Cédula inspector: " & cedula_inspector
                Debug.Print "  • Comercial: " & nombre_comercial
                Debug.Print "  • Reglamento: " & reglamento
                Debug.Print "  • Constructor: " & nombre_constructor & " | CC: " & cedula_constructor
                Debug.Print "-----------------------------------------------"
                

                '💾 Escribir datos en la hoja destino (REGISTRO DICTÁMENES)
                With wsDestino
                    .Range("C5").Offset(F, 0).Value = numero_inspeccion
                    .Range("H5").Offset(F, 0).Value = nombre_proyecto
                    .Range("I5").Offset(F, 0).Value = propietario
                    .Range("L5").Offset(F, 0).Value = contacto
                    .Range("G5").Offset(F, 0).Value = numero_cotizacion
                    .Range("M5").Offset(F, 0).Value = numero_municipio
                    .Range("N5").Offset(F, 0).Value = numero_departamento
                    .Range("D5").Offset(F, 0).Value = Lugar_emision
                    .Range("Q5").Offset(F, 0).Value = Estado_dictamen
                    .Range("P5").Offset(F, 0).Value = Regional
                    .Range("K5").Offset(F, 0).Value = direccion_proyecto
                    .Range("BD5").Offset(F, 0).Value = instalacion
                    .Range("AT5").Offset(F, 0).Value = cedula_inspector
                    .Range("V5").Offset(F, 0).Value = nombre_comercial
                    .Range("AU5").Offset(F, 0).Value = reglamento
                    .Range("AZ5").Offset(F, 0).Value = cedula_constructor
                    .Range("BA5").Offset(F, 0).Value = nombre_constructor
                End With
                
                '📊 Fórmulas dinámicas
                Dim filaExcel As Long: filaExcel = F + 5
                
                With wsDestino
                    .Range("AV5").Offset(F, 0).FormulaLocal = "=BUSCARV(AU" & filaExcel & ";Hoja2!A1:B4;2;0)"
                    .Range("AW5").Offset(F, 0).FormulaLocal = "=BUSCARV(J" & filaExcel & ";Hoja2!A10:B13;2;0)"
                    .Range("AX5").Offset(F, 0).FormulaLocal = "=BUSCARV(AW" & filaExcel & ";Hoja2!D1:E14;2;0)"
                    .Range("AY5").Offset(F, 0).FormulaLocal = "=AX" & filaExcel & " & "" - "" & AW" & filaExcel
                    .Range("BB5").Offset(F, 0).FormulaLocal = "=BUSCARV(M" & filaExcel & ";Hoja2!L:M;2;0)"
                    .Range("BC5").Offset(F, 0).FormulaLocal = "=BUSCARV(N" & filaExcel & ";Hoja2!H:I;2;0)"
                    .Range("BE5").Offset(F, 0).FormulaLocal = "=BUSCARV(BC" & filaExcel & ";DEPARTAMENTOS_Y_MUNICIPIOS!A6:B39;2;0)"
                    .Range("BF5").Offset(F, 0).FormulaLocal = "=BUSCARV(BB" & filaExcel & ";DEPARTAMENTOS_Y_MUNICIPIOS!G7:H1127;2;0)"
                End With
                
                Debug.Print "Fila procesada: " & filaExcel & " - Proyecto: " & nombre_proyecto
                
            Next F
            
            wbFuente.Close SaveChanges:=False
            wsDestino.Range("D2").Value = Contador_filas_TOTAL_2
            
            ' Transferir a SICERCO
            Transferir_a_SICERCO wsDestino, wbDestino
            
        Else
            MsgBox "No se seleccionó ningún archivo.", vbExclamation
        End If
    End With
    
  Finalizar:
      ConfigurarRendimiento True
      Exit Sub
      
  ManejarError:
      MsgBox "⚠️ Error: " & Err.Description, vbExclamation
      Resume Finalizar
End Sub

Private Sub Transferir_a_SICERCO(wsOrigen As Worksheet, wb As Workbook)
    Dim wsSicerco As Worksheet
    Set wsSicerco = wb.Sheets("SICERCO")
    
    Dim fdictamen As Long, fsic As Long, i As Long
    fdictamen = 6
    fsic = 4
    i = 0
    
    wsOrigen.Activate
    wsOrigen.Range("B6").Select
    
    Debug.Print "🔄 Iniciando traslado a SICERCO..."
    
    While Not IsEmpty(ActiveCell.Value)
        wsSicerco.Cells(fsic + i, 1).Value = Cells(fdictamen + i, 2).Value   'Número dictamen
        wsSicerco.Cells(fsic + i, 2).Value = Cells(fdictamen + i, 5).Value   'Fecha expedición
        wsSicerco.Cells(fsic + i, 3).Value = "CC"
        wsSicerco.Cells(fsic + i, 4).Value = Cells(fdictamen + i, 52).Value
        wsSicerco.Cells(fsic + i, 5).Value = Cells(fdictamen + i, 53).Value
        wsSicerco.Cells(fsic + i, 6).FormulaLocal = "=CONCATENAR(A" & fsic + i & ";"".pdf"")"
        wsSicerco.Cells(fsic + i, 7).Value = Cells(fdictamen + i, 48).Value
        wsSicerco.Cells(fsic + i, 8).Value = Cells(fdictamen + i, 51).Value
        wsSicerco.Cells(fsic + i, 9).Value = "Nueva"
        wsSicerco.Cells(fsic + i, 10).Value = "N/A"
        wsSicerco.Cells(fsic + i, 11).Value = Cells(fdictamen + i, 57).Value
        wsSicerco.Cells(fsic + i, 12).Value = Cells(fdictamen + i, 58).Value
        wsSicerco.Cells(fsic + i, 13).Value = Cells(fdictamen + i, 11).Value
        wsSicerco.Cells(fsic + i, 14).Value = "CC"
        wsSicerco.Cells(fsic + i, 15).Value = Cells(fdictamen + i, 46).Value
        wsSicerco.Cells(fsic + i, 16).FormulaLocal = "=BUSCARV(O" & fsic + i & ";Hoja2!R:S;2;0)"
        
        ActiveCell.Offset(1, 0).Select
        i = i + 1
    Wend
    
    Debug.Print "✅ Traslado completado: " & i & " filas transferidas a SICERCO."
End Sub
