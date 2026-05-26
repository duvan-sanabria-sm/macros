

Option Explicit ' Es una buena práctica forzar la declaración de variables

' =========================================================================
' MACRO PRINCIPAL 1: Extracción de Dictámenes Individuales
' =========================================================================


Sub Extraer_informacion_de_Dictamenes()

    On Error GoTo ManejarError
    ConfigurarRendimiento False

    Dim dlg As FileDialog
    Dim lngCount As Long, i As Integer
    Dim Contador_filas_TOTAL As Long
    Dim nombre As String
    Dim identificacion_de_dictamen As String, clase_de_dictamen As String
    Dim Nombre_Inspector As String, numero_dictamen As String
    Dim fecha_emision As Variant, fecha_de_inspeccion As Variant
    Dim direccion_proyecto As String, matricula_inspector As String
    Dim nombre_disenador As String, matricula_disenador As String
    Dim nombre_declarante As String, matricula_declarante As String
    Dim desc_alcance As String
    Dim Subtipo As String, codigoSubtipo As String

    Dim wbDestino As Workbook
    Dim wsDestino As Worksheet
    Dim ws As Worksheet
    Dim wbOrigen As Workbook

    Set wbDestino = ThisWorkbook
    Set wsDestino = wbDestino.Worksheets(1) ' Asume que es la hoja "REGISTRO DICTÁMENES"

    Contador_filas_TOTAL = wsDestino.Range("B2").Value

    Set dlg = Application.FileDialog(msoFileDialogOpen)
    With dlg
        .AllowMultiSelect = True
        .Title = "Seleccione los dictámenes a procesar"
        .Filters.Clear
        .Filters.Add "Archivos Excel", "*.xls; *.xlsx; *.xlsm"

        If .Show = -1 Then
            For lngCount = 1 To .SelectedItems.Count
                nombre = .SelectedItems(lngCount)
                Debug.Print "Intentando abrir: " & nombre

                If Dir(nombre) = "" Then
                    MsgBox "No se encontró el archivo: " & vbCrLf & nombre, vbExclamation, "Archivo no encontrado"
                    Resume SiguienteArchivo
                End If

                On Error Resume Next
                Set wbOrigen = Workbooks.Open(nombre, ReadOnly:=True, CorruptLoad:=xlRepairFile)
                If wbOrigen Is Nothing Then
                    MsgBox "No se pudo abrir el archivo:" & vbCrLf & nombre & vbCrLf & _
                           "Error: " & Err.Description, vbExclamation, "Error al abrir"
                    Err.Clear
                    On Error GoTo ManejarError
                    Resume SiguienteArchivo
                End If
                On Error GoTo ManejarError

                ' === PROCESAR CADA HOJA VISIBLE ===
                For i = 1 To wbOrigen.Worksheets.Count
                    Set ws = wbOrigen.Worksheets(i)

                    ' ? OMITIR hojas ocultas
                    If ws.Visible <> xlSheetVisible Then GoTo SiguienteHoja

                    ' ========== LEER EL TIPO DE DICTAMEN ==========
                    Dim RangoTipo As Range
                    On Error Resume Next
                    Set RangoTipo = ws.Range("TipoInspeccion")
                    On Error GoTo ManejarError

                    If RangoTipo Is Nothing Then
                        identificacion_de_dictamen = "RANGO_NO_ENCONTRADO"
                    Else
                        identificacion_de_dictamen = Trim(RangoTipo.Value)
                    End If
                    Set RangoTipo = Nothing

                    ' ========== CLASIFICACIÓN ==========
                    Select Case identificacion_de_dictamen
                        Case "C. IDENTIFICACIÓN DE LA INSTALACIÓN DE DISTRIBUCIÓN OBJETO DEL DICTAMEN": clase_de_dictamen = "Distribución"
                        Case "C. IDENTIFICACIÓN DE LA INSTALACIÓN DE TRANSFORMACIÓN OBJETO DEL DICTAMEN": clase_de_dictamen = "Subestación"
                        Case "C. IDENTIFICACIÓN DE LA INSTALACIÓN DE USO FINAL OBJETO DEL DICTAMEN": clase_de_dictamen = "Uso Final"
                        Case "C. IDENTIFICACIÓN DEL SISTEMA DE GENERACIÓN OBJETO DEL DICTAMEN": clase_de_dictamen = "Generación"
                        Case "C. IDENTIFICACIÓN DE LA LÍNEA OBJETO DEL DICTAMEN": clase_de_dictamen = "Transmisión"
                        Case "C. IDENTIFICACIÓN DE LA INSTALACIÓN DEL SISTEMA DE ILUMINACIÓN EXTERIOR O ALUMBRADO PÚBLICO OBJETO DEL DICTÁMEN": clase_de_dictamen = "Iluminación Exterior"
                        Case "C. IDENTIFICACIÓN DE LA INSTALACION DEL SISTEMA DE ILUMINACIÓN INTERIOR OBJETO DEL DICTAMEN": clase_de_dictamen = "Iluminación Interior"
                        Case Else: clase_de_dictamen = "SIN IDENTIFICADOR"
                    End Select

                    ' ========== LECTURA DE DATOS ==========
                    If clase_de_dictamen <> "SIN IDENTIFICADOR" And clase_de_dictamen <> "RANGO_NO_ENCONTRADO" Then
                        LeerDatosPorTipo ws, clase_de_dictamen, numero_dictamen, Nombre_Inspector, _
                            direccion_proyecto, matricula_inspector, nombre_disenador, matricula_disenador, _
                            nombre_declarante, matricula_declarante, desc_alcance, fecha_emision, Subtipo, codigoSubtipo
                    Else
                        numero_dictamen = "SIN NUMERO"
                        fecha_emision = "SIN FECHA"
                        If clase_de_dictamen = "RANGO_NO_ENCONTRADO" Then
                            Nombre_Inspector = "RANGO 'TipoInspeccion' NO ENCONTRADO"
                        Else
                            Nombre_Inspector = "ERROR"
                        End If
                        direccion_proyecto = ""
                        matricula_inspector = ""
                        nombre_disenador = ""
                        matricula_disenador = ""
                        nombre_declarante = ""
                        matricula_declarante = ""
                        desc_alcance = ""
                        Subtipo = ""
                        codigoSubtipo = ""
                    End If

                    ' Validación de vacíos
                    If IsEmpty(numero_dictamen) Or numero_dictamen = "" Then numero_dictamen = "SIN NUMERO"
                    If IsEmpty(fecha_emision) Or fecha_emision = "" Then fecha_emision = "SIN FECHA"

                    ' ========== ESCRIBIR EN HOJA DESTINO ==========
                    Contador_filas_TOTAL = Contador_filas_TOTAL + 1
                    With wsDestino
                        .Range("B3").Offset(Contador_filas_TOTAL, 0).Value = numero_dictamen
                        .Range("E3").Offset(Contador_filas_TOTAL, 0).Value = fecha_emision
                        .Range("J3").Offset(Contador_filas_TOTAL, 0).Value = clase_de_dictamen
                        .Range("K3").Offset(Contador_filas_TOTAL, 0).Value = Subtipo
                        .Range("AX3").Offset(Contador_filas_TOTAL, 0).Value = codigoSubtipo
                        .Range("L3").Offset(Contador_filas_TOTAL, 0).Value = direccion_proyecto
                        .Range("P3").Offset(Contador_filas_TOTAL, 0).Value = Nombre_Inspector
                        .Range("AC3").Offset(Contador_filas_TOTAL, 0).Value = matricula_inspector
                        .Range("AL3").Offset(Contador_filas_TOTAL, 0).Value = nombre_disenador
                        .Range("AM3").Offset(Contador_filas_TOTAL, 0).Value = matricula_disenador
                        .Range("AN3").Offset(Contador_filas_TOTAL, 0).Value = nombre_declarante
                        .Range("AO3").Offset(Contador_filas_TOTAL, 0).Value = matricula_declarante
                        .Range("AH3").Offset(Contador_filas_TOTAL, 0).Value = desc_alcance
                    End With

SiguienteHoja:
                Next i

                wbOrigen.Close SaveChanges:=False
SiguienteArchivo:
            Next lngCount

            wsDestino.Range("B2").Value = Contador_filas_TOTAL
            'MsgBox "Extracción completada correctamente.", vbInformation

        Else
            MsgBox "No se seleccionaron archivos.", vbExclamation
        End If
    End With

Finalizar:
    ConfigurarRendimiento True
    Application.CutCopyMode = False
    Exit Sub

ManejarError:
    MsgBox "Error en archivo: " & nombre & vbCrLf & "Detalle: " & Err.Description, vbExclamation
    Resume SiguienteArchivo
End Sub


' =========================================================================
' MACRO PRINCIPAL 2: Extracción desde Archivo de Control
' =========================================================================
Sub extraer_de_control()

    On Error GoTo ManejarError
    ConfigurarRendimiento False
    
    Dim dlg As FileDialog
    Dim nombre As String
    Dim wbFuente As Workbook, wsFuente As Worksheet
    Dim wbDestino As Workbook, wsDestino As Worksheet
    Dim F As Long, filaExcel As Long
    Dim Contador_filas_TOTAL As Long, Contador_filas_TOTAL_2 As Long
    
    Set wbDestino = ThisWorkbook
    Set wsDestino = wbDestino.Worksheets(1)
    
    ' Contadores iniciales
    Contador_filas_TOTAL = wsDestino.Range("B2").Value
    Contador_filas_TOTAL_2 = wsDestino.Range("D2").Value
    
    ' Seleccionar archivo fuente
    Set dlg = Application.FileDialog(msoFileDialogOpen)
    With dlg
        .AllowMultiSelect = False
        .Title = "Seleccionar archivo de control"
        .Filters.Clear
        .Filters.Add "Archivos Excel", "*.xls; *.xlsx; *.xlsm"
        
        If .Show = -1 Then
            nombre = .SelectedItems(1)
            Set wbFuente = Workbooks.Open(nombre)
           
           ' Buscar la primera hoja visible del libro
            Dim hoja As Worksheet
            For Each hoja In wbFuente.Worksheets
                If hoja.Visible = xlSheetVisible Then
                    Set wsFuente = hoja
                    Exit For
                End If
            Next hoja

            ' Validar si no se encontró ninguna hoja visible
            If wsFuente Is Nothing Then
                MsgBox "?? No se encontró ninguna hoja visible en el archivo: " & wbFuente.Name, vbExclamation
                wbFuente.Close SaveChanges:=False
                GoTo Finalizar
            End If

            
            Debug.Print "? Archivo abierto: " & wbFuente.Name


            ' Recorrer filas del control
            For F = Contador_filas_TOTAL_2 To Contador_filas_TOTAL
                Contador_filas_TOTAL_2 = Contador_filas_TOTAL_2 + 1
                filaExcel = F + 3 ' ?? ahora se calcula correctamente
                
                ' Cargar datos desde el archivo fuente
                Dim numero_inspeccion As Variant, nombre_proyecto As Variant, propietario As Variant
                Dim contacto As Variant, numero_cotizacion As Variant, numero_municipio As Variant
                Dim numero_departamento As Variant, Regional As Variant
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
                instalacion = wsFuente.Range("N9").Value
                cedula_inspector = wsFuente.Range("K13").Value
                nombre_comercial = wsFuente.Range("M21").Value
                reglamento = wsFuente.Range("L15").Value
                cedula_constructor = wsFuente.Range("K29").Value
                nombre_constructor = wsFuente.Range("F29").Value
                nombre_comercial = wsFuente.Range("M23").Value
                
                ' Normalizar
                Dim Lugar_emision As String, Estado_dictamen As String
                Lugar_emision = "BOGOTA"
                Estado_dictamen = "APROBADO"
                
                If LCase(instalacion) = "nuevo" Then
                    instalacion = "Nueva"
                Else
                    instalacion = "En funcionamiento"
                End If
                
                ' Escribir datos en hoja destino (REGISTRO DICTÁMENES)
                With wsDestino
                    .Range("C3").Offset(F, 0).Value = numero_inspeccion
                    .Range("H3").Offset(F, 0).Value = nombre_proyecto
                    .Range("I3").Offset(F, 0).Value = propietario
                    .Range("M3").Offset(F, 0).Value = contacto
                    .Range("G3").Offset(F, 0).Value = numero_cotizacion
                    .Range("N3").Offset(F, 0).Value = numero_municipio
                    .Range("O3").Offset(F, 0).Value = numero_departamento
                    .Range("D3").Offset(F, 0).Value = Lugar_emision
                    .Range("R3").Offset(F, 0).Value = Estado_dictamen
                    .Range("Q3").Offset(F, 0).Value = Regional
                    .Range("BD3").Offset(F, 0).Value = instalacion
                    .Range("AT3").Offset(F, 0).Value = cedula_inspector
                    .Range("W3").Offset(F, 0).Value = nombre_comercial
                    .Range("AU3").Offset(F, 0).Value = reglamento
                    .Range("AZ3").Offset(F, 0).Value = cedula_constructor
                    .Range("BA3").Offset(F, 0).Value = nombre_constructor
                    .Range("W3").Offset(F, 0).Value = nombre_comercial

                End With
                
                ' ? Fórmulas dinámicas
                With wsDestino
                    ' 1?? Buscar valor intermedio desde AU en Hoja2!A1:B4
                    .Range("AV3").Offset(F, 0).FormulaLocal = "=BUSCARV(AU" & filaExcel & ";Hoja2!A1:B4;2;FALSO)"
                                    
                    ' 2?? Mantener descripción en AY (desde Hoja2!E17:F23)
                    .Range("AY3").Offset(F, 0).FormulaLocal = "=BUSCARV(AX" & filaExcel & ";Hoja2!E17:F35;2;FALSO)"
                                    
                    ' 3?? Mostrar número + descripción en BG
                    .Range("BG3").Offset(F, 0).FormulaLocal = "=AX" & filaExcel & " & "" - "" & AY" & filaExcel
                                    
                    ' ?? Otros lookups auxiliares
                    .Range("AW3").Offset(F, 0).FormulaLocal = "=BUSCARV(J" & filaExcel & ";Hoja2!A10:B18;2;0)"
                    .Range("BB3").Offset(F, 0).FormulaLocal = "=BUSCARV(N" & filaExcel & ";Hoja2!L:M;2;0)"
                    .Range("BC3").Offset(F, 0).FormulaLocal = "=BUSCARV(O" & filaExcel & ";Hoja2!H:I;2;0)"
                    .Range("BE3").Offset(F, 0).FormulaLocal = "=BUSCARV(BC" & filaExcel & ";DEPARTAMENTOS_Y_MUNICIPIOS!A6:B39;2;0)"
                    .Range("BF3").Offset(F, 0).FormulaLocal = "=BUSCARV(BB" & filaExcel & ";DEPARTAMENTOS_Y_MUNICIPIOS!G7:H1127;2;0)"
                End With

                Debug.Print "? Fila " & filaExcel & " procesada correctamente."
                
            Next F
            
            wbFuente.Close SaveChanges:=False
            wsDestino.Range("D2").Value = Contador_filas_TOTAL_2
            
            ' Transferir datos a hoja SICERCO
            Transferir_a_SICERCO wsDestino, wbDestino
            
        Else
            MsgBox "No se seleccionó ningún archivo.", vbExclamation
        End If
    End With
    
Finalizar:
    ConfigurarRendimiento True
    Application.CalculateFullRebuild
    Exit Sub
    
ManejarError:
    MsgBox "?? Error: " & Err.Description, vbExclamation
    Resume Finalizar
End Sub


' =========================================================================
' MACRO PRINCIPAL 3: Transferencia a Hoja SICERCO
' =========================================================================
Private Sub Transferir_a_SICERCO(wsOrigen As Worksheet, wb As Workbook)
    Dim wsSicerco As Worksheet
    Set wsSicerco = wb.Sheets("SICERCO")
    
    Dim fdictamen As Long, fsic As Long, i As Long
    Dim ultimaFila As Long
    
    fdictamen = 4     ' Fila donde comienzan los dictámenes en la hoja REGISTRO DICTÁMENES (Fila 3 + Offset 1)
    fsic = 4          ' Fila donde comienzan los registros en SICERCO
    
    ' Buscar última fila con datos en columna B (número dictamen)
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, "B").End(xlUp).Row

    Debug.Print "?? Iniciando traslado a SICERCO..."

    For i = 0 To (ultimaFila - fdictamen)
        
        ' Si la fila está completamente vacía, termina el bucle
        If Application.WorksheetFunction.CountA(wsOrigen.Rows(fdictamen + i)) = 0 Then Exit For
        
        ' --- Transferencia de datos ---
        wsSicerco.Cells(fsic + i, 1).Value = wsOrigen.Cells(fdictamen + i, 2).Value   ' Nº Dictamen (col B)
        wsSicerco.Cells(fsic + i, 2).Value = wsOrigen.Cells(fdictamen + i, 5).Value   ' Fecha emisión (col E)
        wsSicerco.Cells(fsic + i, 3).Value = "CC"                                     ' Tipo documento fijo
        wsSicerco.Cells(fsic + i, 4).Value = wsOrigen.Cells(fdictamen + i, 52).Value  ' Columna AZ (CC Constructor)
        wsSicerco.Cells(fsic + i, 5).Value = wsOrigen.Cells(fdictamen + i, 53).Value  ' Columna BA (Nombre Constructor)
        wsSicerco.Cells(fsic + i, 6).FormulaLocal = "=CONCATENAR(A" & fsic + i & ";"".pdf"")" ' Nombre archivo PDF
        
        ' ? Columna H (8): ahora toma el valor de BG (columna 59)
        wsSicerco.Cells(fsic + i, 8).Value = wsOrigen.Cells(fdictamen + i, 59).Value   ' Columna BG
        
        ' Otras columnas
        wsSicerco.Cells(fsic + i, 7).Value = wsOrigen.Cells(fdictamen + i, 48).Value   ' Columna AV (Código)
        wsSicerco.Cells(fsic + i, 9).Value = "Nueva"
        wsSicerco.Cells(fsic + i, 10).Value = "N/A"
        wsSicerco.Cells(fsic + i, 11).Value = wsOrigen.Cells(fdictamen + i, 57).Value  ' BE
        wsSicerco.Cells(fsic + i, 12).Value = wsOrigen.Cells(fdictamen + i, 58).Value  ' BF
        wsSicerco.Cells(fsic + i, 13).Value = wsOrigen.Cells(fdictamen + i, 12).Value  ' L (Dirección)
        wsSicerco.Cells(fsic + i, 14).Value = "CC"
        wsSicerco.Cells(fsic + i, 15).Value = wsOrigen.Cells(fdictamen + i, 46).Value  ' AT (Cédula Inspector)
        wsSicerco.Cells(fsic + i, 16).FormulaLocal = "=BUSCARV(O" & fsic + i & ";Hoja2!R:S;2;0)"
        
    Next i
    
    Debug.Print "? Traslado completado: " & i & " filas transferidas a SICERCO."
End Sub


' =========================================================================
' SUBRUTINA DE APOYO 1: Lectura de datos por tipo
' =========================================================================
Private Sub LeerDatosPorTipo(ws As Worksheet, tipo As String, _
    ByRef numero As String, ByRef inspector As String, ByRef direccion As String, _
    ByRef matricula As String, ByRef disenador As String, ByRef matriculaDis As String, _
    ByRef declarante As String, ByRef matriculaDec As String, ByRef alcance As String, _
    ByRef fechaEmision As Variant, ByRef Subtipo As String, ByRef codigoSubtipo As String)

        Debug.Print "variable: " & Subtipo
    
    Select Case tipo
    
        Case "Distribución"
            inspector = ws.Range("O79").Value
            numero = ws.Range("Q4").Value
            fechaEmision = ws.Range("E4").Value
            direccion = ws.Range("O14").Value
            matricula = ws.Range("O83").Value
            disenador = ws.Range("D24").Value
            matriculaDis = ws.Range("R24").Value
            declarante = ws.Range("D25").Value
            matriculaDec = ws.Range("R25").Value
            alcance = ws.Range("A70").Value
            Subtipo = ws.Range("O16").Value
            codigoSubtipo = ws.Range("AB1").Value
            
        Case "Subestación"
            inspector = ws.Range("O86").Value
            numero = ws.Range("R4").Value
            fechaEmision = ws.Range("E4").Value
            direccion = ws.Range("O14").Value
            matricula = ws.Range("O90").Value
            disenador = ws.Range("D23").Value
            matriculaDis = ws.Range("S23").Value
            declarante = ws.Range("D24").Value
            matriculaDec = ws.Range("S24").Value
            alcance = ws.Range("A77").Value
            Subtipo = ws.Range("O16").Value
            codigoSubtipo = ws.Range("AB1").Value
            
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
            Subtipo = ws.Range("G20").Value
            codigoSubtipo = ws.Range("AB1").Value
            
        ' --- INICIO DE CASOS MODIFICADOS ---
        
        Case "Generación" ' (Caso 4 de la imagen)
            inspector = ws.Range("O79").Value
            numero = ws.Range("Q4").Value
            fechaEmision = ws.Range("E4").Value
            direccion = ws.Range("O14").Value
            matricula = ws.Range("O83").Value
            disenador = ws.Range("D24").Value
            matriculaDis = ws.Range("R24").Value
            declarante = ws.Range("D25").Value
            matriculaDec = ws.Range("R25").Value
            alcance = ws.Range("A70").Value
            Subtipo = ws.Range("O16").Value
            codigoSubtipo = ws.Range("AB1").Value

        Case "Transmisión" ' (Caso 5 de la imagen)
            inspector = ws.Range("O79").Value
            numero = ws.Range("Q4").Value
            fechaEmision = ws.Range("E4").Value
            direccion = ws.Range("O14").Value
            matricula = ws.Range("O83").Value
            disenador = ws.Range("D24").Value
            matriculaDis = ws.Range("R24").Value
            declarante = ws.Range("D25").Value
            matriculaDec = ws.Range("R25").Value
            alcance = ws.Range("A70").Value
            Subtipo = ws.Range("O16").Value  ' <--- AQUÍ ESTABA EL ERROR (Decía "O1Sucho")
            codigoSubtipo = ws.Range("AB1").Value

            
        Case "Iluminación Exterior" ' (Caso 6 de la imagen)
            inspector = ws.Range("N87").Value
            numero = ws.Range("O7").Value
            fechaEmision = ws.Range("E7").Value
            direccion = ws.Range("M17").Value
            matricula = ws.Range("N91").Value
            disenador = ws.Range("C42").Value
            matriculaDis = ws.Range("P42").Value
            declarante = ws.Range("C43").Value
            matriculaDec = ws.Range("P43").Value
            alcance = ws.Range("A81").Value
            Subtipo = ws.Range("W15").Value
            codigoSubtipo = ws.Range("AB1").Value

            
        Case "Iluminación Interior" ' (Caso 7 de la imagen)
            inspector = ws.Range("K63").Value
            numero = ws.Range("L6").Value
            fechaEmision = ws.Range("C6").Value
            direccion = ws.Range("L16").Value
            matricula = ws.Range("K67").Value
            disenador = ws.Range("C25").Value
            matriculaDis = ws.Range("L25").Value
            declarante = ws.Range("C26").Value
            matriculaDec = ws.Range("L26").Value
            alcance = ws.Range("A56").Value
            Subtipo = ws.Range("P13").Value
            codigoSubtipo = ws.Range("AB1").Value

        ' --- Fin de la Modificación ---
            
    End Select
End Sub


' =========================================================================
' SUBRUTINA DE APOYO 2: Configuración de Rendimiento
' =========================================================================
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









