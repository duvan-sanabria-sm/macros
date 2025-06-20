## 🔍 Macro BuscarFacturasLME
Esta macro cruza datos entre un archivo exportado desde NetSuite y otro archivo activo con órdenes y facturas, completando datos faltantes y generando reportes.


## 🚀 ¿Qué hace?

- Busca órdenes de servicio en el archivo `Datos Netsuite 2.xlsx`.
- Completa columnas como número de factura, estado y comercial.
- Genera un archivo de reporte con registros no completados.

## 💻 Código Visual Basic
```visualbasic
Option Explicit
Function WokbookOpen(name As String) As Boolean
    Dim xWb As Workbook
    On Error Resume Next
    Set xWb = Application.Workbooks.Item(name)
    WokbookOpen = (Not xWb Is Nothing)
End Function

Function FindID(hoja As Worksheet, value As String) As Range
    Dim hojaBuscar As Worksheet
    Dim rango As Range
    Dim buscado, _
        fila, _
        columna, _
        cell
    Dim prueba As String
        
    Set hojaBuscar = hoja
    fila = hojaBuscar.Cells.SpecialCells(xlLastCell).Row
    columna = hojaBuscar.Cells.SpecialCells(xlLastCell).Column
    On Error Resume Next
    hojaBuscar.Activate
    
    On Error Resume Next
    Set buscado = hojaBuscar.Range("A1", Cells(fila, columna)).Find(what:=value, LookIn:=xlValues, LookAt _
        :=xlWhole, SearchOrder:=xlByColumns)
    
    If (Not buscado Is Nothing) And (IsNull(buscado) = False) Then
        fila = buscado.Row
        Set rango = Cells(fila, columna)
        Set FindID = rango
    Else
        Set FindID = Nothing
    End If
    
    fila = 0
    columna = 0
    
    Exit Function
    
End Function

Function crearLibro() As String
    Dim libroReporte As Workbook
    
    Application.DisplayAlerts = False
    
    On Error Resume Next
    Workbooks.Add
    
    Set libroReporte = ActiveWorkbook
    
    With libroReporte
        .SaveAs Filename:="C:\Users\duvan.sanabria\OneDrive - Servimeters\Documentos\Macros\LMA -LME - LV\resultados\LV\LV Reporte Facturas" & Format(Now(), "DD-MMM-YYYY hh mm AMPM") & ".xlsx"
        .Worksheets(1).Range("A1").value = "Nro. Orden"
        .Worksheets(1).Range("B1").value = "Cliente"
        .Worksheets(1).Range("C1").value = "Tipo"
        .Worksheets(1).Range("D1").value = "No. Lote"
        .Worksheets(1).Range("E1").value = "Pedido"
        .Worksheets(1).Range("F1").value = "No. Factura"
        .Worksheets(1).Range("G1").value = "Hoja"
        Application.DisplayAlerts = True
        crearLibro = .name
    End With

End Function

Function extraerTexto(texto As String) As String
    texto = Left(texto, Len(texto) - 1)
    extraerTexto = texto
End Function

Sub BuscarFacturasLV()
'
' BuscarFacturas Macro
'
    Application.ScreenUpdating = False

    Dim excelFacturas As Workbook, libroResultados As Workbook, excelActualizar As Workbook
    Dim hojaBuscarExcel As Worksheet
    Dim rango As Range, cell As Range, rangoEncontrado As Range
    Dim os As String, factura As String, osAnterior As String, nombreHoja As String, facAnterior As String, datoFactura As String, nombreLibroResults As String, referenciaOs As String
    Dim celdasCopiadas As String
    Dim ultimaFila As Long
    Dim count As Integer, WS_Count As Integer, I As Integer, valoresErroneos As Integer
    Dim contiene As Boolean
    Dim valorEncontrado As Range, fechaFactura As Range, estadoFactura As Range, idFactura As Range, idOS As Range, comercial_factura As Range
    Dim arrayRange As Variant
    
    Dim listaOs() As String, listaFact() As String
    Dim largoLista As Integer, J As Integer, K As Integer
    Dim variasOs As String, variasFact As String, variasFechas As String, variosEstados As String
    
    contiene = WokbookOpen("Datos Netsuite 2.xlsx")
    
    If contiene Then
'        MsgBox "The file is open", vbInformation, "Kutools for Excel"
        On Error GoTo errorLibro
        Set excelFacturas = Workbooks("Datos Netsuite 2.xlsx")
    Else
'        MsgBox "The file is not open", vbInformation, "Kutools for Excel"
        On Error GoTo errorLibro
        Set excelFacturas = Workbooks.Open("C:\Users\duvan.sanabria\OneDrive - Servimeters\Documentos\Macros\LMA -LME - LV\netsuite\Datos Netsuite 2.xlsx")
    End If
    
    On Error GoTo errorLibroResultados
    nombreLibroResults = crearLibro()
    
    Set excelActualizar = Workbooks("LV 2025 GMM-RG-54 CONT SEGUI EQUIP V1.xlsx")
    
    WS_Count = excelActualizar.Worksheets.count
    
    For I = 1 To WS_Count
        With excelActualizar.Worksheets(I)
        '''
            nombreHoja = .name
            ultimaFila = .Cells(.Rows.count, "A").End(xlUp).Row
            Set rango = .Range("N15", .Cells(ultimaFila, "N"))
            For Each cell In rango
                os = cell.value
                os = Replace(os, " ", "")
                If ((UCase$(os) Like "PEDIDO" = True) Or (UCase$(os) Like "N/A" = True)) Then
                    celdasCopiadas = celdasCopiadas & "-" & "A" & cell.Row & ",B" & cell.Row & ",G" & cell.Row & ",N" & cell.Row & ",M" & cell.Row & ",O" & cell.Row
                    arrayRange = Split(celdasCopiadas, "-")
                    GoTo Continue
                End If

                factura = cell.Offset(0, 1).value
                factura = Replace(factura, " ", "")
                
                If (InStr(1, os, "/") > 0) Then
                    listaOs() = Split(os, "/")
                ElseIf (InStr(1, os, "-") > 0) Then
                    listaOs() = Split(os, "-")
                ElseIf (InStr(1, os, ",") > 0) Then
                    listaOs() = Split(os, ",")
                Else
                    listaOs() = Split(os)
                End If
                
                largoLista = UBound(listaOs)
                
                If Not largoLista = -1 Then
                    For J = 0 To largoLista
                        os = listaOs(J)
                        os = Replace(os, " ", "")
                        
                        '''
                        If (os Like "*[0-9]" = True) Then
                        
                            If ((UCase$(factura) Like "PEND" = True) Or (StrComp(factura, "", vbBinaryCompare) = 0) Or (factura Like "*[0-9]" = True)) Then
                                If (StrComp(os, osAnterior, vbBinaryCompare) = 0) Then
                                    
                                    If (idFactura Like "*[0-9]" = True) Then
                                        referenciaOs = os
                                        datoFactura = cell.Offset(0, 2).value
                                        
                                        If (largoLista > 0) Then
                                            cell.Offset(0, 1).value = variasFact
                                            cell.Offset(0, 2).value = variosEstados
                                            cell.Offset(0, 3).value = variasFechas
                                        ElseIf ((IsNull(datoFactura) = True) Or (StrComp(datoFactura, "", vbBinaryCompare) = 0)) Then
                                            cell.Offset(0, 1).value = idFactura
                                            
                                            If (IsNull(estadoFactura) = False) Then
                                                cell.Offset(0, 2).value = estadoFactura
                                            End If
                                        
                                            If (IsNull(comercial_factura) = False) Then
                                                cell.Offset(0, 3).value = comercial_factura
                                            End If
                                        Else
                                            cell.Offset(0, 1).value = idFactura
                                            
                                            If (IsNull(comercial_factura) = False) Then
                                                cell.Offset(0, 3).value = comercial_factura
                                            End If
                                        End If
                                    Else
                                        celdasCopiadas = celdasCopiadas & "-" & "A" & cell.Row & ",B" & cell.Row & ",G" & cell.Row & ",N" & cell.Row & ",M" & cell.Row & ",O" & cell.Row
                                        arrayRange = Split(celdasCopiadas, "-")
                                    End If
                                Else
                                    If (J = 0) Then
                                        variasFact = ""
                                        variasOs = ""
                                        variasFechas = ""
                                        variosEstados = ""
                                    End If
                                    
                                    On Error GoTo errorHoja
                                    Set hojaBuscarExcel = excelFacturas.Sheets(2)
                                    Set valorEncontrado = FindID(hojaBuscarExcel, os)
                                    If Not valorEncontrado Is Nothing Then
                                        referenciaOs = hojaBuscarExcel.Cells(valorEncontrado.Row, 1)
                                        If (referenciaOs Like "*[0-9]" = True) Then
                                            On Error GoTo errorHoja
                                            Set hojaBuscarExcel = excelFacturas.Sheets(1)
                                            Set rangoEncontrado = FindID(hojaBuscarExcel, referenciaOs)
                                            If (Not rangoEncontrado Is Nothing) Then
                                                Set idFactura = hojaBuscarExcel.Cells(rangoEncontrado.Row, rangoEncontrado.Column - 1)
                                                If (idFactura Like "*[0-9]" = True) Then
                                                    osAnterior = os
                                                    facAnterior = idFactura
                                                    
                                                    
                                                    Set comercial_factura = Cells(rangoEncontrado.Row, rangoEncontrado.Column - 2)
                                                    Set estadoFactura = rangoEncontrado
                                                    
                                                    datoFactura = cell.Offset(0, 2).value
                                                    
                                                    If largoLista > 0 Then
                                                        If J = largoLista Then
                                                            variasFact = variasFact & idFactura
                                                            variosEstados = variosEstados & estadoFactura
                                                            variasFechas = variasFechas & comercial_factura
                                                        Else
                                                            variasFact = variasFact & idFactura & "/ "
                                                            variosEstados = variosEstados & estadoFactura & "/ "
                                                            variasFechas = variasFechas & comercial_factura & "/ "
                                                        End If
                                                        
                                                        
                                                        cell.Offset(0, 1).value = variasFact
                                                        cell.Offset(0, 2).value = variosEstados
                                                        cell.Offset(0, 3).value = variasFechas
                                                    ElseIf ((IsNull(datoFactura) = True) Or (StrComp(datoFactura, "", vbBinaryCompare) = 0)) Then
                                                        cell.Offset(0, 1).value = idFactura
                                                        cell.Offset(0, 2).value = estadoFactura
                                                        cell.Offset(0, 3).value = comercial_factura
                                                    Else
                                                        cell.Offset(0, 1).value = idFactura
                                                        cell.Offset(0, 3).value = comercial_factura
                                                    End If
                                                    
                                                Else
                                                    celdasCopiadas = celdasCopiadas & "-" & "A" & cell.Row & ",B" & cell.Row & ",G" & cell.Row & ",N" & cell.Row & ",M" & cell.Row & ",O" & cell.Row
                                                    arrayRange = Split(celdasCopiadas, "-")
                                                End If
                                            Else
                                                celdasCopiadas = celdasCopiadas & "-" & "A" & cell.Row & ",B" & cell.Row & ",G" & cell.Row & ",N" & cell.Row & ",M" & cell.Row & ",O" & cell.Row
                                                arrayRange = Split(celdasCopiadas, "-")
                                            End If
                                        Else
                                            celdasCopiadas = celdasCopiadas & "-" & "A" & cell.Row & ",B" & cell.Row & ",G" & cell.Row & ",N" & cell.Row & ",M" & cell.Row & ",O" & cell.Row
                                            arrayRange = Split(celdasCopiadas, "-")
                                        End If
                                    Else
                                    'BUSCAR POR FACTURA
                                        '''
                                        
                                        If (InStr(1, factura, "/") > 0) Then
                                            listaFact() = Split(factura, "/")
                                        ElseIf (InStr(1, factura, "-") > 0) Then
                                            listaFact() = Split(factura, "-")
                                        ElseIf (InStr(1, factura, ",") > 0) Then
                                            listaFact() = Split(factura, ",")
                                        Else
                                            listaFact() = Split(factura)
                                        End If
                                        
                                        largoLista = UBound(listaFact)
                                            
                                        If Not largoLista = -1 Then
                                        
                                            For K = 0 To largoLista
                                                
                                                factura = listaFact(K)
                                                factura = Replace(factura, " ", "")
                                                
                                                If (K = 0) Then
                                                    variasFact = ""
                                                    variasOs = ""
                                                    variasFechas = ""
                                                    variosEstados = ""
                                                End If
                                                
                                                On Error GoTo errorHoja
                                                Set hojaBuscarExcel = excelFacturas.Sheets(1)
                                                Set valorEncontrado = FindID(hojaBuscarExcel, factura)
                                                If Not valorEncontrado Is Nothing Then
                                                    os = hojaBuscarExcel.Cells(valorEncontrado.Row, 3)
                                                    If (os Like "*[0-9]" = True) Then
                                                        Set comercial_factura = Cells(valorEncontrado.Row, valorEncontrado.Column - 2)
                                                        Set estadoFactura = valorEncontrado
                                                        datoFactura = cell.Offset(0, 2).value
                                                        
                                                        If largoLista > 0 Then
                                                            If K = largoLista Then
                                                                variosEstados = variosEstados & estadoFactura
                                                                variasFechas = variasFechas & comercial_factura
                                                            Else
                                                                variosEstados = variosEstados & estadoFactura & "/ "
                                                                variasFechas = variasFechas & comercial_factura & "/ "
                                                            End If
                                                            
                                                            cell.Offset(0, 2).value = variosEstados
                                                            cell.Offset(0, 3).value = variasFechas
                                                        ElseIf ((IsNull(datoFactura) = True) Or (StrComp(datoFactura, "", vbBinaryCompare) = 0)) Then
                                                            cell.Offset(0, 2).value = estadoFactura
                                                            cell.Offset(0, 3).value = comercial_factura
                                                        Else
                                                            cell.Offset(0, 3).value = comercial_factura
                                                        End If
                                                
                                                        On Error GoTo errorHoja
                                                        Set hojaBuscarExcel = excelFacturas.Sheets(2)
                                                        Set rangoEncontrado = FindID(hojaBuscarExcel, os)
                                                        If Not rangoEncontrado Is Nothing Then
                                                            Set idOS = hojaBuscarExcel.Cells(rangoEncontrado.Row, rangoEncontrado.Column - 1)
                                                            If (idOS Like "*[0-9]" = True) Then
                                                                facAnterior = factura
                                                                osAnterior = idOS
                                                                
                                                                If largoLista > 0 Then
                                                                    If K = largoLista Then
                                                                        variasOs = variasOs & idOS
                                                                    Else
                                                                        variasOs = variasOs & idOS & "/ "
                                                                    End If
                                                                    
                                                                    cell.value = variasOs
                                                                Else
                                                                    cell.value = idOS
                                                                End If
                                                            Else
                                                                celdasCopiadas = celdasCopiadas & "-" & "A" & cell.Row & ",B" & cell.Row & ",G" & cell.Row & ",N" & cell.Row & ",M" & cell.Row & ",O" & cell.Row
                                                                arrayRange = Split(celdasCopiadas, "-")
                                                            End If
                                                        Else
                                                            celdasCopiadas = celdasCopiadas & "-" & "A" & cell.Row & ",B" & cell.Row & ",G" & cell.Row & ",N" & cell.Row & ",M" & cell.Row & ",O" & cell.Row
                                                            arrayRange = Split(celdasCopiadas, "-")
                                                        End If
                                                    Else
                                                        celdasCopiadas = celdasCopiadas & "-" & "A" & cell.Row & ",B" & cell.Row & ",G" & cell.Row & ",N" & cell.Row & ",M" & cell.Row & ",O" & cell.Row
                                                        arrayRange = Split(celdasCopiadas, "-")
                                                    End If
                                                Else
                                                    celdasCopiadas = celdasCopiadas & "-" & "A" & cell.Row & ",B" & cell.Row & ",G" & cell.Row & ",N" & cell.Row & ",M" & cell.Row & ",O" & cell.Row
                                                    arrayRange = Split(celdasCopiadas, "-")
                                                End If
                                                
                                            Next K
                                        Else
                                            celdasCopiadas = celdasCopiadas & "-" & "A" & cell.Row & ",B" & cell.Row & ",G" & cell.Row & ",N" & cell.Row & ",M" & cell.Row & ",O" & cell.Row
                                            arrayRange = Split(celdasCopiadas, "-")
                                        End If
                                        '''
                                    End If
                                End If
                            Else
                                celdasCopiadas = celdasCopiadas & "-" & "A" & cell.Row & ",B" & cell.Row & ",G" & cell.Row & ",N" & cell.Row & ",M" & cell.Row & ",O" & cell.Row
                                arrayRange = Split(celdasCopiadas, "-")
                            End If
                        ElseIf (factura Like "*[0-9]" = True) Then
                            
                            If (InStr(1, factura, "/") > 0) Then
                                listaFact() = Split(factura, "/")
                            ElseIf (InStr(1, factura, "-") > 0) Then
                                listaFact() = Split(factura, "-")
                            ElseIf (InStr(1, factura, ",") > 0) Then
                                listaFact() = Split(factura, ",")
                            Else
                                listaFact() = Split(factura)
                            End If
                            
                            largoLista = UBound(listaFact)
                                
                            If Not largoLista = -1 Then
                                For K = 0 To largoLista
                                    factura = listaFact(K)
                                    factura = Replace(factura, " ", "")
                                    
                                    If (StrComp(factura, facAnterior, vbBinaryCompare) = 0) Then
                                        If (osAnterior Like "*[0-9]" = True) Then
                                        
                                            datoFactura = cell.Offset(0, 2).value
                                            
                                            If (largoLista > 0) Then
                                                cell.value = variasOs
                                                cell.Offset(0, 2).value = variosEstados
                                                cell.Offset(0, 3).value = variasFechas
                                            ElseIf ((IsNull(datoFactura) = True) Or (StrComp(datoFactura, "", vbBinaryCompare) = 0)) Then
                                                cell.value = osAnterior
                                            
                                                If (IsNull(estadoFactura) = False) Then
                                                    cell.Offset(0, 2).value = estadoFactura
                                                End If
                                            
                                                If (IsNull(comercial_factura) = False) Then
                                                    cell.Offset(0, 3).value = comercial_factura
                                                End If
                                            Else
                                                cell.value = osAnterior
                                                
                                                If (IsNull(comercial_factura) = False) Then
                                                    cell.Offset(0, 3).value = comercial_factura
                                                End If
                                            End If
                                        Else
                                                celdasCopiadas = celdasCopiadas & "-" & "A" & cell.Row & ",B" & cell.Row & ",G" & cell.Row & ",N" & cell.Row & ",M" & cell.Row & ",O" & cell.Row
                                                arrayRange = Split(celdasCopiadas, "-")
                                        End If
                                    Else
                                        If (K = 0) Then
                                            variasFact = ""
                                            variasOs = ""
                                            variasFechas = ""
                                            variosEstados = ""
                                        End If
                                        On Error GoTo errorHoja
                                        Set hojaBuscarExcel = excelFacturas.Sheets(1)
                                        Set valorEncontrado = FindID(hojaBuscarExcel, factura)
                                        If Not valorEncontrado Is Nothing Then
                                            os = hojaBuscarExcel.Cells(valorEncontrado.Row, 3)
                                            If (os Like "*[0-9]" = True) Then
                                                Set comercial_factura = Cells(valorEncontrado.Row, valorEncontrado.Column - 2)
                                                Set estadoFactura = valorEncontrado
                                                datoFactura = cell.Offset(0, 2).value
                                                
                                                If largoLista > 0 Then
                                                    
                                                    If K = largoLista Then
                                                        variosEstados = variosEstados & estadoFactura
                                                        variasFechas = variasFechas & comercial_factura
                                                    Else
                                                        variosEstados = variosEstados & estadoFactura & "/ "
                                                        variasFechas = variasFechas & comercial_factura & "/ "
                                                    End If
                                                    
                                                    cell.Offset(0, 2).value = variosEstados
                                                    cell.Offset(0, 3).value = variasFechas
                                                ElseIf ((IsNull(datoFactura) = True) Or (StrComp(datoFactura, "", vbBinaryCompare) = 0)) Then
                                                    cell.Offset(0, 2).value = estadoFactura
                                                    cell.Offset(0, 3).value = comercial_factura
                                                Else
                                                    cell.Offset(0, 3).value = comercial_factura
                                                End If
                                        
                                                On Error GoTo errorHoja
                                                Set hojaBuscarExcel = excelFacturas.Sheets(2)
                                                Set rangoEncontrado = FindID(hojaBuscarExcel, os)
                                                If Not rangoEncontrado Is Nothing Then
                                                    Set idOS = hojaBuscarExcel.Cells(rangoEncontrado.Row, rangoEncontrado.Column - 1)
                                                    If (idOS Like "*[0-9]" = True) Then
                                                        facAnterior = factura
                                                        osAnterior = idOS
                                                        
                                                        If largoLista > 0 Then
                                                            
                                                            If K = largoLista Then
                                                                variasOs = variasOs & idOS
                                                            Else
                                                                variasOs = variasOs & idOS & "/ "
                                                            End If
                                                            
                                                            cell.value = variasOs
                                                        Else
                                                            cell.value = idOS
                                                        End If
                                                    Else
                                                        celdasCopiadas = celdasCopiadas & "-" & "A" & cell.Row & ",B" & cell.Row & ",G" & cell.Row & ",N" & cell.Row & ",M" & cell.Row & ",O" & cell.Row
                                                        arrayRange = Split(celdasCopiadas, "-")
                                                    End If
                                                Else
                                                    celdasCopiadas = celdasCopiadas & "-" & "A" & cell.Row & ",B" & cell.Row & ",G" & cell.Row & ",N" & cell.Row & ",M" & cell.Row & ",O" & cell.Row
                                                    arrayRange = Split(celdasCopiadas, "-")
                                                End If
                                            Else
                                                celdasCopiadas = celdasCopiadas & "-" & "A" & cell.Row & ",B" & cell.Row & ",G" & cell.Row & ",N" & cell.Row & ",M" & cell.Row & ",O" & cell.Row
                                                arrayRange = Split(celdasCopiadas, "-")
                                            End If
                                        Else
                                            celdasCopiadas = celdasCopiadas & "-" & "A" & cell.Row & ",B" & cell.Row & ",G" & cell.Row & ",N" & cell.Row & ",M" & cell.Row & ",O" & cell.Row
                                            arrayRange = Split(celdasCopiadas, "-")
                                        End If
                                    End If
                                    
                                Next K
                            Else
                                celdasCopiadas = celdasCopiadas & "-" & "A" & cell.Row & ",B" & cell.Row & ",G" & cell.Row & ",N" & cell.Row & ",M" & cell.Row & ",O" & cell.Row
                                arrayRange = Split(celdasCopiadas, "-")
                            End If
                        Else
                            celdasCopiadas = celdasCopiadas & "-" & "A" & cell.Row & ",B" & cell.Row & ",G" & cell.Row & ",N" & cell.Row & ",M" & cell.Row & ",O" & cell.Row
                            arrayRange = Split(celdasCopiadas, "-")
                        End If
                        '''
                        
                    Next J
                Else
                    celdasCopiadas = celdasCopiadas & "-" & "A" & cell.Row & ",B" & cell.Row & ",G" & cell.Row & ",N" & cell.Row & ",M" & cell.Row & ",O" & cell.Row
                    arrayRange = Split(celdasCopiadas, "-")
                End If
Continue:
            Next cell
            
            On Error Resume Next
            For count = 1 To UBound(arrayRange)
                .Activate
                .Range(arrayRange(count)).Select
                Selection.Copy
                Windows(nombreLibroResults).Activate
                Set libroResultados = ActiveWorkbook
                On Error Resume Next
                With libroResultados.ActiveSheet
                    ultimaFila = .Cells(.Rows.count, "A").End(xlUp).Row
                    .Range("A" & ultimaFila + 1).Select
                    .Paste
                    .Range("G" & ultimaFila + 1).value = nombreHoja
                End With
                Application.CutCopyMode = False
            Next count
        
        '''
        End With
        
        celdasCopiadas = ""
        arrayRange = ""
    Next I
    
    excelFacturas.Close SaveChanges:=True
    libroResultados.Close SaveChanges:=True
    
    Application.ScreenUpdating = True
    
    End
    
errorLibro:
    MsgBox "El Libro de datos no esta disponible " & vbCrLf & Err.Description
    End
errorHoja:
    MsgBox "La hoja de Excel no esta disponible " & vbCrLf & Err.Description
    End
errorLibroResultados:
    MsgBox "Error en la hoja de Resultados " & vbCrLf & Err.Description
    End
End Sub