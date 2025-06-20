## 🔍 Macro BuscarFacturasLMA
Esta macro cruza datos entre un archivo exportado desde NetSuite y otro archivo activo con órdenes y facturas, completando datos faltantes y generando reportes.


## 🚀 ¿Qué hace?

- Busca órdenes de servicio en el archivo `Datos Netsuite 2.xlsx`.
- Completa columnas como número de factura, estado y comercial.
- Genera un archivo de reporte con registros no completados.

## 💻 Código Visual Basic
```visualbasic
Attribute VB_Name = "M�dulo7"
Option Explicit
Function WokbookOpen(name As String) As Boolean
    Dim xWb As Workbook
    On Error Resume Next
    Set xWb = Application.Workbooks.Item(name)
    WokbookOpen = (Not xWb Is Nothing)
End Function

Function crearLibro(nombre As String) As String
    Dim libroReporte As Workbook
    Application.DisplayAlerts = False
    
    On Error Resume Next
    Workbooks.Add
    
    Set libroReporte = ActiveWorkbook
    
    With libroReporte
        .SaveAs Filename:="\\14.14.14.1\Comunicados\TICS\PROGRAMASEQUIPOSNUEVOS\Temporal\Facturas - Programacion en sitio - " & nombre & ".xlsx"
        .Worksheets(1).Range("A1").value = "Fecha Prog."
        .Worksheets(1).Range("B1").value = "N� Cotizacion"
        .Worksheets(1).Range("C1").value = "Valor cotizado"
        .Worksheets(1).Range("D1").value = "Empresa"
        .Worksheets(1).Range("E1").value = "Rep. Ventas"
        .Worksheets(1).Range("F1").value = "N� Pedido"
        .Worksheets(1).Range("G1").value = "Estado de Pedido"
        .Worksheets(1).Range("H1").value = "Lote"
        .Worksheets(1).Range("I1").value = "N� Factura"
        .Worksheets(1).Range("J1").value = "Estado de Factura"
        .Worksheets(1).Range("K1").value = "Valor Facturado"
        Application.DisplayAlerts = True
        crearLibro = .name
    End With

End Function

Function EncontrarValor(hoja As Worksheet, value As Variant, letra As String, numero As Integer) As Variant
    Dim hojaBuscar As Worksheet
    Dim rango As Range
    Dim buscado, _
        fila, _
        columna, _
        cell
    Dim filas_encontradas As String
        
    Set hojaBuscar = hoja
    fila = hojaBuscar.Cells.SpecialCells(xlLastCell).Row
    columna = hojaBuscar.Cells.SpecialCells(xlLastCell).Column
    On Error Resume Next
    hojaBuscar.Activate
    
    On Error Resume Next
    With hojaBuscar.Range(letra & "2", Cells(fila, numero))
    
        Set buscado = .Find(what:=value, LookIn:=xlValues, LookAt _
        :=xlWhole, SearchOrder:=xlByColumns)
    
        If (Not buscado Is Nothing) And (IsNull(buscado) = False) Then
            Do
                If (InStr(1, filas_encontradas, buscado.Row) > 0) Then
                    Exit Do
                End If
                filas_encontradas = filas_encontradas & "/" & buscado.Row
                Set buscado = .FindNext(buscado)
    '           Set rango = Cells(fila, columna)
            Loop While Not buscado Is Nothing
            EncontrarValor = filas_encontradas
        Else
            EncontrarValor = Empty
        End If
    
    End With
    
    fila = 0
    columna = 0
    
    Exit Function
    
End Function

Sub PruebaMacro()
'
' PruebaMacro Macro
'

Application.ScreenUpdating = False

'VARIABLES DE REPORTE FACTURA
Dim cotizacion As String, valor_cotizado As String, cliente As String, rep_ventas As String, pedido As String, estado_pedido As String, _
    lote As String, factura As String, estado_factura As String, valor_facturado As Variant, factura_anterior As String, _
    valor_facturado_dos As Variant, comparar_valor_facturado As Variant, factura_flujo As String, valor_cotizado_dos As String, _
    factura_actual As String
    
'    LISTA DE VALORES DE REPORTE FACTURAS
Dim list_cotizacion As String, list_valor_cotizado As String, list_cliente As String, list_rep_ventas As String, _
    list_pedido As String, list_estado_pedido As String, list_lote As String, list_factura As String, _
    list_estado_factura As String, list_valor_facturado As String, list_results As String

Dim buscar_columna As String, valor As String, columna_ref As String, fila_ref As String, nombre_libro_resultados As String, _
    lista_registros() As String, nombre_hoja As String, mes As String, dia As String, dia_registro As String, fecha_registro As String
    
Dim es_lote_columna As Boolean, es_abierto As Boolean

Dim conteo_columnas As Integer, index As Integer, largo_lista As Integer, indice As Integer, count As Integer, i As Integer, _
    largo_filas_encontradas As Integer, recuento_filas As Integer, conteo_campos As Integer, bajar_filas As Integer

Dim filas_encontradas As Variant, ultima_fila As Variant, array_campos() As String, num_pedido As Variant, _
    num_pedido_anterior As Variant, dias_array As Variant

Dim libro_de_datos As Workbook, libro_de_agenda As Workbook, wb As Workbook, libro_resultados As Workbook

Dim buscar_en_hoja As Worksheet


'CAMBIAR VALOR
es_abierto = WokbookOpen("Prog en Sitio.xlsx")

'SI EL LIBRO DE DATOS NO ESTA BIERTO LO ABRE
If es_abierto Then
    On Error GoTo errorLibro
'    CAMBIAR VALOR
    Set libro_de_datos = Workbooks("Prog en Sitio.xlsx")
Else
    On Error GoTo errorLibro
'    CAMBIAR VALOR
    Set libro_de_datos = Workbooks.Open("C:\Users\william.enciso\Documents\Excel\Prog en Sitio.xlsx")
End If

For Each wb In Workbooks
'CAMBIAR VALOR
    If wb.name = "Prog en Sitio.xlsx" Then
        wb.Activate
        Set buscar_en_hoja = wb.Sheets(1)
        On Error GoTo errorDesconocido
        ActiveWorkbook.RefreshAll
        DoEvents
        Exit For
    End If
Next wb

Set libro_de_agenda = Workbooks("PROGRAMACION EN SITIO 2019.xlsx")
libro_de_agenda.Activate

conteo_columnas = 7

With libro_de_agenda.ActiveSheet
    ultima_fila = .Cells(.Rows.count, "A").End(xlUp).Row
    .Range("A1").Select
    mes = ActiveCell.Offset(0, 1).value
    ActiveCell.Offset(1, 0).Select
    For index = 1 To conteo_columnas
        ActiveCell.Offset(0, 1).Select
        dia = dia & "," & ActiveCell.value
    Next index
    dias_array = Split(dia, ",")
    .Range("A3").Select
    nombre_hoja = .name
End With

On Error GoTo errorLibroResultados
nombre_libro_resultados = crearLibro(nombre_hoja)
Set libro_resultados = Workbooks(nombre_libro_resultados)

buscar_columna = "LOTE"
es_lote_columna = False

libro_de_agenda.Activate
Do Until es_lote_columna

    If conteo_campos = 26 Then
        dia = ""
        bajar_filas = 2
        conteo_campos = 1
        fila_ref = ActiveCell.Row
        mes = ActiveCell.Offset(0, 1).value
        If mes = "" Then
            ActiveCell.Offset(1, 0).Select
            mes = ActiveCell.Offset(0, 1).value
            bajar_filas = 3
        End If
        ActiveCell.Offset(1, 0).Select
        For index = 1 To conteo_columnas
            ActiveCell.Offset(0, 1).Select
            dia = dia & "," & ActiveCell.value
        Next index
        dias_array = Split(dia, ",")
        Range("A" & fila_ref + bajar_filas).Select
    Else
        conteo_campos = conteo_campos + 1
    End If

    If ActiveCell.value = buscar_columna Then
    
        fila_ref = ActiveCell.Row
        For index = 1 To conteo_columnas
        
            libro_de_agenda.Activate
        
            list_cliente = Empty
            list_cotizacion = Empty
            list_estado_factura = Empty
            list_estado_pedido = Empty
            list_factura = Empty
            list_lote = Empty
            list_pedido = Empty
            list_rep_ventas = Empty
            list_valor_cotizado = Empty
            list_valor_facturado = Empty
            
            cliente = ""
            cotizacion = ""
            estado_factura = ""
            estado_pedido = ""
            factura = Empty
            lote = ""
            pedido = ""
            rep_ventas = ""
            valor_cotizado = Empty
            valor_cotizado_dos = Empty
            valor_facturado = Empty
            valor_facturado_dos = Empty
        
            ActiveCell.Offset(0, 2).Select
            
            valor = ActiveCell.value
            valor = Replace(valor, " ", "")
            
'        BUSCAR SI TIENE VARIOS VALORES EN LA CELDA
            If (InStr(1, valor, "/") > 0) Then
                lista_registros() = Split(valor, "/")
            ElseIf (InStr(1, valor, "|") > 0) Then
                lista_registros() = Split(valor, "|")
            ElseIf (InStr(1, valor, ",") > 0) Then
                lista_registros() = Split(valor, ",")
            Else
                lista_registros() = Split(valor)
            End If
            
            largo_lista = UBound(lista_registros)
            
            If Not largo_lista = -1 Then
            
                For indice = 0 To largo_lista
                
                    cliente = ""
                    cotizacion = ""
                    estado_factura = ""
                    estado_pedido = ""
                    factura = Empty
                    lote = ""
                    pedido = ""
                    rep_ventas = ""
                    valor_cotizado = Empty
                    valor_cotizado_dos = Empty
                    valor_facturado = Empty
                    valor_facturado_dos = Empty
        
                    valor = lista_registros(indice)
                    
                    If (valor Like "*[0-9]" = True) Then
'            SEPARAR EL LOTE DEL A�O
                        num_pedido = Split(valor, "-")
                        num_pedido = num_pedido(0)
                        num_pedido = Replace(num_pedido, " ", "")
                        
                        If (num_pedido = num_pedido_anterior) Then
                            GoTo next_indice
                        End If
                        
                        filas_encontradas = EncontrarValor(buscar_en_hoja, num_pedido, "A", 1)
                        
                        If (IsEmpty(filas_encontradas) = True) Then
                            filas_encontradas = EncontrarValor(buscar_en_hoja, num_pedido, "K", 11)
                        End If
                        
                        
                        If (IsEmpty(filas_encontradas) = False) Then
                        
                            filas_encontradas = Replace(filas_encontradas, "/", "", 1, 1)
                            filas_encontradas = Split(filas_encontradas, "/")
                            largo_filas_encontradas = UBound(filas_encontradas)
                            
                            dia_registro = dias_array(index)
                            fecha_registro = mes & " " & dia_registro
                            
                            For count = 0 To largo_filas_encontradas
                                factura_actual = buscar_en_hoja.Cells(filas_encontradas(count), 16)
                                
                                If ((factura_anterior = factura_actual) And (factura = "")) Or (factura = "") Then
                                    cotizacion = "," & buscar_en_hoja.Cells(filas_encontradas(count), 1)
                                ElseIf (factura_anterior <> factura_actual) Then
                                    cotizacion = cotizacion & ", " & buscar_en_hoja.Cells(filas_encontradas(count), 1)
                                End If
                                
                                If ((factura_anterior = factura_actual) And (factura = "")) Or (factura = "") Then
                                    
                                    For i = 0 To largo_filas_encontradas
                                        comparar_valor_facturado = buscar_en_hoja.Cells(filas_encontradas(i), 5)
                                        factura_flujo = buscar_en_hoja.Cells(filas_encontradas(i), 16)
                                        
                                        If (valor_cotizado <= comparar_valor_facturado) And (factura_flujo = factura_actual) Then
                                            valor_cotizado = comparar_valor_facturado
                                        End If
                                    Next i
                                    
                                ElseIf (factura_anterior <> factura_actual) Then
                                    
                                    For i = 0 To largo_filas_encontradas
                                        comparar_valor_facturado = buscar_en_hoja.Cells(filas_encontradas(i), 5)
                                        factura_flujo = buscar_en_hoja.Cells(filas_encontradas(i), 16)
                                        
                                        If (valor_cotizado_dos <= comparar_valor_facturado) And (factura_flujo = factura_actual) Then
                                            valor_cotizado_dos = comparar_valor_facturado
                                        End If
                                    Next i
                                    
                                    If (IsEmpty(valor_cotizado_dos) = False) Then
                                        valor_cotizado = valor_cotizado & ", " & valor_cotizado_dos
                                    End If
                                End If
                                
                                If ((factura_anterior = factura_actual) And (factura = "")) Or (factura = "") Then
                                    cliente = "," & buscar_en_hoja.Cells(filas_encontradas(count), 3)
                                ElseIf (factura_anterior <> factura_actual) Then
                                    cliente = cliente & ", " & buscar_en_hoja.Cells(filas_encontradas(count), 3)
                                End If
                                
                                If ((factura_anterior = factura_actual) And (factura = "")) Or (factura = "") Then
                                    rep_ventas = "," & buscar_en_hoja.Cells(filas_encontradas(count), 4)
                                ElseIf (factura_anterior <> factura_actual) Then
                                    rep_ventas = rep_ventas & ", " & buscar_en_hoja.Cells(filas_encontradas(count), 4)
                                End If
                                
                                If ((factura_anterior = factura_actual) And (factura = "")) Or (factura = "") Then
                                    pedido = "," & buscar_en_hoja.Cells(filas_encontradas(count), 10)
                                ElseIf (factura_anterior <> factura_actual) Then
                                    pedido = pedido & ", " & buscar_en_hoja.Cells(filas_encontradas(count), 10)
                                End If
                                
                                If ((factura_anterior = factura_actual) And (factura = "")) Or (factura = "") Then
                                    estado_pedido = "," & buscar_en_hoja.Cells(filas_encontradas(count), 12)
                                ElseIf (factura_anterior <> factura_actual) Then
                                    estado_pedido = estado_pedido & ", " & buscar_en_hoja.Cells(filas_encontradas(count), 12)
                                End If
                                
                                If ((factura_anterior = factura_actual) And (factura = "")) Or (factura = "") Then
                                    lote = "," & buscar_en_hoja.Cells(filas_encontradas(count), 11)
                                ElseIf (factura_anterior <> factura_actual) Then
                                    lote = lote & ", " & buscar_en_hoja.Cells(filas_encontradas(count), 11)
                                End If
                                
                                If ((factura_anterior = factura_actual) And (factura = "")) Or (factura = "") Then
                                    estado_factura = "," & buscar_en_hoja.Cells(filas_encontradas(count), 17)
                                ElseIf (factura_anterior <> factura_actual) Then
                                    estado_factura = estado_factura & ", " & buscar_en_hoja.Cells(filas_encontradas(count), 17)
                                End If
                                
                                If ((factura_anterior = factura_actual) And (factura = "")) Or (factura = "") Then
'                                LOGICA DE VALOR

                                    For i = 0 To largo_filas_encontradas
                                        comparar_valor_facturado = buscar_en_hoja.Cells(filas_encontradas(i), 18)
                                        factura_flujo = buscar_en_hoja.Cells(filas_encontradas(i), 16)
                                        
                                        If (valor_facturado <= comparar_valor_facturado) And (factura_flujo = factura_actual) Then
                                            valor_facturado = comparar_valor_facturado
                                            
                                            comparar_valor_facturado = buscar_en_hoja.Cells(filas_encontradas(i), 19)
                                            
                                            If (valor_facturado <= comparar_valor_facturado) And (factura_flujo = factura_actual) Then
                                                valor_facturado = comparar_valor_facturado
                                            End If
                                            
                                        End If
                                    Next i
                                    
                                ElseIf (factura_anterior <> factura_actual) Then
                                
                                    For i = 0 To largo_filas_encontradas
                                        comparar_valor_facturado = buscar_en_hoja.Cells(filas_encontradas(i), 18)
                                        factura_flujo = buscar_en_hoja.Cells(filas_encontradas(i), 16)
                                        
                                        If (valor_facturado_dos <= comparar_valor_facturado) And (factura_flujo = factura_actual) Then
                                            valor_facturado_dos = comparar_valor_facturado
                                            
                                            comparar_valor_facturado = buscar_en_hoja.Cells(filas_encontradas(i), 19)
                                            
                                            If (valor_facturado_dos <= comparar_valor_facturado) And (factura_flujo = factura_actual) Then
                                                valor_facturado_dos = comparar_valor_facturado
                                            End If
                                            
                                        End If
                                    Next i
                                    
                                    If (IsEmpty(valor_facturado_dos) = False) Then
                                        valor_facturado = valor_facturado & ", " & valor_facturado_dos
                                    End If
                                    
                                End If
                                
                                If ((factura_anterior = factura_actual) And (factura = "")) Or (factura = "") Then
                                    factura = "," & buscar_en_hoja.Cells(filas_encontradas(count), 16)
                                ElseIf (factura_anterior <> factura_actual) Then
                                    factura = factura & ", " & buscar_en_hoja.Cells(filas_encontradas(count), 16)
                                End If
                                
                                factura_anterior = buscar_en_hoja.Cells(filas_encontradas(count), 16)
                                
                            Next count
                            
                            If (InStr(1, cotizacion, ",") > 0) Then
                                cotizacion = Replace(cotizacion, ",", "", 1, 1)
                            End If
                            
'                            If (InStr(1, valor_cotizado, ",") > 0) Then
'                                valor_cotizado = Replace(valor_cotizado, ",", "", 1, 1)
'                            End If
                            
                            If (InStr(1, cliente, ",") > 0) Then
                                cliente = Replace(cliente, ",", "", 1, 1)
                            End If
                            
                            If (InStr(1, rep_ventas, ",") > 0) Then
                                rep_ventas = Replace(rep_ventas, ",", "", 1, 1)
                            End If
                            
                            If (InStr(1, pedido, ",") > 0) Then
                                pedido = Replace(pedido, ",", "", 1, 1)
                            End If
                            
                            If (InStr(1, estado_pedido, ",") > 0) Then
                                estado_pedido = Replace(estado_pedido, ",", "", 1, 1)
                            End If
                            
                            If (InStr(1, lote, ",") > 0) Then
                                lote = Replace(lote, ",", "", 1, 1)
                            End If
                            
                            If (InStr(1, factura, ",") > 0) Then
                                factura = Replace(factura, ",", "", 1, 1)
                            End If
                            
                            If (InStr(1, estado_factura, ",") > 0) Then
                                estado_factura = Replace(estado_factura, ",", "", 1, 1)
                            End If
                            
                            list_cotizacion = list_cotizacion & " - " & cotizacion
                            list_valor_cotizado = list_valor_cotizado & " - " & valor_cotizado
                            list_cliente = list_cliente & " - " & cliente
                            list_rep_ventas = list_rep_ventas & " - " & rep_ventas
                            list_pedido = list_pedido & " - " & pedido
                            list_estado_pedido = list_estado_pedido & " - " & estado_pedido
                            list_lote = list_lote & " - " & lote
                            list_factura = list_factura & " - " & factura
                            list_estado_factura = list_estado_factura & " - " & estado_factura
                            list_valor_facturado = list_valor_facturado & " - " & valor_facturado
                            
                        End If
                        
                    End If
                      
next_indice:
                num_pedido_anterior = num_pedido
                    
                Next indice
                
                list_cotizacion = Replace(list_cotizacion, "-", "", 1, 1)
                list_valor_cotizado = Replace(list_valor_cotizado, "-", "", 1, 1)
                list_cliente = Replace(list_cliente, "-", "", 1, 1)
                list_rep_ventas = Replace(list_rep_ventas, "-", "", 1, 1)
                list_pedido = Replace(list_pedido, "-", "", 1, 1)
                list_estado_pedido = Replace(list_estado_pedido, "-", "", 1, 1)
                list_lote = Replace(list_lote, "-", "", 1, 1)
                list_factura = Replace(list_factura, "-", "", 1, 1)
                list_estado_factura = Replace(list_estado_factura, "-", "", 1, 1)
                list_valor_facturado = Replace(list_valor_facturado, "-", "", 1, 1)
            
                If (list_cotizacion <> "") Then
                    list_results = list_results & "/" & fecha_registro & "|" & list_cotizacion & "|" & list_valor_cotizado & "|" & list_cliente & "|" & list_rep_ventas & "|" & list_pedido & "|" & list_estado_pedido & "|" & list_lote & "|" & list_factura & "|" & list_estado_factura & "|" & list_valor_facturado
                    array_campos() = Split(list_results, "/")
                    
                    Call agregarValores(libro_resultados, array_campos)
                    
                    libro_de_agenda.Activate
                    list_results = Empty
                    
                    
                End If

                
            End If
            
        Next index
        
        ActiveSheet.Range("A" & fila_ref).Select
        
    End If
    
    ActiveCell.Offset(1, 0).Select
    
    If ActiveCell.Row = ultima_fila Then
        es_lote_columna = True
    End If
Loop

Application.DisplayAlerts = False

libro_de_datos.Close
libro_resultados.Close SaveChanges:=True

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End

errorLibroResultados:
    MsgBox "Error en la hoja de Resultados " & vbCrLf & Err
    Application.ScreenUpdating = True
    End
    
errorDesconocido:
    MsgBox " " & vbCrLf & Err
    Application.ScreenUpdating = True
    End
    
errorLibro:
    MsgBox "El Libro de datos no esta disponible " & vbCrLf & Err
    Application.ScreenUpdating = True
    End

End Sub

Sub agregarValores(libroResultados As Workbook, arrayValores() As String)
    Dim list_campos As Variant, ultima_fila As Variant
    Dim indice As Integer, count As Integer
    Dim libro As Workbook
    
    Dim array_columns As Variant
    array_columns = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K")
    
    Set libro = libroResultados
    libro.Activate
                    
    On Error Resume Next
    With libro.Worksheets(1)
        On Error Resume Next
        For count = 1 To UBound(arrayValores)
            ultima_fila = .Cells(.Rows.count, "A").End(xlUp).Row
            If ultima_fila = 199 Then
                MsgBox "Llegue a " & ultima_fila
            End If
            ultima_fila = ultima_fila + 1
'            .Range("A" & ultima_fila).Select
            list_campos = Split(arrayValores(count), "|")
            For indice = 0 To UBound(list_campos)
                .Range(array_columns(indice) & ultima_fila).value = list_campos(indice)
'                ActiveCell.Offset(0, 1).Select
            Next indice
        Next count
    End With
End Sub



