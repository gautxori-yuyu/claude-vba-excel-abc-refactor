# Técnicas de Optimización en VBA

Referencia detallada de técnicas para mejorar el rendimiento del código VBA.

## Control del entorno de ejecución

### Gestor de optimización reutilizable

```vba
'==============================================================
' Módulo: modOptimizacion
'--------------------------------------------------------------
' Gestión centralizada de optimización del entorno Excel
'==============================================================
Option Explicit

Private Type TEstadoAplicacion
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    Calculation As XlCalculation
    DisplayStatusBar As Boolean
    StatusBar As Variant
    DisplayAlerts As Boolean
    Cursor As XlMousePointer
End Type

Private mEstadoOriginal As TEstadoAplicacion
Private mblnOptimizado As Boolean

'--------------------------------------------------------------
' Activar modo optimizado
'--------------------------------------------------------------
Public Sub ActivarOptimizacion(Optional ByVal strMensajeEstado As String = "Procesando...")
    If mblnOptimizado Then Exit Sub
    
    ' Guardar estado original
    With Application
        mEstadoOriginal.ScreenUpdating = .ScreenUpdating
        mEstadoOriginal.EnableEvents = .EnableEvents
        mEstadoOriginal.Calculation = .Calculation
        mEstadoOriginal.DisplayStatusBar = .DisplayStatusBar
        mEstadoOriginal.StatusBar = .StatusBar
        mEstadoOriginal.DisplayAlerts = .DisplayAlerts
        mEstadoOriginal.Cursor = .Cursor
    End With
    
    ' Aplicar configuración optimizada
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayStatusBar = True
        .StatusBar = strMensajeEstado
        .DisplayAlerts = False
        .Cursor = xlWait
    End With
    
    mblnOptimizado = True
End Sub

'--------------------------------------------------------------
' Restaurar estado original
'--------------------------------------------------------------
Public Sub DesactivarOptimizacion()
    If Not mblnOptimizado Then Exit Sub
    
    With Application
        .ScreenUpdating = mEstadoOriginal.ScreenUpdating
        .EnableEvents = mEstadoOriginal.EnableEvents
        .Calculation = mEstadoOriginal.Calculation
        .DisplayStatusBar = mEstadoOriginal.DisplayStatusBar
        .StatusBar = mEstadoOriginal.StatusBar
        .DisplayAlerts = mEstadoOriginal.DisplayAlerts
        .Cursor = mEstadoOriginal.Cursor
    End With
    
    mblnOptimizado = False
End Sub

'--------------------------------------------------------------
' Actualizar barra de estado durante proceso largo
'--------------------------------------------------------------
Public Sub ActualizarProgreso(ByVal lngActual As Long, ByVal lngTotal As Long, _
                               Optional ByVal strPrefijo As String = "Procesando")
    If lngTotal > 0 Then
        Application.StatusBar = strPrefijo & ": " & _
                                Format(lngActual / lngTotal, "0%") & _
                                " (" & lngActual & " de " & lngTotal & ")"
    End If
    
    ' Permitir que Excel procese eventos del sistema cada 100 iteraciones
    If lngActual Mod 100 = 0 Then DoEvents
End Sub
```

## Procesamiento eficiente de rangos

### Leer rango a array

```vba
'--------------------------------------------------------------
' Procesar rango mediante array en memoria
'--------------------------------------------------------------
Public Function ProcesarRangoConArray(ByVal rngDatos As Range) As Variant
    Dim arrDatos As Variant
    Dim lngFila As Long
    Dim lngCol As Long
    Dim lngFilas As Long
    Dim lngCols As Long
    
    ' Leer todo el rango a memoria de una sola vez
    arrDatos = rngDatos.Value2
    
    ' Obtener dimensiones
    lngFilas = UBound(arrDatos, 1)
    lngCols = UBound(arrDatos, 2)
    
    ' Procesar en memoria (mucho más rápido)
    For lngFila = 1 To lngFilas
        For lngCol = 1 To lngCols
            ' Ejemplo: convertir a mayúsculas si es texto
            If VarType(arrDatos(lngFila, lngCol)) = vbString Then
                arrDatos(lngFila, lngCol) = UCase(arrDatos(lngFila, lngCol))
            End If
        Next lngCol
    Next lngFila
    
    ProcesarRangoConArray = arrDatos
End Function

'--------------------------------------------------------------
' Escribir array a rango
'--------------------------------------------------------------
Public Sub EscribirArrayARango(ByVal rngDestino As Range, ByRef arrDatos As Variant)
    ' Redimensionar destino al tamaño del array
    Dim lngFilas As Long
    Dim lngCols As Long
    
    lngFilas = UBound(arrDatos, 1) - LBound(arrDatos, 1) + 1
    lngCols = UBound(arrDatos, 2) - LBound(arrDatos, 2) + 1
    
    ' Escribir de una sola vez
    rngDestino.Resize(lngFilas, lngCols).Value2 = arrDatos
End Sub
```

### Comparación de rendimiento

```vba
'--------------------------------------------------------------
' Demostración: celda a celda vs array
'--------------------------------------------------------------
Public Sub CompararRendimiento()
    Const FILAS As Long = 10000
    Const COLS As Long = 10
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    
    Dim rng As Range
    Set rng = ws.Range("A1").Resize(FILAS, COLS)
    rng.Formula = "=RAND()"
    rng.Value = rng.Value ' Convertir a valores
    
    Dim dblInicio As Double
    Dim dblTiempo1 As Double
    Dim dblTiempo2 As Double
    
    ' Método 1: Celda a celda (LENTO)
    dblInicio = Timer
    Dim i As Long, j As Long
    For i = 1 To FILAS
        For j = 1 To COLS
            ws.Cells(i, j).Value = ws.Cells(i, j).Value * 2
        Next j
    Next i
    dblTiempo1 = Timer - dblInicio
    
    ' Restaurar valores
    rng.Formula = "=RAND()"
    rng.Value = rng.Value
    
    ' Método 2: Array en memoria (RÁPIDO)
    dblInicio = Timer
    Dim arr As Variant
    arr = rng.Value2
    For i = 1 To FILAS
        For j = 1 To COLS
            arr(i, j) = arr(i, j) * 2
        Next j
    Next i
    rng.Value2 = arr
    dblTiempo2 = Timer - dblInicio
    
    ' Resultado típico: Array es 50-100x más rápido
    MsgBox "Celda a celda: " & Format(dblTiempo1, "0.000") & " seg" & vbCrLf & _
           "Array: " & Format(dblTiempo2, "0.000") & " seg" & vbCrLf & _
           "Factor: " & Format(dblTiempo1 / dblTiempo2, "0.0") & "x"
    
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
End Sub
```

## Búsquedas optimizadas

### Usar Find en lugar de bucles

```vba
'--------------------------------------------------------------
' Búsqueda optimizada con Find
'--------------------------------------------------------------
Public Function BuscarValor(ByVal rngBuscar As Range, _
                            ByVal varValor As Variant) As Range
    Set BuscarValor = rngBuscar.Find( _
        What:=varValor, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)
End Function

'--------------------------------------------------------------
' Buscar todas las coincidencias
'--------------------------------------------------------------
Public Function BuscarTodos(ByVal rngBuscar As Range, _
                            ByVal varValor As Variant) As Collection
    Dim colResultados As New Collection
    Dim rngPrimero As Range
    Dim rngActual As Range
    
    Set rngPrimero = rngBuscar.Find(What:=varValor, LookAt:=xlWhole)
    
    If Not rngPrimero Is Nothing Then
        Set rngActual = rngPrimero
        Do
            colResultados.Add rngActual
            Set rngActual = rngBuscar.FindNext(rngActual)
        Loop While Not rngActual Is Nothing And rngActual.Address <> rngPrimero.Address
    End If
    
    Set BuscarTodos = colResultados
End Function
```

### Usar Match para búsquedas de índice

```vba
'--------------------------------------------------------------
' Búsqueda de posición con Match (más rápido que Find para índices)
'--------------------------------------------------------------
Public Function ObtenerFilaPorValor(ByVal rngColumna As Range, _
                                     ByVal varValor As Variant) As Long
    On Error Resume Next
    ObtenerFilaPorValor = Application.Match(varValor, rngColumna, 0)
    If Err.Number <> 0 Then ObtenerFilaPorValor = 0
    On Error GoTo 0
End Function

'--------------------------------------------------------------
' Búsqueda bidimensional con Match + Index
'--------------------------------------------------------------
Public Function BuscarEnTabla(ByVal rngTabla As Range, _
                               ByVal varBuscar As Variant, _
                               ByVal lngColBusqueda As Long, _
                               ByVal lngColResultado As Long) As Variant
    On Error Resume Next
    BuscarEnTabla = Application.Index( _
        rngTabla.Columns(lngColResultado), _
        Application.Match(varBuscar, rngTabla.Columns(lngColBusqueda), 0))
    If Err.Number <> 0 Then BuscarEnTabla = CVErr(xlErrNA)
    On Error GoTo 0
End Function
```

### Dictionary para lookups repetitivos

```vba
'--------------------------------------------------------------
' Crear índice de búsqueda con Dictionary
'--------------------------------------------------------------
Public Function CrearIndice(ByVal rngClaves As Range, _
                            ByVal rngValores As Range) As Object
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare ' Ignorar mayúsculas/minúsculas
    
    Dim arrClaves As Variant
    Dim arrValores As Variant
    arrClaves = rngClaves.Value2
    arrValores = rngValores.Value2
    
    Dim i As Long
    For i = 1 To UBound(arrClaves, 1)
        If Not dic.Exists(arrClaves(i, 1)) Then
            dic.Add arrClaves(i, 1), arrValores(i, 1)
        End If
    Next i
    
    Set CrearIndice = dic
End Function

' Uso: O(1) en lugar de O(n) para cada búsqueda
Public Sub EjemploUsoIndice()
    Dim dicClientes As Object
    Set dicClientes = CrearIndice(Range("A:A"), Range("B:B"))
    
    ' Búsquedas instantáneas
    Dim strNombre As String
    If dicClientes.Exists("CLI001") Then
        strNombre = dicClientes("CLI001")
    End If
End Sub
```

## Evitar Select y Activate

### Ejemplos de código mejorado

```vba
'--------------------------------------------------------------
' MAL: Usa Select/Activate
'--------------------------------------------------------------
Public Sub CodigoLento()
    Sheets("Datos").Select
    Range("A1").Select
    Selection.Copy
    Sheets("Destino").Select
    Range("B1").Select
    ActiveSheet.Paste
End Sub

'--------------------------------------------------------------
' BIEN: Referencias directas
'--------------------------------------------------------------
Public Sub CodigoRapido()
    ThisWorkbook.Worksheets("Datos").Range("A1").Copy _
        Destination:=ThisWorkbook.Worksheets("Destino").Range("B1")
End Sub

'--------------------------------------------------------------
' BIEN: Con variables para mayor claridad
'--------------------------------------------------------------
Public Sub CodigoRapidoClaro()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    
    Set wsOrigen = ThisWorkbook.Worksheets("Datos")
    Set wsDestino = ThisWorkbook.Worksheets("Destino")
    
    wsOrigen.Range("A1").Copy Destination:=wsDestino.Range("B1")
End Sub
```

## Optimización de bucles

### Salida anticipada

```vba
'--------------------------------------------------------------
' Salir del bucle en cuanto se encuentre el resultado
'--------------------------------------------------------------
Public Function ExisteValor(ByVal rngBuscar As Range, _
                            ByVal varValor As Variant) As Boolean
    Dim celda As Range
    
    For Each celda In rngBuscar
        If celda.Value2 = varValor Then
            ExisteValor = True
            Exit Function ' Salir inmediatamente
        End If
    Next celda
    
    ExisteValor = False
End Function
```

### Reducir iteraciones innecesarias

```vba
'--------------------------------------------------------------
' Procesar solo celdas con datos
'--------------------------------------------------------------
Public Sub ProcesarSoloCeldasConDatos(ByVal ws As Worksheet)
    Dim rngUsado As Range
    On Error Resume Next
    Set rngUsado = ws.UsedRange.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    
    If rngUsado Is Nothing Then Exit Sub
    
    Dim celda As Range
    For Each celda In rngUsado
        ' Procesar solo celdas que tienen datos
    Next celda
End Sub
```

### Procesar por bloques

```vba
'--------------------------------------------------------------
' Procesar rangos grandes por bloques para gestionar memoria
'--------------------------------------------------------------
Public Sub ProcesarPorBloques(ByVal rngTotal As Range, _
                               ByVal lngTamanioBloque As Long)
    Dim lngFilaInicio As Long
    Dim lngFilaFin As Long
    Dim lngFilasTotales As Long
    Dim rngBloque As Range
    
    lngFilasTotales = rngTotal.Rows.Count
    lngFilaInicio = 1
    
    Do While lngFilaInicio <= lngFilasTotales
        lngFilaFin = Application.Min(lngFilaInicio + lngTamanioBloque - 1, lngFilasTotales)
        
        Set rngBloque = rngTotal.Rows(lngFilaInicio & ":" & lngFilaFin)
        
        ' Procesar bloque
        ProcesarBloque rngBloque
        
        lngFilaInicio = lngFilaFin + 1
        DoEvents ' Permitir respuesta del sistema
    Loop
End Sub
```

## Value vs Value2

```vba
'--------------------------------------------------------------
' Value2 es más rápido para datos numéricos
'--------------------------------------------------------------
Public Sub DemostrarDiferenciaValue()
    ' Value convierte Currency y Date a sus tipos nativos de VBA
    ' Value2 devuelve el valor subyacente sin conversión
    
    Dim rng As Range
    Set rng = Range("A1")
    
    ' Si A1 contiene una fecha:
    ' .Value devuelve Date (tipo VBA)
    ' .Value2 devuelve Double (número de serie de Excel)
    
    ' Si A1 contiene moneda:
    ' .Value devuelve Currency (tipo VBA)
    ' .Value2 devuelve Double
    
    ' Para operaciones masivas con números, Value2 es más eficiente
    Dim arr As Variant
    arr = Range("A1:Z10000").Value2 ' Más rápido
    ' arr = Range("A1:Z10000").Value ' Más lento por conversiones
End Sub
```

## Liberación de memoria

```vba
'--------------------------------------------------------------
' Patrón de limpieza de recursos
'--------------------------------------------------------------
Public Sub ProcesoConLimpieza()
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim arr As Variant
    
    Set wb = Workbooks.Open("C:\Datos.xlsx")
    Set ws = wb.Worksheets(1)
    Set rng = ws.UsedRange
    arr = rng.Value2
    
    ' ... procesar datos ...
    
CleanExit:
    ' Liberar en orden inverso a la creación
    Erase arr
    Set rng = Nothing
    Set ws = Nothing
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
        Set wb = Nothing
    End If
    Exit Sub
    
ErrorHandler:
    ' Asegurar limpieza incluso en error
    Resume CleanExit
End Sub
```
