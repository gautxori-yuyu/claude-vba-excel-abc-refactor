# Patrones de Diseño en VBA

Referencia detallada de implementación de patrones de diseño adaptados a VBA.

## Factory Pattern

VBA no permite constructores con parámetros. El patrón Factory resuelve esto:

```vba
'==============================================================
' Clase: clsOferta
'--------------------------------------------------------------
' Representa una oferta comercial con sus datos básicos
'==============================================================
Option Explicit

Private mlngId As Long
Private mstrCliente As String
Private mdteFecha As Date
Private mdblImporte As Double

'--------------------------------------------------------------
' Propiedades
'--------------------------------------------------------------
Public Property Get Id() As Long
    Id = mlngId
End Property

Friend Property Let Id(ByVal lngValue As Long)
    mlngId = lngValue
End Property

Public Property Get Cliente() As String
    Cliente = mstrCliente
End Property

Public Property Let Cliente(ByVal strValue As String)
    mstrCliente = strValue
End Property

' ... más propiedades

'--------------------------------------------------------------
' Método de inicialización (llamado por Factory)
'--------------------------------------------------------------
Friend Sub Init(ByVal lngId As Long, ByVal strCliente As String, _
                ByVal dteFecha As Date, ByVal dblImporte As Double)
    mlngId = lngId
    mstrCliente = strCliente
    mdteFecha = dteFecha
    mdblImporte = dblImporte
End Sub
```

```vba
'==============================================================
' Módulo: modOfertaFactory
'--------------------------------------------------------------
' Factoría para crear instancias de clsOferta
'==============================================================
Option Explicit

Public Function CrearOferta(ByVal lngId As Long, ByVal strCliente As String, _
                            ByVal dteFecha As Date, ByVal dblImporte As Double) As clsOferta
    Dim objOferta As clsOferta
    Set objOferta = New clsOferta
    objOferta.Init lngId, strCliente, dteFecha, dblImporte
    Set CrearOferta = objOferta
End Function

Public Function CrearOfertaDesdeRango(ByVal rngFila As Range) As clsOferta
    Set CrearOfertaDesdeRango = CrearOferta( _
        CLng(rngFila.Cells(1, 1).Value2), _
        CStr(rngFila.Cells(1, 2).Value2), _
        CDate(rngFila.Cells(1, 3).Value), _
        CDbl(rngFila.Cells(1, 4).Value2))
End Function
```

## Singleton Pattern

Para gestores globales que deben tener una única instancia:

```vba
'==============================================================
' Módulo: modLogger
'--------------------------------------------------------------
' Singleton para gestión centralizada de logging
'==============================================================
Option Explicit

Private mobjLogger As clsLogger

Public Property Get Logger() As clsLogger
    If mobjLogger Is Nothing Then
        Set mobjLogger = New clsLogger
        mobjLogger.Init
    End If
    Set Logger = mobjLogger
End Property

Public Sub LiberarLogger()
    If Not mobjLogger Is Nothing Then
        mobjLogger.Cerrar
        Set mobjLogger = Nothing
    End If
End Sub
```

```vba
'==============================================================
' Clase: clsLogger
'--------------------------------------------------------------
' Gestiona el registro de eventos y errores
'==============================================================
Option Explicit

Private mlngNivel As Long
Private mstrRutaFichero As String
Private mobjFSO As Object ' Scripting.FileSystemObject
Private mobjFichero As Object ' Scripting.TextStream

Public Enum NivelLog
    nlError = 1
    nlWarning = 2
    nlInfo = 3
    nlDebug = 4
End Enum

Friend Sub Init(Optional ByVal strRuta As String = "", _
                Optional ByVal lngNivel As NivelLog = nlInfo)
    mlngNivel = lngNivel
    If Len(strRuta) = 0 Then
        mstrRutaFichero = ThisWorkbook.Path & "\log_" & Format(Date, "yyyymmdd") & ".txt"
    Else
        mstrRutaFichero = strRuta
    End If
    Set mobjFSO = CreateObject("Scripting.FileSystemObject")
    Set mobjFichero = mobjFSO.OpenTextFile(mstrRutaFichero, 8, True) ' ForAppending
End Sub

Public Sub Registrar(ByVal lngNivel As NivelLog, ByVal strMensaje As String, _
                     Optional ByVal strOrigen As String = "")
    If lngNivel <= mlngNivel Then
        Dim strLinea As String
        strLinea = Format(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & _
                   ObtenerEtiquetaNivel(lngNivel) & vbTab
        If Len(strOrigen) > 0 Then strLinea = strLinea & "[" & strOrigen & "] "
        strLinea = strLinea & strMensaje
        mobjFichero.WriteLine strLinea
    End If
End Sub

Private Function ObtenerEtiquetaNivel(ByVal lngNivel As NivelLog) As String
    Select Case lngNivel
        Case nlError: ObtenerEtiquetaNivel = "ERROR"
        Case nlWarning: ObtenerEtiquetaNivel = "WARN "
        Case nlInfo: ObtenerEtiquetaNivel = "INFO "
        Case nlDebug: ObtenerEtiquetaNivel = "DEBUG"
    End Select
End Function

Friend Sub Cerrar()
    If Not mobjFichero Is Nothing Then
        mobjFichero.Close
        Set mobjFichero = Nothing
    End If
    Set mobjFSO = Nothing
End Sub
```

Uso:
```vba
Logger.Registrar nlInfo, "Proceso iniciado", "modMain.ProcesoPrincipal"
Logger.Registrar nlError, "No se pudo conectar", "clsConexion.Abrir"
```

## Repository Pattern

Encapsula el acceso a datos:

```vba
'==============================================================
' Clase: clsOfertaRepository
'--------------------------------------------------------------
' Acceso a datos de ofertas desde hoja de cálculo
'==============================================================
Option Explicit

Private mwsDatos As Worksheet
Private mstrRangoTabla As String

Friend Sub Init(ByVal wsDatos As Worksheet, Optional ByVal strRango As String = "A1")
    Set mwsDatos = wsDatos
    mstrRangoTabla = strRango
End Sub

Public Function ObtenerPorId(ByVal lngId As Long) As clsOferta
    Dim rngEncontrado As Range
    Set rngEncontrado = mwsDatos.Columns(1).Find(What:=lngId, LookAt:=xlWhole)
    
    If Not rngEncontrado Is Nothing Then
        Set ObtenerPorId = modOfertaFactory.CrearOfertaDesdeRango( _
            rngEncontrado.Resize(1, 4))
    End If
End Function

Public Function ObtenerTodas() As Collection
    Dim colResultado As Collection
    Set colResultado = New Collection
    
    Dim lngFilaUlt As Long
    lngFilaUlt = mwsDatos.Cells(mwsDatos.Rows.Count, 1).End(xlUp).Row
    
    Dim lngFila As Long
    For lngFila = 2 To lngFilaUlt ' Asume cabecera en fila 1
        colResultado.Add modOfertaFactory.CrearOfertaDesdeRango( _
            mwsDatos.Range("A" & lngFila).Resize(1, 4))
    Next lngFila
    
    Set ObtenerTodas = colResultado
End Function

Public Function ObtenerPorCliente(ByVal strCliente As String) As Collection
    Dim colResultado As Collection
    Set colResultado = New Collection
    
    Dim objOferta As clsOferta
    For Each objOferta In ObtenerTodas
        If objOferta.Cliente = strCliente Then
            colResultado.Add objOferta
        End If
    Next objOferta
    
    Set ObtenerPorCliente = colResultado
End Function

Public Sub Guardar(ByVal objOferta As clsOferta)
    Dim rngExistente As Range
    Set rngExistente = mwsDatos.Columns(1).Find(What:=objOferta.Id, LookAt:=xlWhole)
    
    Dim lngFila As Long
    If rngExistente Is Nothing Then
        ' Insertar nuevo
        lngFila = mwsDatos.Cells(mwsDatos.Rows.Count, 1).End(xlUp).Row + 1
    Else
        ' Actualizar existente
        lngFila = rngExistente.Row
    End If
    
    With mwsDatos
        .Cells(lngFila, 1).Value2 = objOferta.Id
        .Cells(lngFila, 2).Value2 = objOferta.Cliente
        .Cells(lngFila, 3).Value = objOferta.Fecha
        .Cells(lngFila, 4).Value2 = objOferta.Importe
    End With
End Sub
```

## Observer Pattern con WithEvents

Para comunicación desacoplada entre componentes:

```vba
'==============================================================
' Clase: clsMonitorCambios
'--------------------------------------------------------------
' Monitoriza cambios en un rango y notifica a suscriptores
'==============================================================
Option Explicit

Public Event CambioCelda(ByVal rngCelda As Range, ByVal varValorAnterior As Variant)
Public Event CambioMultiple(ByVal rngCambios As Range)

Private mwsHoja As Worksheet
Private mrngMonitorizado As Range
Private mdicValoresAnteriores As Object ' Scripting.Dictionary

Friend Sub Init(ByVal rngMonitorizar As Range)
    Set mrngMonitorizado = rngMonitorizar
    Set mwsHoja = rngMonitorizar.Worksheet
    Set mdicValoresAnteriores = CreateObject("Scripting.Dictionary")
    
    ' Capturar valores iniciales
    Dim celda As Range
    For Each celda In mrngMonitorizado
        mdicValoresAnteriores(celda.Address) = celda.Value2
    Next celda
End Sub

Public Sub VerificarCambios()
    Dim celda As Range
    Dim varValorActual As Variant
    Dim varValorAnterior As Variant
    Dim blnHuboCambios As Boolean
    
    For Each celda In mrngMonitorizado
        varValorActual = celda.Value2
        If mdicValoresAnteriores.Exists(celda.Address) Then
            varValorAnterior = mdicValoresAnteriores(celda.Address)
            If varValorActual <> varValorAnterior Then
                RaiseEvent CambioCelda(celda, varValorAnterior)
                mdicValoresAnteriores(celda.Address) = varValorActual
                blnHuboCambios = True
            End If
        End If
    Next celda
    
    If blnHuboCambios Then
        RaiseEvent CambioMultiple(mrngMonitorizado)
    End If
End Sub
```

```vba
'==============================================================
' Clase: clsGestorValidacion
'--------------------------------------------------------------
' Suscriptor que valida cambios detectados
'==============================================================
Option Explicit

Private WithEvents mobjMonitor As clsMonitorCambios

Public Sub Suscribir(ByVal objMonitor As clsMonitorCambios)
    Set mobjMonitor = objMonitor
End Sub

Private Sub mobjMonitor_CambioCelda(ByVal rngCelda As Range, _
                                     ByVal varValorAnterior As Variant)
    ' Lógica de validación
    If Not ValidarCelda(rngCelda) Then
        MsgBox "Valor no válido en " & rngCelda.Address
        rngCelda.Value = varValorAnterior ' Revertir
    End If
End Sub

Private Function ValidarCelda(ByVal rngCelda As Range) As Boolean
    ' Implementar lógica de validación específica
    ValidarCelda = True
End Function
```

## Strategy Pattern

Para algoritmos intercambiables:

```vba
'==============================================================
' Clase: IEstrategiaCalculo (Interfaz)
'--------------------------------------------------------------
' Define contrato para estrategias de cálculo
'==============================================================
Option Explicit

Public Function Calcular(ByVal dblBase As Double, _
                         ByVal dblParametro As Double) As Double
End Function

Public Property Get Nombre() As String
End Property
```

```vba
'==============================================================
' Clase: clsCalculoLineal
'--------------------------------------------------------------
' Implementa cálculo lineal: base * parámetro
'==============================================================
Option Explicit
Implements IEstrategiaCalculo

Private Function IEstrategiaCalculo_Calcular(ByVal dblBase As Double, _
                                              ByVal dblParametro As Double) As Double
    IEstrategiaCalculo_Calcular = dblBase * dblParametro
End Function

Private Property Get IEstrategiaCalculo_Nombre() As String
    IEstrategiaCalculo_Nombre = "Cálculo Lineal"
End Property
```

```vba
'==============================================================
' Clase: clsCalculoEscalonado
'--------------------------------------------------------------
' Implementa cálculo escalonado con tramos
'==============================================================
Option Explicit
Implements IEstrategiaCalculo

Private marrTramos() As Double
Private marrFactores() As Double

Friend Sub Init(ByRef arrTramos() As Double, ByRef arrFactores() As Double)
    marrTramos = arrTramos
    marrFactores = arrFactores
End Sub

Private Function IEstrategiaCalculo_Calcular(ByVal dblBase As Double, _
                                              ByVal dblParametro As Double) As Double
    Dim dblResultado As Double
    Dim i As Long
    
    For i = UBound(marrTramos) To LBound(marrTramos) Step -1
        If dblBase >= marrTramos(i) Then
            dblResultado = dblBase * marrFactores(i) * dblParametro
            Exit For
        End If
    Next i
    
    IEstrategiaCalculo_Calcular = dblResultado
End Function

Private Property Get IEstrategiaCalculo_Nombre() As String
    IEstrategiaCalculo_Nombre = "Cálculo Escalonado"
End Property
```

Uso:
```vba
Public Sub EjemploStrategy()
    Dim objEstrategia As IEstrategiaCalculo
    Dim dblResultado As Double
    
    ' Seleccionar estrategia según contexto
    If blnUsarEscalonado Then
        Dim objEscalonado As New clsCalculoEscalonado
        objEscalonado.Init arrTramos, arrFactores
        Set objEstrategia = objEscalonado
    Else
        Set objEstrategia = New clsCalculoLineal
    End If
    
    ' Usar estrategia de forma uniforme
    dblResultado = objEstrategia.Calcular(1000, 1.5)
    Debug.Print "Resultado con " & objEstrategia.Nombre & ": " & dblResultado
End Sub
```

## MVP para UserForms

Separación de responsabilidades en formularios:

```vba
'==============================================================
' Clase: clsOfertaModelo
'--------------------------------------------------------------
' Modelo de datos para el formulario de ofertas
'==============================================================
Option Explicit

Private mobjOferta As clsOferta
Private mblnModificado As Boolean

Public Event DatosModificados()

Public Property Get Oferta() As clsOferta
    Set Oferta = mobjOferta
End Property

Public Property Set Oferta(ByVal objValue As clsOferta)
    Set mobjOferta = objValue
    mblnModificado = False
End Property

Public Property Get Modificado() As Boolean
    Modificado = mblnModificado
End Property

Public Sub ActualizarCliente(ByVal strCliente As String)
    If mobjOferta.Cliente <> strCliente Then
        mobjOferta.Cliente = strCliente
        mblnModificado = True
        RaiseEvent DatosModificados
    End If
End Sub

Public Sub ActualizarImporte(ByVal dblImporte As Double)
    If mobjOferta.Importe <> dblImporte Then
        mobjOferta.Importe = dblImporte
        mblnModificado = True
        RaiseEvent DatosModificados
    End If
End Sub
```

```vba
'==============================================================
' Clase: clsOfertaPresentador
'--------------------------------------------------------------
' Presentador que coordina vista y modelo
'==============================================================
Option Explicit

Private WithEvents mobjModelo As clsOfertaModelo
Private mobjVista As frmOferta

Public Sub Init(ByVal objVista As frmOferta, ByVal objModelo As clsOfertaModelo)
    Set mobjVista = objVista
    Set mobjModelo = objModelo
    
    ' Cargar datos iniciales en vista
    ActualizarVista
End Sub

Private Sub ActualizarVista()
    With mobjVista
        .txtCliente.Value = mobjModelo.Oferta.Cliente
        .txtImporte.Value = Format(mobjModelo.Oferta.Importe, "#,##0.00")
        .lblEstado.Caption = IIf(mobjModelo.Modificado, "Modificado", "Sin cambios")
    End With
End Sub

' Llamado desde la vista cuando cambia el cliente
Public Sub VistaClienteCambiado(ByVal strNuevoValor As String)
    If ValidarCliente(strNuevoValor) Then
        mobjModelo.ActualizarCliente strNuevoValor
    Else
        MsgBox "Cliente no válido", vbExclamation
        mobjVista.txtCliente.Value = mobjModelo.Oferta.Cliente
    End If
End Sub

Private Function ValidarCliente(ByVal strCliente As String) As Boolean
    ValidarCliente = Len(Trim(strCliente)) >= 3
End Function

Private Sub mobjModelo_DatosModificados()
    ActualizarVista
End Sub
```

```vba
'==============================================================
' UserForm: frmOferta
'--------------------------------------------------------------
' Vista del formulario - solo presentación y eventos
'==============================================================
Option Explicit

Private mobjPresentador As clsOfertaPresentador

Public Sub Init(ByVal objPresentador As clsOfertaPresentador)
    Set mobjPresentador = objPresentador
End Sub

Private Sub txtCliente_AfterUpdate()
    mobjPresentador.VistaClienteCambiado txtCliente.Value
End Sub

Private Sub btnGuardar_Click()
    mobjPresentador.Guardar
End Sub

Private Sub btnCancelar_Click()
    mobjPresentador.Cancelar
    Unload Me
End Sub
```

Instanciación:
```vba
Public Sub MostrarFormularioOferta(ByVal lngIdOferta As Long)
    Dim objModelo As New clsOfertaModelo
    Dim objPresentador As New clsOfertaPresentador
    Dim objVista As New frmOferta
    
    ' Cargar oferta en modelo
    Set objModelo.Oferta = gRepositorio.ObtenerPorId(lngIdOferta)
    
    ' Conectar componentes
    objPresentador.Init objVista, objModelo
    objVista.Init objPresentador
    
    objVista.Show vbModal
End Sub
```
