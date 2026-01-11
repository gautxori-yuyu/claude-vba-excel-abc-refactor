VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRibbonEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "2-Servicios.Excel.Ribbon"
Option Explicit

'=============================================================================
' clsRibbonEvents - Eventos de acciones del Ribbon
'=============================================================================
' Responsabilidades:
'   1. Declarar eventos publicos para acciones del Ribbon
'   2. Proporcionar metodos On* que disparan los eventos
'   3. Evaluar estado enabled de controles del Ribbon
'=============================================================================
' USO:
'   - clsEventCoordinator se suscribe con WithEvents
'   - modCALLBACKSRibbon llama a los metodos On* cuando el usuario pulsa botones
'=============================================================================

Private Const MODULE_NAME As String = "clsRibbonEvents"

'=============================================================================
' EVENTOS PUBLICOS - Para suscripcion con WithEvents
'=============================================================================
Public Event GenerarGraficosDesdeCurvasRto()
Public Event InvertirEjes()
Public Event FormatearCGASING()
Public Event MostrarConfigurador()
Public Event NuevaOportunidad()
Public Event ReplaceWithNamesInValidations()
Public Event ToggleRibbonMode()

'=============================================================================
' METODOS PUBLICOS - Disparan eventos (llamados desde callbacks del Ribbon)
'=============================================================================

Public Sub OnGenerarGraficosDesdeCurvasRto()
    LogDebug MODULE_NAME, "[OnGenerarGraficosDesdeCurvasRto]"
    RaiseEvent GenerarGraficosDesdeCurvasRto
End Sub

Public Sub OnInvertirEjes()
    LogDebug MODULE_NAME, "[OnInvertirEjes]"
    RaiseEvent InvertirEjes
End Sub

Public Sub OnFormatearCGASING()
    LogDebug MODULE_NAME, "[OnFormatearCGASING]"
    RaiseEvent FormatearCGASING
End Sub

Public Sub OnMostrarConfigurador()
    LogDebug MODULE_NAME, "[OnMostrarConfigurador]"
    RaiseEvent MostrarConfigurador
End Sub

Public Sub OnNuevaOportunidad()
    LogDebug MODULE_NAME, "[OnNuevaOportunidad]"
    RaiseEvent NuevaOportunidad
End Sub

Public Sub OnReplaceWithNamesInValidations()
    LogDebug MODULE_NAME, "[OnReplaceWithNamesInValidations]"
    RaiseEvent ReplaceWithNamesInValidations
End Sub

Public Sub OnToggleRibbonMode()
    LogDebug MODULE_NAME, "[OnToggleRibbonMode]"
    RaiseEvent ToggleRibbonMode
End Sub

'=============================================================================
' EVALUACION DE ESTADO DE CONTROLES
'=============================================================================

'@Description: Determina si un control del Ribbon debe estar habilitado
Public Function GetRibbonControlEnabled(control As IRibbonControl) As Boolean
    Dim enabled As Boolean

    Select Case control.Id
    Case "btnGenerarGraficos"
        enabled = EsFicheroOportunidad() And EsValidoGenerarGrafico()

    Case "btnInvertirSeries"
        enabled = EsFicheroOportunidad()
        If enabled Then enabled = EsValidoInvertirEjes()

    Case "btnCGASING"
        enabled = EsFicheroOportunidad()
        If enabled Then enabled = IsDefaultCGasIngSheet()

    Case "btnNuevaOp"
        enabled = True

    Case Else
        enabled = True
    End Select

    GetRibbonControlEnabled = enabled
    LogDebug MODULE_NAME, "[GetRibbonControlEnabled] " & control.Id & " = " & enabled
End Function

'=============================================================================
' CICLO DE VIDA
'=============================================================================

Private Sub Class_Initialize()
    LogDebug MODULE_NAME, "[Class_Initialize]"
End Sub

Private Sub Class_Terminate()
    LogDebug MODULE_NAME, "[Class_Terminate]"
End Sub

