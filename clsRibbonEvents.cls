VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRibbonEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ==========================================
' CLASE DE EVENTOS DEL RIBBON
' ==========================================
' Esta clase envuelve el objeto IRibbonUI y gestiona
' su ciclo de vida con proteccion y logging.
'
' CAMBIOS IMPORTANTES:
' - ribbonUI ahora es privado con Property Get
' - mWasEverInitialized es informativo, no bloquea
' ==========================================

'@Folder "2-Servicios.Excel.Ribbon"
Option Explicit

'=============================================================================
' IMPLEMENTACION IService
'=============================================================================
Implements IService

Private mIsInitialized As Boolean

'=============================================================================
' TU CODIGO ORIGINAL
'=============================================================================

Private Const MODULE_NAME As String = "clsRibbonEvents"


'-----------------------------------------------------------------------------
' IService Implementation
'-----------------------------------------------------------------------------

Private Sub IService_Initialize(ByVal dependencies As Object)
    If mIsInitialized Then Exit Sub
    
    LogInfo MODULE_NAME, "[IService_Initialize]"
    mWasEverInitialized = False
    mIsRecovering = False
    mIsInitialized = True
End Sub

Private Sub IService_Dispose()
    LogInfo MODULE_NAME, "[IService_Dispose]"
    Call StopEvents
    mIsInitialized = False
End Sub

Private Property Get IService_IsInitialized() As Boolean
    IService_IsInitialized = mIsInitialized
End Property

Private Property Get IService_ServiceName() As String
    IService_ServiceName = "RibbonEvents"
End Property

' ==========================================
' INICIALIZACION Y LIMPIEZA
' ==========================================

Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
    mWasEverInitialized = False
    mIsRecovering = False
End Sub

Private Sub Class_Terminate()
    Call IService_Dispose
    LogInfo MODULE_NAME, "[Class_Terminate]"
End Sub

Public Sub OnGenerarGraficosDesdeCurvasRto()
    RaiseEvent GenerarGraficosDesdeCurvasRto
End Sub

Public Sub OnInvertirEjes()
    RaiseEvent InvertirEjes
End Sub

Public Sub OnFormatearCGASING()
    RaiseEvent FormatearCGASING
End Sub

Public Sub OnConfigurador()
    RaiseEvent Configurador
End Sub

Public Sub OnNuevaOportunidad()
    RaiseEvent NuevaOportunidad
End Sub

Public Sub OnReplaceWithNamesInValidations()
    RaiseEvent ReplaceWithNamesInValidations
End Sub

Public Function GetRibbonControlEnabled(control As IRibbonControl) As Boolean
    Dim enabled As Boolean
    Select Case control.id
    Case "btnGenerarGraficos"                    ' GetGraficoEnabled
        enabled = EsFicheroOportunidad() And EsValidoGenerarGrafico()
    Case "btnInvertirSeries"                     ' GetInvertirEjesEnabled
        enabled = EsFicheroOportunidad()
        'If enabled Then enabled = App.bChartActive ' comprueba que hay un grafico seleccionado
        If enabled Then enabled = EsValidoInvertirEjes()
        ' HAY QUE PASAR A ESTO: el orquestador determina el estado del ribbon:
        ' ... PERO DE MOMENTO, NO ESTA TERMINADO DE IMPLEMENTAR clsExcelFile y clsFileManager
        ' enabled = App.bCanInvertAxes
    Case "btnCGASING"                            ' GetCGASINGEnabled
        enabled = EsFicheroOportunidad()
        If enabled Then enabled = IsDefaultCGasIngSheet()
    Case "btnNuevaOp"                            ' GetNuevaOportunidadEnabled
        enabled = True
    End Select
    GetRibbonControlEnabled = enabled
    'InvalidarControl control.id ' esto ESTA DE SOBRA!!
    LogDebug MODULE_NAME, "[GetRibbonControlEnabled] - ¿" & control.id & "?: " & CBool(enabled)
End Function


' ==========================================
' gestion del ribbon
' ==========================================
'@Description: Conmuta cíclicamente el modo del Ribbon (admin ? opp ? user ? admin)
'@Scope: público
'@Category: UI / Ribbon
Public Sub ToggleRibbonMode()
    mRibbonState.ToggleModo
    Call ActivarTabSiProcede
End Sub

Private Sub ActivarTabSiProcede()
    On Error Resume Next

    If evRibbon Is Nothing Then Exit Sub

    If EsFicheroOportunidad() Then
        evRibbon.ActivarTab "tabOfertasEspecial"
    End If
End Sub

