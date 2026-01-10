VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventCoordinator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "2.Infrastructure"
Option Explicit

'=============================================================================
' clsEventCoordinator - Coordinador centralizado de eventos
'=============================================================================
' Responsabilidades:
'   1. Suscribirse a eventos de servicios (WithEvents)
'   2. Coordinar reacciones entre servicios
'   3. Actualizar ApplicationContext cuando proceda
'   4. Propagar eventos entre componentes desacoplados
'=============================================================================

Private Const MODULE_NAME As String = "clsEventCoordinator"

'=============================================================================
' ESTADO
'=============================================================================

Private mIsInitialized As Boolean

'=============================================================================
' REFERENCIAS A INFRAESTRUCTURA
'=============================================================================

Private mServiceManager As clsServiceManager
Private mAppContext As clsApplicationContext

'=============================================================================
' SUSCRIPCIONES A EVENTOS (WithEvents)
' Cada servicio que emite eventos tiene su variable WithEvents aqui
'=============================================================================

Private WithEvents ctx As clsExecutionContext
Attribute ctx.VB_VarHelpID = -1
Private WithEvents mOpportunities As clsOpportunitiesMgr
Attribute mOpportunities.VB_VarHelpID = -1
Private WithEvents mChartManager As clsChartEventsManager
Attribute mChartManager.VB_VarHelpID = -1
Private WithEvents mFSMonitoringCoord As clsFSMonitoringCoord
Attribute mFSMonitoringCoord.VB_VarHelpID = -1
Private WithEvents mRibbonState As clsRibbonState
Attribute mRibbonState.VB_VarHelpID = -1
Private WithEvents evRibbon As clsRibbonEvents
Attribute evRibbon.VB_VarHelpID = -1


' ==========================================
' EVENTOS DE APLICACION
' ==========================================
Public Event GenerarGraficosDesdeCurvasRto()
Public Event InvertirEjes()
Public Event FormatearCGASING()
Public Event Configurador()
Public Event NuevaOportunidad()
Public Event ReplaceWithNamesInValidations()


'=============================================================================
' INICIALIZACION
'=============================================================================

Public Sub Initialize(ByVal serviceManager As clsServiceManager, _
                      ByVal appContext As clsApplicationContext)
    On Error GoTo ErrHandler

    If mIsInitialized Then
        LogWarning MODULE_NAME, "Ya inicializado, ignorando llamada"
        Exit Sub
    End If

    Set mServiceManager = serviceManager
    Set mAppContext = appContext

    LogInfo MODULE_NAME, "Suscribiendo a eventos de servicios..."

    ' Suscribirse a cada servicio que emite eventos
    ' Los nombres deben coincidir con IService_ServiceName de cada clase
    On Error Resume Next  ' Algunos servicios pueden no estar registrados

    Set mOpportunities = mServiceManager.GetService("OpportunitiesMgr")
    If mOpportunities Is Nothing Then LogDebug MODULE_NAME, "OpportunitiesMgr no disponible"

    Set mChartManager = mServiceManager.GetService("ChartEventsManager")
    If mChartManager Is Nothing Then LogDebug MODULE_NAME, "ChartEventsManager no disponible"

    Set mFSMonitoringCoord = mServiceManager.GetService("FSMonitoringCoord")
    If mFSMonitoringCoord Is Nothing Then LogDebug MODULE_NAME, "FSMonitoringCoord no disponible"

    Set mRibbonState = mServiceManager.GetService("RibbonState")
    If mRibbonState Is Nothing Then LogDebug MODULE_NAME, "RibbonState no disponible"

    Set evRibbon = mServiceManager.GetService("RibbonEvents")
    If evRibbon Is Nothing Then LogDebug MODULE_NAME, "RibbonEvents no disponible"

    Set ctx = mServiceManager.GetService("ExecutionContext")
    If ctx Is Nothing Then LogDebug MODULE_NAME, "ExecutionContext no disponible"

    On Error GoTo ErrHandler

    mIsInitialized = True
    LogInfo MODULE_NAME, "EventCoordinator inicializado - suscripciones activas"

    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "Error durante inicializacion", Err.Number, Err.Description
    Call ReleaseEventSubscriptions
    Err.Raise Err.Number, MODULE_NAME, Err.Description
End Sub

'@Description: Indica si el coordinador esta inicializado
Public Property Get IsInitialized() As Boolean
    IsInitialized = mIsInitialized
End Property

'=============================================================================
' MANEJADORES DE EVENTOS
'=============================================================================
' TODO: Migrar los manejadores de eventos desde clsAplicacion de rama main

Private Sub mRibbonState_StateChanged()
    ' Invalidar ribbon cuando cambia estado
    App.RibbonUI.InvalidarRibbon
End Sub

'=============================================================================
' LIMPIEZA
'=============================================================================

Public Sub Dispose()
    If Not mIsInitialized Then
        LogDebug MODULE_NAME, "No inicializado, nada que liberar"
        Exit Sub
    End If

    LogInfo MODULE_NAME, "Liberando EventCoordinator..."

    ' Liberar suscripciones a eventos
    Call ReleaseEventSubscriptions

    ' Liberar referencias de infraestructura
    Set mAppContext = Nothing
    Set mServiceManager = Nothing

    mIsInitialized = False

    LogInfo MODULE_NAME, "EventCoordinator liberado correctamente"
End Sub

'@Description: Libera todas las suscripciones a eventos (WithEvents)
Private Sub ReleaseEventSubscriptions()
    On Error Resume Next

    LogDebug MODULE_NAME, "Liberando suscripciones a eventos..."

    Set mOpportunities = Nothing
    Set mChartManager = Nothing
    Set mFSMonitoringCoord = Nothing
    Set mRibbonState = Nothing
    Set evRibbon = Nothing
    Set ctx = Nothing

    LogDebug MODULE_NAME, "Suscripciones liberadas"

    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Call Dispose
End Sub
