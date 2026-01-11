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
' clsEventCoordinator - Mediator centralizado de eventos
'=============================================================================
' Responsabilidades:
'   1. Suscribirse a eventos de servicios (WithEvents)
'   2. Coordinar reacciones entre servicios
'   3. Actualizar ApplicationContext cuando proceda
'   4. Ser el punto unico de coordinacion (Mediator Pattern)
'=============================================================================
' ARQUITECTURA:
'   [Servicio] --evento--> [EventCoordinator] --coordina--> [Otros servicios/Estado]
'=============================================================================

Private Const MODULE_NAME As String = "clsEventCoordinator"

'=============================================================================
' REFERENCIAS A INFRAESTRUCTURA
'=============================================================================
Private mServiceManager As clsServiceManager
Private mAppContext As clsApplicationContext

'=============================================================================
' SUSCRIPCIONES A EVENTOS (WithEvents)
'=============================================================================
' Servicios que emiten eventos
Private WithEvents mExecutionContext As clsExecutionContext
Attribute mExecutionContext.VB_VarHelpID = -1
Private WithEvents mOpportunities As clsOpportunitiesMgr
Attribute mOpportunities.VB_VarHelpID = -1
Private WithEvents mChartManager As clsChartEventsManager
Attribute mChartManager.VB_VarHelpID = -1
Private WithEvents mFSMonitoringCoord As clsFSMonitoringCoord
Attribute mFSMonitoringCoord.VB_VarHelpID = -1

' Estado que emite eventos (NO es un servicio)
Private WithEvents mRibbonState As clsRibbonState
Attribute mRibbonState.VB_VarHelpID = -1

' Ribbon callbacks (eventos de usuario)
Private WithEvents mRibbonEvents As clsRibbonEvents
Attribute mRibbonEvents.VB_VarHelpID = -1

'=============================================================================
' INICIALIZACION
'=============================================================================

Public Sub Initialize(ByVal serviceManager As clsServiceManager, _
                      ByVal appContext As clsApplicationContext)

    Set mServiceManager = serviceManager
    Set mAppContext = appContext

    LogInfo MODULE_NAME, "Suscribiendo a eventos de servicios..."

    ' Suscribirse a servicios usando acceso tipado (sin strings)
    Set mExecutionContext = mServiceManager.ExecutionContext
    Set mOpportunities = mServiceManager.OpportunitiesMgr
    Set mChartManager = mServiceManager.ChartEventsManager
    Set mFSMonitoringCoord = mServiceManager.FSMonitoringCoord

    ' Suscribirse al estado de Ribbon (desde ApplicationContext, NO es un servicio)
    Set mRibbonState = mAppContext.RibbonState

    ' RibbonEvents: Crear instancia local si no existe como servicio
    ' (gestiona callbacks del XML del ribbon)
    Set mRibbonEvents = New clsRibbonEvents

    LogInfo MODULE_NAME, "EventCoordinator inicializado - suscrito a " & _
            "ExecutionContext, OpportunitiesMgr, ChartEventsManager, FSMonitoringCoord, RibbonState, RibbonEvents"
End Sub

'=============================================================================
' MANEJADORES DE EVENTOS - ExecutionContext
'=============================================================================

Private Sub mExecutionContext_WorkbookActivated(ByVal wb As Workbook)
    LogDebug MODULE_NAME, "[WorkbookActivated] " & wb.Name

    ' Actualizar FileManager para tracking
    mServiceManager.FileManager.GetOrTrackWorkbook wb

    ' Invalidar ribbon para actualizar estado
    mServiceManager.RibbonUI.InvalidarRibbon
End Sub

Private Sub mExecutionContext_WorkbookOpened(ByVal wb As Workbook)
    LogDebug MODULE_NAME, "[WorkbookOpened] " & wb.Name

    ' Tracking automatico
    mServiceManager.FileManager.GetOrTrackWorkbook wb
End Sub

Private Sub mExecutionContext_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
    LogDebug MODULE_NAME, "[WorkbookBeforeClose] " & wb.Name

    ' Limpiar tracking
    mServiceManager.FileManager.UntrackWorkbook wb
End Sub

Private Sub mExecutionContext_SheetActivated(ByVal sh As Object)
    LogDebug MODULE_NAME, "[SheetActivated] " & sh.Name

    ' Actualizar observacion de graficos
    mServiceManager.ChartEventsManager.WatchSheet sh

    ' Invalidar ribbon
    mServiceManager.RibbonUI.InvalidarRibbon
End Sub

Private Sub mExecutionContext_SheetDeactivated(ByVal sh As Object)
    LogDebug MODULE_NAME, "[SheetDeactivated] " & sh.Name

    ' Detener observacion de graficos de la hoja anterior
    mServiceManager.ChartEventsManager.StopWatching
End Sub

'=============================================================================
' MANEJADORES DE EVENTOS - ChartEventsManager
'=============================================================================

Private Sub mChartManager_ChartActivated(Chart As Chart)
    LogInfo MODULE_NAME, "[ChartActivated] Grafico activado"

    ' Actualizar estado en ApplicationContext
    mAppContext.ChartState.IsActive = True
    Set mAppContext.ChartState.CurrentChart = Chart

    ' Verificar si se puede invertir ejes
    mAppContext.ChartState.CanInvertAxes = EsValidoInvertirEjes(Chart)

    ' Invalidar controles de grafico en ribbon
    mServiceManager.RibbonUI.InvalidarRibbon
End Sub

Private Sub mChartManager_ChartDeactivated(Chart As Chart)
    LogInfo MODULE_NAME, "[ChartDeactivated] Grafico desactivado"

    ' Actualizar estado
    mAppContext.ChartState.IsActive = False
    Set mAppContext.ChartState.CurrentChart = Nothing
    mAppContext.ChartState.CanInvertAxes = False

    ' Invalidar ribbon
    mServiceManager.RibbonUI.InvalidarRibbon
End Sub

Private Sub mChartManager_HojaConGraficosCambiada(sePuedeInvertir As Boolean)
    LogDebug MODULE_NAME, "[HojaConGraficosCambiada] sePuedeInvertir=" & sePuedeInvertir

    ' Actualizar estado de inversion
    mAppContext.ChartState.CanInvertAxes = sePuedeInvertir

    ' Invalidar ribbon
    mServiceManager.RibbonUI.InvalidarRibbon
End Sub

'=============================================================================
' MANEJADORES DE EVENTOS - OpportunitiesMgr
'=============================================================================

Private Sub mOpportunities_currOpportunityChanged(ByVal Index As Long, ByVal Path As String)
    LogInfo MODULE_NAME, "[currOpportunityChanged] Index=" & Index & ", Path=" & Path

    ' Invalidar dropdown de oportunidades en ribbon
    mServiceManager.RibbonUI.InvalidarControl "ddlOportunidades"
End Sub

Private Sub mOpportunities_OpportunityCollectionUpdate(ByVal cambios As String)
    LogInfo MODULE_NAME, "[OpportunityCollectionUpdate] Cambios: " & cambios

    ' Invalidar dropdown de oportunidades
    mServiceManager.RibbonUI.InvalidarControl "ddlOportunidades"
End Sub

'=============================================================================
' MANEJADORES DE EVENTOS - FSMonitoringCoord
'=============================================================================

Private Sub mFSMonitoringCoord_OpportunityCreated(ByVal parentFolder As String, ByVal subfolderName As String)
    LogInfo MODULE_NAME, "[OpportunityCreated] " & subfolderName

    ' Actualizar coleccion de oportunidades
    mServiceManager.OpportunitiesMgr.ProcesarCambiosEnOportunidades subfolderName
End Sub

Private Sub mFSMonitoringCoord_OpportunityDeleted(ByVal parentFolder As String, ByVal subfolderName As String)
    LogInfo MODULE_NAME, "[OpportunityDeleted] " & subfolderName

    ' Actualizar coleccion
    mServiceManager.OpportunitiesMgr.actualizarColeccionOportunidades
    mServiceManager.RibbonUI.InvalidarControl "ddlOportunidades"
End Sub

Private Sub mFSMonitoringCoord_OpportunityRenamed(ByVal parentFolder As String, ByVal oldName As String, ByVal newName As String)
    LogInfo MODULE_NAME, "[OpportunityRenamed] " & oldName & " -> " & newName

    ' Actualizar coleccion
    mServiceManager.OpportunitiesMgr.actualizarColeccionOportunidades
    mServiceManager.RibbonUI.InvalidarControl "ddlOportunidades"
End Sub

Private Sub mFSMonitoringCoord_MonitoringError(ByVal folder As String, ByVal errorMessage As String)
    LogError MODULE_NAME, "[MonitoringError] " & folder & ": " & errorMessage
End Sub

Private Sub mFSMonitoringCoord_MonitoringReconnected(ByVal folder As String, ByVal attempts As Long)
    LogInfo MODULE_NAME, "[MonitoringReconnected] " & folder & " tras " & attempts & " intentos"
End Sub

Private Sub mFSMonitoringCoord_MonitoringFailed(ByVal folder As String, ByVal reason As String)
    LogError MODULE_NAME, "[MonitoringFailed] " & folder & ": " & reason
End Sub

'=============================================================================
' MANEJADORES DE EVENTOS - RibbonState
'=============================================================================

Private Sub mRibbonState_StateChanged()
    LogDebug MODULE_NAME, "[RibbonState_StateChanged] Modo=" & mRibbonState.RibbonStateDescription

    ' Invalidar ribbon cuando cambia estado
    mServiceManager.RibbonUI.InvalidarRibbon
End Sub

'=============================================================================
' MANEJADORES DE EVENTOS - RibbonEvents (acciones de usuario)
'=============================================================================

Private Sub mRibbonEvents_GenerarGraficosDesdeCurvasRto()
    LogInfo MODULE_NAME, "[RibbonAction] GenerarGraficosDesdeCurvasRto"

    ' Llamar a la funcion existente (en modulo de graficos)
    On Error GoTo ErrHandler
    Call GenerarGraficosDesdeCurvasRto_Impl
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[GenerarGraficosDesdeCurvasRto] Error", Err.Number, Err.Description
    MsgBox "Error al generar graficos: " & Err.Description, vbCritical
End Sub

Private Sub mRibbonEvents_InvertirEjes()
    LogInfo MODULE_NAME, "[RibbonAction] InvertirEjes"

    On Error GoTo ErrHandler
    Call InvertirEjesGraficoActivo
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[InvertirEjes] Error", Err.Number, Err.Description
    MsgBox "Error al invertir ejes: " & Err.Description, vbCritical
End Sub

Private Sub mRibbonEvents_FormatearCGASING()
    LogInfo MODULE_NAME, "[RibbonAction] FormatearCGASING"

    On Error GoTo ErrHandler
    Call FormatCGASINGSheet
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[FormatearCGASING] Error", Err.Number, Err.Description
    MsgBox "Error al formatear CGASING: " & Err.Description, vbCritical
End Sub

Private Sub mRibbonEvents_MostrarConfigurador()
    LogInfo MODULE_NAME, "[RibbonAction] MostrarConfigurador"

    On Error GoTo ErrHandler
    frmConfiguracion.Show
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[MostrarConfigurador] Error", Err.Number, Err.Description
    MsgBox "Error al mostrar configurador: " & Err.Description, vbCritical
End Sub

Private Sub mRibbonEvents_NuevaOportunidad()
    LogInfo MODULE_NAME, "[RibbonAction] NuevaOportunidad"

    On Error GoTo ErrHandler
    mServiceManager.OpportunitiesMgr.CreaOportunidad
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[NuevaOportunidad] Error", Err.Number, Err.Description
    MsgBox "Error al crear oportunidad: " & Err.Description, vbCritical
End Sub

Private Sub mRibbonEvents_ToggleRibbonMode()
    LogInfo MODULE_NAME, "[RibbonAction] ToggleRibbonMode"

    ' Cambiar modo del ribbon
    mAppContext.RibbonState.ToggleModo
End Sub

'=============================================================================
' PROPIEDADES PUBLICAS
'=============================================================================

'@Description: Expone RibbonEvents para que los callbacks del Ribbon puedan disparar eventos
Public Property Get RibbonEvents() As clsRibbonEvents
    Set RibbonEvents = mRibbonEvents
End Property

'=============================================================================
' LIMPIEZA
'=============================================================================

Public Sub Dispose()
    LogInfo MODULE_NAME, "Liberando EventCoordinator..."

    ' Liberar referencias WithEvents
    Set mExecutionContext = Nothing
    Set mOpportunities = Nothing
    Set mChartManager = Nothing
    Set mFSMonitoringCoord = Nothing
    Set mRibbonState = Nothing
    Set mRibbonEvents = Nothing

    Set mAppContext = Nothing
    Set mServiceManager = Nothing

    LogInfo MODULE_NAME, "EventCoordinator liberado"
End Sub

