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
'=============================================================================

Private Const MODULE_NAME As String = "clsEventCoordinator"

'=============================================================================
' REFERENCIAS
'=============================================================================

Private mServiceManager As clsServiceManager
Private mAppContext As clsApplicationContext

'=============================================================================
' SUSCRIPCIONES A EVENTOS (WithEvents)
'=============================================================================
' TODO: Aquí irán TODOS los WithEvents de tus servicios
'
' Ejemplo de lo que moveremos desde clsAplicacion:
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
' INICIALIZACIÓN
'=============================================================================

Public Sub Initialize(ByVal serviceManager As clsServiceManager, _
                      ByVal appContext As clsApplicationContext)
    
    Set mServiceManager = serviceManager
    Set mAppContext = appContext
    
    LogInfo MODULE_NAME, "Suscribiendo a eventos de servicios..."
    
    ' TODO: Obtener servicios y asignar a variables WithEvents
    '
    ' Ejemplo:
     Set mOpportunities = mServiceManager.GetService("IOpportunities")
     Set mChartManager = mServiceManager.GetService("IChartManager")
     Set mFSMonitoringCoord = mServiceManager.GetService("IFSMonitoring")
     Set mRibbonState = mServiceManager.GetService("IRibbonState")
     Set evRibbon = mServiceManager.GetService("IRibbonEvents")
     Set ctx = mServiceManager.GetService("IExecutionContext")
    
    LogInfo MODULE_NAME, "EventCoordinator inicializado"
End Sub

'=============================================================================
' MANEJADORES DE EVENTOS
'=============================================================================
' TODO: Aquí irán TODOS los manejadores de eventos (los 22 que tienes)
'
' ESTRUCTURA TÍPICA de cada manejador:
'
' Private Sub mOpportunities_currOpportunityChanged(ByVal Index As Long, ByVal Path As String)
'     ' 1. Actualizar contexto si procede
'     Dim opp As Object
'     Set opp = mOpportunities.GetByIndex(Index)
'     Set mAppContext.CurrentOpportunity = opp
'
'     ' 2. Coordinar con otros servicios
'     Dim ribbonSvc As Object
'     Set ribbonSvc = mServiceManager.GetService("IRibbon")
'     ribbonSvc.InvalidateControl "ddlOportunidades"
'
'     ' 3. Logging
'     LogInfo MODULE_NAME, "Oportunidad cambiada: " & Path
' End Sub
'
'
' Private Sub mChartManager_ChartActivated(Chart As Chart)
'     ' 1. Actualizar contexto
'     mAppContext.IsChartActive = True
'
'     ' 2. Coordinar con ribbon
'     Dim ribbonSvc As Object
'     Set ribbonSvc = mServiceManager.GetService("IRibbon")
'     ribbonSvc.EnableChartControls
'
'     ' 3. Logging
'     LogInfo MODULE_NAME, "Gráfico activado"
' End Sub
'
'
' Private Sub mChartManager_ChartDeactivated(Chart As Chart)
'     mAppContext.IsChartActive = False
'
'     Dim ribbonSvc As Object
'     Set ribbonSvc = mServiceManager.GetService("IRibbon")
'     ribbonSvc.DisableChartControls
'
'     LogInfo MODULE_NAME, "Gráfico desactivado"
' End Sub
'
'
' Private Sub evRibbon_GenerarGraficosDesdeCurvasRto()
'     ' Coordinar: Llamar al servicio de gráficos
'     Dim chartGen As Object
'     Set chartGen = mServiceManager.GetService("IChartGenerator")
'     chartGen.GenerarGraficoSensibilidad
' End Sub
'

Private Sub mRibbonState_StateChanged()
    ' Invalidar ribbon cuando cambia estado
    App.RibbonUI.InvalidarRibbon
End Sub

'
' Private Sub evRibbon_FormatearCGASING()
'     ' Coordinar: Llamar al formateador
'     Dim formatter As IFormatter
'     Set formatter = mServiceManager.GetService("ICGASINGFormatter")
'
'     If formatter.CanFormat(ActiveWorkbook) Then
'         formatter.Format ActiveWorkbook
'     End If
' End Sub
'
'
' Private Sub mFSMonitoringCoord_OpportunityCreated(ByVal parentFolder As String, _
'                                                     ByVal subfolderName As String)
'     ' Coordinar: Actualizar lista de oportunidades
'     Dim oppSvc As Object
'     Set oppSvc = mServiceManager.GetService("IOpportunities")
'     oppSvc.RefreshCollection
'
'     ' Actualizar ribbon
'     Dim ribbonSvc As Object
'     Set ribbonSvc = mServiceManager.GetService("IRibbon")
'     ribbonSvc.InvalidateControl "ddlOportunidades"
'
'     LogInfo MODULE_NAME, "Oportunidad creada: " & subfolderName
' End Sub
'
' ... Y así con los 22 manejadores de eventos

'=============================================================================
' LIMPIEZA
'=============================================================================

Public Sub Dispose()
    LogInfo MODULE_NAME, "Liberando EventCoordinator..."
    
    ' Liberar referencias WithEvents
    ' TODO: Cuando añadas los WithEvents, libéralos aquí
    '
    ' Ejemplo:
    ' Set mOpportunities = Nothing
    ' Set mChartManager = Nothing
    ' Set mFSMonitoringCoord = Nothing
    ' Set mRibbonState = Nothing
    ' Set evRibbon = Nothing
    ' Set ctx = Nothing
    
    Set mAppContext = Nothing
    Set mServiceManager = Nothing
    
    LogInfo MODULE_NAME, "EventCoordinator liberado"
End Sub
