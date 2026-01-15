VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAplicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "1.Application"
Option Explicit

'=============================================================================
' clsAplicacion - Composition Root (punto de entrada)
'=============================================================================
' Responsabilidades:
'   1. Crear ApplicationContext (estado compartido - NO es un servicio)
'   2. Crear ServiceManager e inyectarle ApplicationContext
'   3. Registrar servicios
'   4. Crear EventCoordinator (Mediator centralizado)
'   5. Exponer acceso tipado a servicios y contexto
'   6. Gestionar ciclo de vida (Terminate)
'=============================================================================
' ARQUITECTURA:
'   - ApplicationContext: Estado compartido (RibbonState, ChartState, etc.)
'   - ServiceManager: Contenedor DI con acceso tipado (sin strings)
'   - EventCoordinator: Mediator para eventos entre servicios
'   - Servicios: Implementan IService, acceden al estado via AppContext
'=============================================================================

Private Const MODULE_NAME As String = "clsAplicacion"

'=============================================================================
' INFRAESTRUCTURA (3 componentes principales)
'=============================================================================
Private mServiceManager As clsServiceManager
Private mEventCoordinator As clsEventCoordinator
Private mAppContext As clsApplicationContext

'=============================================================================
' ESTADO (componentes que viven en ApplicationContext - NO son servicios)
'=============================================================================
Private mRibbonState As clsRibbonState
Private mChartState As clsChartState

'=============================================================================
' FLAG DE INICIALIZACION
'=============================================================================
Private mIsInitialized As Boolean

'=============================================================================
' INICIALIZACION
'=============================================================================

Private Sub Class_Initialize()
    ' NO hacer nada pesado aqui
    ' Esperar a que ThisWorkbook llame Initialize()
End Sub

Public Sub Initialize()
    If mIsInitialized Then Exit Sub

    LogInfo MODULE_NAME, "=========================================="
    LogInfo MODULE_NAME, "Iniciando aplicacion..."
    LogInfo MODULE_NAME, "=========================================="

    ' PASO 1: Crear objetos de estado (NO son servicios)
    Call CreateStateObjects

    ' PASO 2: Crear ApplicationContext y componerlo con los estados
    Call CreateApplicationContext

    ' PASO 3: Crear ServiceManager e inyectarle ApplicationContext
    Set mServiceManager = New clsServiceManager
    mServiceManager.Initialize mAppContext

    ' PASO 4: Registrar servicios (solo los que implementan IService)
    Call RegisterServices

    ' PASO 5: Crear EventCoordinator (Mediator) y conectarlo
    Set mEventCoordinator = New clsEventCoordinator
    mEventCoordinator.Initialize mServiceManager, mAppContext

    mIsInitialized = True

    LogInfo MODULE_NAME, "=========================================="
    LogInfo MODULE_NAME, "Aplicacion inicializada - " & _
            mServiceManager.GetServiceCount & " servicios registrados"
    LogInfo MODULE_NAME, "=========================================="
End Sub

'=============================================================================
' CREACION DE OBJETOS DE ESTADO
'=============================================================================

Private Sub CreateStateObjects()
    LogInfo MODULE_NAME, "Creando objetos de estado..."

    ' Estado del Ribbon (NO es un servicio)
    Set mRibbonState = New clsRibbonState

    ' Estado de graficos (NO es un servicio)
    Set mChartState = New clsChartState

    LogInfo MODULE_NAME, "Objetos de estado creados"
End Sub

'=============================================================================
' CREACION DE APPLICATION CONTEXT
'=============================================================================

Private Sub CreateApplicationContext()
    LogInfo MODULE_NAME, "Creando ApplicationContext..."

    Set mAppContext = New clsApplicationContext

    ' Nota: ExecutionContext se crea como servicio y se pasa despues
    ' Por ahora inicializamos con Nothing y se actualizara despues
    ' de registrar los servicios

    LogInfo MODULE_NAME, "ApplicationContext creado"
End Sub

'=============================================================================
' REGISTRO DE SERVICIOS
'=============================================================================

Private Sub RegisterServices()
    LogInfo MODULE_NAME, "Registrando servicios..."

    '------------------------------------------------------------------------
    ' SERVICIOS SINGLETON (implementan IService, tienen ciclo de vida)
    '------------------------------------------------------------------------

    ' Configuration - carga desde registro
    mServiceManager.RegisterSingleton New clsConfiguration

    ' ExecutionContext - suscripcion a eventos de Application
    Dim execCtx As clsExecutionContext
    Set execCtx = New clsExecutionContext
    mServiceManager.RegisterSingleton execCtx

    ' Ahora que tenemos ExecutionContext, inicializar ApplicationContext completo
    mAppContext.Initialize mRibbonState, execCtx, mChartState

    ' FileManager - tracking de archivos
    mServiceManager.RegisterSingleton New clsFileManager

    ' OpportunitiesMgr - gestion de carpetas de oportunidades
    mServiceManager.RegisterSingleton New clsOpportunitiesMgr

    ' ChartEventsManager - gestion de eventos de graficos
    mServiceManager.RegisterSingleton New clsChartEventsManager

    ' FSMonitoringCoord - monitoreo del sistema de archivos
    mServiceManager.RegisterSingleton New clsFSMonitoringCoord

    ' RibbonUI - gestion del puntero IRibbonUI
    mServiceManager.RegisterSingleton New clsRibbonUI

    LogInfo MODULE_NAME, "Servicios registrados: " & mServiceManager.GetServiceCount
End Sub

'=============================================================================
' API PUBLICA - Acceso a infraestructura
'=============================================================================

'@Description: Acceso al gestor de servicios (con propiedades tipadas)
Public Property Get Services() As clsServiceManager
    Set Services = mServiceManager
End Property

'@Description: Acceso al contexto de aplicacion (estado compartido)
Public Property Get Context() As clsApplicationContext
    Set Context = mAppContext
End Property

'@Description: Acceso al coordinador de eventos (Mediator)
Public Property Get Events() As clsEventCoordinator
    Set Events = mEventCoordinator
End Property

'@Description: Indica si la aplicacion esta inicializada
Public Property Get IsInitialized() As Boolean
    IsInitialized = mIsInitialized
End Property

'=============================================================================
' API PUBLICA - Acceso tipado a servicios (facade conveniente)
'=============================================================================
' Estos accesos delegan al ServiceManager para consistencia.
' Son atajos para evitar escribir App.Services.XXX
'=============================================================================

Public Property Get Configuration() As clsConfiguration
    Set Configuration = mServiceManager.Configuration
End Property

Public Property Get ExecutionContext() As clsExecutionContext
    Set ExecutionContext = mServiceManager.ExecutionContext
End Property

Public Property Get FileManager() As clsFileManager
    Set FileManager = mServiceManager.FileManager
End Property

Public Property Get OpportunitiesMgr() As clsOpportunitiesMgr
    Set OpportunitiesMgr = mServiceManager.OpportunitiesMgr
End Property

Public Property Get ChartEventsManager() As clsChartEventsManager
    Set ChartEventsManager = mServiceManager.ChartEventsManager
End Property

Public Property Get FSMonitoringCoord() As clsFSMonitoringCoord
    Set FSMonitoringCoord = mServiceManager.FSMonitoringCoord
End Property

Public Property Get RibbonUI() As clsRibbonUI
    Set RibbonUI = mServiceManager.RibbonUI
End Property

'=============================================================================
' API PUBLICA - Acceso a estado (desde ApplicationContext)
'=============================================================================

Public Property Get RibbonState() As clsRibbonState
    Set RibbonState = mRibbonState
End Property

Public Property Get ChartState() As clsChartState
    Set ChartState = mChartState
End Property

'=============================================================================
' TERMINACION
'=============================================================================

Public Sub Terminate()
    If Not mIsInitialized Then Exit Sub

    LogInfo MODULE_NAME, "=========================================="
    LogInfo MODULE_NAME, "Terminando aplicacion..."
    LogInfo MODULE_NAME, "=========================================="

    ' Liberar en orden inverso a la creacion
    On Error Resume Next

    ' 1. EventCoordinator primero (deja de escuchar eventos)
    If Not mEventCoordinator Is Nothing Then
        mEventCoordinator.Dispose
        Set mEventCoordinator = Nothing
    End If

    ' 2. ServiceManager (dispone todos los servicios)
    If Not mServiceManager Is Nothing Then
        mServiceManager.DisposeAll
        Set mServiceManager = Nothing
    End If

    ' 3. Estado
    Set mRibbonState = Nothing
    Set mChartState = Nothing

    ' 4. ApplicationContext
    Set mAppContext = Nothing

    mIsInitialized = False

    LogInfo MODULE_NAME, "Aplicacion terminada"
    LogInfo MODULE_NAME, "=========================================="
End Sub

Private Sub Class_Terminate()
    Call Terminate
End Sub

