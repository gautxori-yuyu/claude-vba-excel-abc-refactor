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
' clsAplicacion - Coordinador principal
'=============================================================================
' Responsabilidades:
'   1. Crear ServiceManager
'   2. Registrar servicios
'   3. Crear EventCoordinator
'   4. Exponer ServiceManager
'   5. Limpiar al cerrar
'=============================================================================

Private Const MODULE_NAME As String = "clsAplicacion"

'=============================================================================
' DEPENDENCIAS (3)
'=============================================================================
Private mServiceManager As clsServiceManager
Private mEventCoordinator As clsEventCoordinator
Private mAppContext As clsApplicationContext

'=============================================================================
' FACADE para servicios
'=============================================================================
Private mConfiguration As clsConfiguration
Private ctx As clsExecutionContext
Private mFileManager As clsFileManager
Private mOpportunitiesMgr As clsOpportunitiesMgr
Private mChartsMgr As clsChartEventsManager
Private fsMon As clsFSMonitoringCoord
Private mRibbonUI As clsRibbonUI
Private mRibbonState As clsRibbonState

'=============================================================================
' ESTADO
'=============================================================================
Private mIsInitialized As Boolean

'=============================================================================
' INICIALIZACI�N
'=============================================================================

Private Sub Class_Initialize()
    ' NO hacer nada pesado aqu�
    ' Esperar a que ThisWorkbook llame Initialize()
End Sub

Public Sub Initialize()
    On Error GoTo ErrHandler

    If mIsInitialized Then
        LogWarning MODULE_NAME, "Aplicacion ya inicializada, ignorando llamada"
        Exit Sub
    End If

    LogInfo MODULE_NAME, "=== INICIANDO APLICACION ==="

    ' PASO 1: Crear infraestructura base
    LogDebug MODULE_NAME, "Paso 1/3: Creando infraestructura..."
    Set mServiceManager = New clsServiceManager
    Set mAppContext = New clsApplicationContext
    Set mEventCoordinator = New clsEventCoordinator

    ' PASO 2: Registrar todos los servicios
    LogDebug MODULE_NAME, "Paso 2/3: Registrando servicios..."
    Call RegisterServices

    ' PASO 3: Inicializar coordinador de eventos
    LogDebug MODULE_NAME, "Paso 3/3: Inicializando EventCoordinator..."
    mEventCoordinator.Initialize mServiceManager, mAppContext

    mIsInitialized = True

    LogInfo MODULE_NAME, "=== APLICACION INICIALIZADA ==="
    LogInfo MODULE_NAME, "Servicios: " & mServiceManager.GetRegisteredServiceNames

    Exit Sub

ErrHandler:
    LogCritical MODULE_NAME, "Error fatal durante inicializacion", Err.Number, Err.Description
    Call CleanupOnError
    Err.Raise Err.Number, MODULE_NAME, "Fallo en Initialize: " & Err.Description
End Sub

'@Description: Limpieza en caso de error durante inicializacion
Private Sub CleanupOnError()
    On Error Resume Next

    Set mEventCoordinator = Nothing
    Set mAppContext = Nothing

    If Not mServiceManager Is Nothing Then
        mServiceManager.DisposeAll
        Set mServiceManager = Nothing
    End If

    mIsInitialized = False

    LogWarning MODULE_NAME, "Limpieza de emergencia completada"
End Sub

'=============================================================================
' REGISTRO DE SERVICIOS
'=============================================================================

Private Sub RegisterServices()
    On Error GoTo ErrHandler

    LogInfo MODULE_NAME, "Registrando servicios Singleton..."

    '----------------------------------------------------------------------------
    ' SINGLETON: Servicios que viven durante toda la aplicacion
    '----------------------------------------------------------------------------

    ' 1. ApplicationContext (estado global compartido)
    mServiceManager.RegisterSingleton mAppContext

    ' 2. Configuration (rutas y configuracion persistente)
    Set mConfiguration = New clsConfiguration
    mServiceManager.RegisterSingleton mConfiguration

    ' 3. ExecutionContext (contexto de ejecucion actual)
    Set ctx = New clsExecutionContext
    mServiceManager.RegisterSingleton ctx

    ' 4. FileManager (gestion de archivos Excel)
    Set mFileManager = New clsFileManager
    mServiceManager.RegisterSingleton mFileManager

    ' 5. OpportunitiesMgr (gestion de oportunidades)
    Set mOpportunitiesMgr = New clsOpportunitiesMgr
    mServiceManager.RegisterSingleton mOpportunitiesMgr

    ' 6. ChartEventsManager (eventos de graficos)
    Set mChartsMgr = New clsChartEventsManager
    mServiceManager.RegisterSingleton mChartsMgr

    ' 7. FSMonitoringCoord (monitoreo de carpetas)
    Set fsMon = New clsFSMonitoringCoord
    mServiceManager.RegisterSingleton fsMon

    ' 8. RibbonUI (interfaz del ribbon)
    Set mRibbonUI = New clsRibbonUI
    mServiceManager.RegisterSingleton mRibbonUI

    ' 9. RibbonState (estado del ribbon)
    Set mRibbonState = New clsRibbonState
    mServiceManager.RegisterSingleton mRibbonState

    '----------------------------------------------------------------------------
    ' TRANSIENT: Servicios creados bajo demanda
    ' (Descomentar cuando se implementen las clases)
    '----------------------------------------------------------------------------
    ' mServiceManager.RegisterTransient "IChartGenerator", "CreateChartGenerator"
    ' mServiceManager.RegisterTransient "ICGASINGFormatter", "CreateCGASINGFormatter"

    LogInfo MODULE_NAME, "Servicios registrados: " & mServiceManager.GetServiceCount

    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "Error registrando servicios", Err.Number, Err.Description
    Err.Raise Err.Number, MODULE_NAME, "Fallo en RegisterServices: " & Err.Description
End Sub

'=============================================================================
' PATRON FACADE PARA SERVICIOS REGISTRADOS
'=============================================================================
 Public Property Get Configuration() As clsConfiguration
     Set Configuration = mConfiguration
 End Property

 Public Property Get ExecutionContext() As clsExecutionContext
     Set ExecutionContext = ctx
 End Property

 Public Property Get FileManager() As clsFileManager
     Set FileManager = mFileManager
 End Property

 Public Property Get OpportunitiesMgr() As clsOpportunitiesMgr
     Set OpportunitiesMgr = mOpportunitiesMgr
 End Property

 Public Property Get ChartEventsManager() As clsChartEventsManager
     Set ChartEventsManager = mChartsMgr
 End Property

 Public Property Get FSMonitoringCoord() As clsFSMonitoringCoord
     Set FSMonitoringCoord = fsMon
 End Property

 Public Property Get RibbonUI() As clsRibbonUI
     Set RibbonUI = mRibbonUI
 End Property

 Public Property Get RibbonState() As clsRibbonState
     Set RibbonState = mRibbonState
 End Property

'=============================================================================
' API P�BLICA
'=============================================================================

'@Description: Acceso al gestor de servicios
Public Property Get Services() As clsServiceManager
    Set Services = mServiceManager
End Property

'@Description: Acceso al contexto de aplicaci�n
Public Property Get Context() As clsApplicationContext
    Set Context = mAppContext
End Property

'@Description: Indica si la aplicaci�n est� inicializada
Public Property Get IsInitialized() As Boolean
    IsInitialized = mIsInitialized
End Property

'=============================================================================
' TERMINACION
'=============================================================================

Public Sub Terminate()
    If Not mIsInitialized Then
        LogDebug MODULE_NAME, "Aplicacion no inicializada, nada que terminar"
        Exit Sub
    End If

    LogInfo MODULE_NAME, "=== TERMINANDO APLICACION ==="

    On Error Resume Next

    ' 1. Liberar coordinador de eventos primero
    If Not mEventCoordinator Is Nothing Then
        LogDebug MODULE_NAME, "Liberando EventCoordinator..."
        mEventCoordinator.Dispose
        Set mEventCoordinator = Nothing
    End If

    ' 2. Liberar referencias locales del facade
    LogDebug MODULE_NAME, "Liberando referencias del facade..."
    Set mConfiguration = Nothing
    Set ctx = Nothing
    Set mFileManager = Nothing
    Set mOpportunitiesMgr = Nothing
    Set mChartsMgr = Nothing
    Set fsMon = Nothing
    Set mRibbonUI = Nothing
    Set mRibbonState = Nothing

    ' 3. Liberar todos los servicios (en orden inverso)
    If Not mServiceManager Is Nothing Then
        LogDebug MODULE_NAME, "Liberando ServiceManager..."
        mServiceManager.DisposeAll
        Set mServiceManager = Nothing
    End If

    ' 4. Liberar contexto
    Set mAppContext = Nothing

    mIsInitialized = False

    LogInfo MODULE_NAME, "=== APLICACION TERMINADA ==="

    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Call Terminate
End Sub
