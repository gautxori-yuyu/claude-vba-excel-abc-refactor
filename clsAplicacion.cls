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
' INICIALIZACIÓN
'=============================================================================

Private Sub Class_Initialize()
    ' NO hacer nada pesado aquí
    ' Esperar a que ThisWorkbook llame Initialize()
End Sub

Public Sub Initialize()
    If mIsInitialized Then Exit Sub
    
    LogInfo MODULE_NAME, "Iniciando aplicación..."
    
    ' PASO 1: Crear infraestructura
    Set mServiceManager = New clsServiceManager
    Set mAppContext = New clsApplicationContext
    Set mEventCoordinator = New clsEventCoordinator
    
    ' PASO 2: Registrar servicios
    Call RegisterServices
    
    ' PASO 3: Inicializar EventCoordinator (con acceso a servicios)
    mEventCoordinator.Initialize mServiceManager, mAppContext
    
    mIsInitialized = True
    
    LogInfo MODULE_NAME, "Aplicación inicializada - " & _
            mServiceManager.GetServiceCount & " servicios registrados"
End Sub

'=============================================================================
' REGISTRO DE SERVICIOS
'=============================================================================

Private Sub RegisterServices()
    LogInfo MODULE_NAME, "Registrando servicios..."
    
    '----------------------------------------------------------------------------
    ' SINGLETON: Viven toda la aplicación
    '----------------------------------------------------------------------------
    
    ' ApplicationContext (estado compartido)
    mServiceManager.RegisterSingleton mAppContext
    
    ' servicios existentes (Configuration, FileManager, etc.)
    Set mConfiguration = New clsConfiguration
    mServiceManager.RegisterSingleton mConfiguration

    Set ctx = New clsExecutionContext
    mServiceManager.RegisterSingleton ctx

    Set mFileManager = New clsFileManager
    mServiceManager.RegisterSingleton mFileManager

    Set mOpportunitiesMgr = New clsOpportunitiesMgr
    mServiceManager.RegisterSingleton mOpportunitiesMgr

    Set mChartsMgr = New clsChartEventsManager
    mServiceManager.RegisterSingleton mChartsMgr

    Set fsMon = New clsFSMonitoringCoord
    mServiceManager.RegisterSingleton fsMon

    Set mRibbonUI = New clsRibbonUI
    mServiceManager.RegisterSingleton mRibbonUI

    Set mRibbonState = New clsRibbonState
    mServiceManager.RegisterSingleton mRibbonState
    
    '----------------------------------------------------------------------------
    ' TRANSIENT: Se crean bajo demanda
    '----------------------------------------------------------------------------
    
    ' TODO: ChartGenerator, CGASINGFormatter, etc.
    ' Se añadirán cuando creemos esas clases
    '
    ' Ejemplo:
    ' mServiceManager.RegisterTransient "IChartGenerator", "CreateChartGenerator"
    
    LogInfo MODULE_NAME, "Servicios registrados: " & mServiceManager.GetServiceCount
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
' API PÚBLICA
'=============================================================================

'@Description: Acceso al gestor de servicios
Public Property Get Services() As clsServiceManager
    Set Services = mServiceManager
End Property

'@Description: Acceso al contexto de aplicación
Public Property Get Context() As clsApplicationContext
    Set Context = mAppContext
End Property

'@Description: Indica si la aplicación está inicializada
Public Property Get IsInitialized() As Boolean
    IsInitialized = mIsInitialized
End Property

'=============================================================================
' TERMINACIÓN
'=============================================================================

Public Sub Terminate()
    If Not mIsInitialized Then Exit Sub
    
    LogInfo MODULE_NAME, "Terminando aplicación..."
    
    ' Liberar en orden inverso
    On Error Resume Next
    
    If Not mEventCoordinator Is Nothing Then
        mEventCoordinator.Dispose
        Set mEventCoordinator = Nothing
    End If
    
    If Not mServiceManager Is Nothing Then
        mServiceManager.DisposeAll
        Set mServiceManager = Nothing
    End If
    
    Set mAppContext = Nothing
    
    mIsInitialized = False
    
    LogInfo MODULE_NAME, "Aplicación terminada"
End Sub

Private Sub Class_Terminate()
    Call Terminate
End Sub
