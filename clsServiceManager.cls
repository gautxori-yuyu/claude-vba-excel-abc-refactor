VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsServiceManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "2.Infrastructure"
Option Explicit

'=============================================================================
' clsServiceManager - Gestor de servicios con resolución por tipo
'=============================================================================
' Responsabilidades:
'   1. Registrar servicios Singleton (usando TypeName como clave)
'   2. Resolver servicios mediante propiedades tipadas
'   3. Gestionar ciclo de vida (Initialize/Dispose vía IService)
'   4. Inyectar dependencias (ApplicationContext + ServiceManager)
'=============================================================================
' DISEÑO:
'   - NO usa nombres de servicio como strings
'   - Usa TypeName() internamente como clave del diccionario
'   - Expone propiedades tipadas para acceso a servicios
'   - Todos los servicios son Singleton
'=============================================================================

Private Const MODULE_NAME As String = "clsServiceManager"

'=============================================================================
' ESTADO
'=============================================================================
Private mServices As Object                         ' Dictionary: TypeName -> IService
Private mAppContext As clsApplicationContext        ' Contexto compartido (inyectado)
Private mIsInitialized As Boolean

'=============================================================================
' INICIALIZACIÓN
'=============================================================================

Private Sub Class_Initialize()
    Set mServices = CreateObject("Scripting.Dictionary")
    mServices.CompareMode = vbTextCompare
End Sub

Private Sub Class_Terminate()
    Call DisposeAll
End Sub

'@Description: Inicializa el ServiceManager con el ApplicationContext
Public Sub Initialize(ByVal appContext As clsApplicationContext)
    If mIsInitialized Then Exit Sub

    Set mAppContext = appContext
    mIsInitialized = True

    LogInfo MODULE_NAME, "[Initialize] ServiceManager inicializado con ApplicationContext"
End Sub

'=============================================================================
' REGISTRO DE SERVICIOS
'=============================================================================

'@Description: Registra un servicio Singleton (clave = TypeName automático)
Public Sub RegisterSingleton(ByVal instance As IService)
    On Error GoTo ErrHandler

    Dim key As String
    key = TypeName(instance)

    If mServices.Exists(key) Then
        ' Reemplazar servicio existente
        Set mServices(key) = instance
        LogInfo MODULE_NAME, "Singleton actualizado: " & key
    Else
        ' Añadir nuevo servicio
        mServices.Add key, instance
        LogInfo MODULE_NAME, "Singleton registrado: " & key
    End If

    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[RegisterSingleton] Error", Err.Number, Err.Description
    Err.Raise Err.Number, MODULE_NAME, "Error registrando servicio: " & Err.Description
End Sub

'=============================================================================
' INICIALIZACIÓN DE SERVICIOS (Lazy)
'=============================================================================

'@Description: Inicializa un servicio si aún no está inicializado
Private Sub EnsureInitialized(ByVal service As IService)
    If service.IsInitialized Then Exit Sub

    ' Inyectar dependencias: ServiceManager (Me) para acceso a otros servicios
    service.Initialize Me

    LogInfo MODULE_NAME, "Servicio inicializado: " & TypeName(service)
End Sub

'@Description: Inicializa todos los servicios registrados
Public Sub InitializeAll()
    Dim key As Variant
    Dim service As IService

    LogInfo MODULE_NAME, "Inicializando " & mServices.Count & " servicios..."

    For Each key In mServices.Keys
        Set service = mServices(key)
        Call EnsureInitialized(service)
    Next key

    LogInfo MODULE_NAME, "Todos los servicios inicializados"
End Sub

'=============================================================================
' PROPIEDADES TIPADAS - Acceso a servicios sin strings
'=============================================================================
' Cada propiedad devuelve el servicio específico con su tipo correcto.
' Lazy initialization: el servicio se inicializa la primera vez que se accede.
'=============================================================================

Public Property Get Configuration() As clsConfiguration
    Const KEY As String = "clsConfiguration"
    If Not mServices.Exists(KEY) Then Exit Property

    Dim service As IService
    Set service = mServices(KEY)
    Call EnsureInitialized(service)

    Set Configuration = mServices(KEY)
End Property

Public Property Get ExecutionContext() As clsExecutionContext
    Const KEY As String = "clsExecutionContext"
    If Not mServices.Exists(KEY) Then Exit Property

    Dim service As IService
    Set service = mServices(KEY)
    Call EnsureInitialized(service)

    Set ExecutionContext = mServices(KEY)
End Property

Public Property Get FileManager() As clsFileManager
    Const KEY As String = "clsFileManager"
    If Not mServices.Exists(KEY) Then Exit Property

    Dim service As IService
    Set service = mServices(KEY)
    Call EnsureInitialized(service)

    Set FileManager = mServices(KEY)
End Property

Public Property Get OpportunitiesMgr() As clsOpportunitiesMgr
    Const KEY As String = "clsOpportunitiesMgr"
    If Not mServices.Exists(KEY) Then Exit Property

    Dim service As IService
    Set service = mServices(KEY)
    Call EnsureInitialized(service)

    Set OpportunitiesMgr = mServices(KEY)
End Property

Public Property Get ChartEventsManager() As clsChartEventsManager
    Const KEY As String = "clsChartEventsManager"
    If Not mServices.Exists(KEY) Then Exit Property

    Dim service As IService
    Set service = mServices(KEY)
    Call EnsureInitialized(service)

    Set ChartEventsManager = mServices(KEY)
End Property

Public Property Get FSMonitoringCoord() As clsFSMonitoringCoord
    Const KEY As String = "clsFSMonitoringCoord"
    If Not mServices.Exists(KEY) Then Exit Property

    Dim service As IService
    Set service = mServices(KEY)
    Call EnsureInitialized(service)

    Set FSMonitoringCoord = mServices(KEY)
End Property

Public Property Get RibbonUI() As clsRibbonUI
    Const KEY As String = "clsRibbonUI"
    If Not mServices.Exists(KEY) Then Exit Property

    Dim service As IService
    Set service = mServices(KEY)
    Call EnsureInitialized(service)

    Set RibbonUI = mServices(KEY)
End Property

'=============================================================================
' PROPIEDADES DE CONSULTA
'=============================================================================

'@Description: Devuelve el número de servicios registrados
Public Function GetServiceCount() As Long
    GetServiceCount = mServices.Count
End Function

'@Description: Verifica si un servicio está registrado (por TypeName)
Public Function IsRegistered(ByVal typeName As String) As Boolean
    IsRegistered = mServices.Exists(typeName)
End Function

'@Description: Acceso al ApplicationContext
Public Property Get AppContext() As clsApplicationContext
    Set AppContext = mAppContext
End Property

'=============================================================================
' GESTIÓN DE CICLO DE VIDA
'=============================================================================

'@Description: Libera todos los servicios
Public Sub DisposeAll()
    On Error Resume Next

    LogInfo MODULE_NAME, "Liberando " & mServices.Count & " servicios..."

    Dim key As Variant
    Dim service As IService

    For Each key In mServices.Keys
        Set service = mServices(key)
        If service.IsInitialized Then
            service.Dispose
            LogInfo MODULE_NAME, "Servicio liberado: " & key
        End If
    Next key

    mServices.RemoveAll
    Set mAppContext = Nothing
    mIsInitialized = False

    LogInfo MODULE_NAME, "Servicios liberados"
End Sub

'=============================================================================
' MÉTODO LEGACY (Para compatibilidad temporal)
'=============================================================================

'@Description: Obtiene un servicio por nombre (DEPRECATED - usar propiedades tipadas)
'@Note: Mantener solo para migración, eliminar cuando todo use propiedades tipadas
Public Function GetService(ByVal serviceName As String) As Object
    LogWarning MODULE_NAME, "[GetService] DEPRECATED: Usar propiedades tipadas en lugar de GetService(""" & serviceName & """)"

    ' Mapear nombres legacy a TypeNames
    Dim key As String
    Select Case serviceName
        Case "IConfiguration": key = "clsConfiguration"
        Case "IExecutionContext": key = "clsExecutionContext"
        Case "IFileManager": key = "clsFileManager"
        Case "IOpportunities": key = "clsOpportunitiesMgr"
        Case "IChartManager": key = "clsChartEventsManager"
        Case "IFSMonitoring": key = "clsFSMonitoringCoord"
        Case "IRibbonUI": key = "clsRibbonUI"
        Case Else: key = serviceName
    End Select

    If Not mServices.Exists(key) Then
        Err.Raise 5001, MODULE_NAME, "Servicio no registrado: " & serviceName & " (key: " & key & ")"
    End If

    Dim service As IService
    Set service = mServices(key)
    Call EnsureInitialized(service)

    Set GetService = mServices(key)
End Function

