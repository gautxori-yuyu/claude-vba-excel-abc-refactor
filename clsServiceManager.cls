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
' clsServiceManager - Gestor de servicios con ciclo de vida
'=============================================================================
' Responsabilidades:
'   1. Registrar servicios (Singleton o Transient)
'   2. Resolver servicios (con lazy initialization)
'   3. Gestionar ciclo de vida (Initialize/Dispose vía IService)
'   4. Inyectar dependencias
'=============================================================================

Private Const MODULE_NAME As String = "clsServiceManager"

'=============================================================================
' ENUMS
'=============================================================================

Public Enum ServiceLifetime
    Singleton = 1   ' Una instancia para toda la aplicación
    Transient = 2   ' Nueva instancia cada vez que se solicita
End Enum

'=============================================================================
' TYPES
'=============================================================================

Private Type ServiceDescriptor
    ServiceName As String
    Lifetime As ServiceLifetime
    instance As Object              ' Para Singleton
    factoryFunction As String       ' Para Transient (nombre función)
    IsInitialized As Boolean
End Type

'=============================================================================
' ESTADO
'=============================================================================
Private mServices() As ServiceDescriptor
Private mServiceCount As Long

'=============================================================================
' INICIALIZACIÓN
'=============================================================================

Private Sub Class_Initialize()
    mServiceCount = 0
    ReDim mServices(0 To 9) ' Preasignar espacio inicial (opcional)
End Sub

Private Sub Class_Terminate()
    Call DisposeAll
End Sub

'=============================================================================
' REGISTRO DE SERVICIOS
'=============================================================================

'@Description: Registra un servicio Singleton (una instancia compartida)
Public Sub RegisterSingleton(ByVal instance As Object)
    Dim instanceService As IService
    Set instanceService = instance
    
    ' TODO: añadir chequero de errores
    
    Dim Index As Long
    Index = FindServiceIndex(instanceService.ServiceName)
    
    Dim descriptor As ServiceDescriptor
    
    descriptor.ServiceName = instanceService.ServiceName
    descriptor.Lifetime = ServiceLifetime.Singleton
    Set descriptor.instance = instance
    descriptor.IsInitialized = False
    descriptor.factoryFunction = ""
    
    If Index >= 0 Then
        ' Reemplazar
        mServices(Index) = descriptor
        LogInfo MODULE_NAME, "Singleton actualizado: " & ServiceName
    Else
        ' Añadir nuevo
        Call AddService(descriptor)
        LogInfo MODULE_NAME, "Singleton registrado: " & ServiceName
    End If
End Sub

'@Description: Registra un servicio Transient (nueva instancia cada vez)
Public Sub RegisterTransient(ByVal ServiceName As String, ByVal factoryFunction As String)
    Dim Index As Long
    Index = FindServiceIndex(ServiceName)
    
    Dim descriptor As ServiceDescriptor
    
    descriptor.ServiceName = ServiceName
    descriptor.Lifetime = ServiceLifetime.Transient
    Set descriptor.instance = Nothing
    descriptor.IsInitialized = False
    descriptor.factoryFunction = factoryFunction
    
    If Index >= 0 Then
        mServices(Index) = descriptor
        LogInfo MODULE_NAME, "Transient actualizado: " & ServiceName
    Else
        Call AddService(descriptor)
        LogInfo MODULE_NAME, "Transient registrado: " & ServiceName
    End If
End Sub

'=============================================================================
' RESOLUCIÓN DE SERVICIOS
'=============================================================================

'@Description: Obtiene un servicio (lo inicializa si es necesario)
Public Function GetService(ByVal ServiceName As String) As Object
    Dim Index As Long
    Index = FindServiceIndex(ServiceName)
    
    If Index < 0 Then
        Err.Raise 5001, MODULE_NAME, "Servicio no registrado: " & ServiceName
    End If
    
    Dim descriptor As ServiceDescriptor
    descriptor = mServices(Index)
    
    Select Case descriptor.Lifetime
        Case ServiceLifetime.Singleton
            ' Inicializar si hace falta (lazy initialization)
            If Not descriptor.IsInitialized Then
                Call InitializeServiceInstance(descriptor.instance)
                descriptor.IsInitialized = True
                mServices(Index) = descriptor ' Guardar estado actualizado
            End If
            
            Set GetService = descriptor.instance
            
        Case ServiceLifetime.Transient
            ' Crear nueva instancia
            Dim newInstance As Object
            Set newInstance = CreateInstance(descriptor.factoryFunction)
            
            ' Inicializar
            Call InitializeServiceInstance(newInstance)
            
            Set GetService = newInstance
    End Select
End Function

'@Description: Verifica si un servicio está registrado
Public Function IsRegistered(ByVal ServiceName As String) As Boolean
    IsRegistered = (FindServiceIndex(ServiceName) >= 0)
End Function

'@Description: Devuelve el número de servicios registrados
Public Function GetServiceCount() As Long
    GetServiceCount = mServiceCount
End Function

'=============================================================================
' GESTIÓN DE CICLO DE VIDA
'=============================================================================

'@Description: Inicializa un servicio Singleton
Private Sub InitializeService(ByRef descriptor As ServiceDescriptor)
    Call InitializeServiceInstance(descriptor.instance)
End Sub

'@Description: Inicializa una instancia de servicio (llama a IService.Initialize)
Private Sub InitializeServiceInstance(ByVal service As Object)
    On Error Resume Next
    
    ' Intentar obtener como IService
    Dim svc As IService
    Set svc = service
    
    If Not svc Is Nothing Then
        ' Pasar el ServiceManager como dependencia (DI)
        svc.Initialize Me
        
        LogInfo MODULE_NAME, "Servicio inicializado: " & svc.ServiceName
    End If
    
    On Error GoTo 0
End Sub

'@Description: Libera todos los servicios
Public Sub DisposeAll()
    On Error Resume Next
    LogInfo MODULE_NAME, "Liberando " & mServiceCount & " servicios..."
    
    Dim i As Long
    For i = 0 To mServiceCount - 1
        With mServices(i)
            If .Lifetime = ServiceLifetime.Singleton Then
                Call DisposeService(.instance)
            End If
        End With
    Next i
    
    mServiceCount = 0
    Erase mServices
    ReDim mServices(0 To 9)
    
    LogInfo MODULE_NAME, "Servicios liberados"
End Sub

'@Description: Libera un servicio específico
Private Sub DisposeService(ByVal service As Object)
    On Error Resume Next
    
    Dim svc As IService
    Set svc = service
    
    If Not svc Is Nothing Then
        svc.Dispose
        LogInfo MODULE_NAME, "Servicio liberado: " & svc.ServiceName
    End If
    
    Set service = Nothing
End Sub

'=============================================================================
' HELPERS
'=============================================================================
Private Function FindServiceIndex(ByVal ServiceName As String) As Long
    Dim i As Long
    For i = 0 To mServiceCount - 1
        If StrComp(mServices(i).ServiceName, ServiceName, vbTextCompare) = 0 Then
            FindServiceIndex = i
            Exit Function
        End If
    Next i
    FindServiceIndex = -1
End Function

Private Sub AddService(descriptor As ServiceDescriptor)
    ' Asegurar espacio
    If mServiceCount >= UBound(mServices) Then
        ReDim Preserve mServices(0 To UBound(mServices) + 10)
    End If
    mServices(mServiceCount) = descriptor
    mServiceCount = mServiceCount + 1
End Sub

'@Description: Crea una instancia usando una función factory
Private Function CreateInstance(ByVal factoryFunction As String) As Object
    On Error GoTo ErrHandler
    
    ' Llamar a función factory pública
    Set CreateInstance = Application.Run(factoryFunction)
    Exit Function
    
ErrHandler:
    Err.Raise 5002, MODULE_NAME, _
        "Error creando instancia con factory: " & factoryFunction & vbCrLf & Err.Description
End Function


