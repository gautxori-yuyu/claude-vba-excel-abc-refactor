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
'   3. Gestionar ciclo de vida (Initialize/Dispose v�a IService)
'   4. Inyectar dependencias
'=============================================================================

Private Const MODULE_NAME As String = "clsServiceManager"

'=============================================================================
' ENUMS
'=============================================================================

Public Enum ServiceLifetime
    Singleton = 1   ' Una instancia para toda la aplicaci�n
    Transient = 2   ' Nueva instancia cada vez que se solicita
End Enum

'=============================================================================
' TYPES
'=============================================================================

Private Type ServiceDescriptor
    ServiceName As String
    Lifetime As ServiceLifetime
    instance As Object              ' Para Singleton
    factoryFunction As String       ' Para Transient (nombre funci�n)
    IsInitialized As Boolean
End Type

'=============================================================================
' ESTADO
'=============================================================================
Private mServices() As ServiceDescriptor
Private mServiceCount As Long

'=============================================================================
' INICIALIZACI�N
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
    On Error GoTo ErrHandler

    If instance Is Nothing Then
        Err.Raise 5003, MODULE_NAME, "No se puede registrar un servicio Nothing"
    End If

    Dim instanceService As IService
    Set instanceService = instance

    If instanceService Is Nothing Then
        Err.Raise 5004, MODULE_NAME, "El objeto no implementa IService"
    End If

    Dim svcName As String
    svcName = instanceService.ServiceName

    If Len(Trim(svcName)) = 0 Then
        Err.Raise 5005, MODULE_NAME, "ServiceName no puede estar vacio"
    End If

    Dim Index As Long
    Index = FindServiceIndex(svcName)

    Dim descriptor As ServiceDescriptor

    descriptor.ServiceName = svcName
    descriptor.Lifetime = ServiceLifetime.Singleton
    Set descriptor.instance = instance
    descriptor.IsInitialized = False
    descriptor.factoryFunction = ""

    If Index >= 0 Then
        ' Reemplazar servicio existente
        mServices(Index) = descriptor
        LogInfo MODULE_NAME, "Singleton actualizado: " & svcName
    Else
        ' Agregar nuevo servicio
        Call AddService(descriptor)
        LogInfo MODULE_NAME, "Singleton registrado: " & svcName
    End If

    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "Error registrando singleton", Err.Number, Err.Description
    Err.Raise Err.Number, MODULE_NAME, Err.Description
End Sub

'@Description: Registra un servicio Transient (nueva instancia cada vez)
Public Sub RegisterTransient(ByVal ServiceName As String, ByVal factoryFunction As String)
    On Error GoTo ErrHandler

    If Len(Trim(ServiceName)) = 0 Then
        Err.Raise 5006, MODULE_NAME, "ServiceName no puede estar vacio para Transient"
    End If

    If Len(Trim(factoryFunction)) = 0 Then
        Err.Raise 5007, MODULE_NAME, "factoryFunction no puede estar vacio"
    End If

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

    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "Error registrando transient: " & ServiceName, Err.Number, Err.Description
    Err.Raise Err.Number, MODULE_NAME, Err.Description
End Sub

'=============================================================================
' RESOLUCI�N DE SERVICIOS
'=============================================================================

'@Description: Obtiene un servicio (lo inicializa si es necesario)
Public Function GetService(ByVal ServiceName As String) As Object
    On Error GoTo ErrHandler

    If Len(Trim(ServiceName)) = 0 Then
        Err.Raise 5008, MODULE_NAME, "ServiceName no puede estar vacio"
    End If

    Dim Index As Long
    Index = FindServiceIndex(ServiceName)

    If Index < 0 Then
        Err.Raise 5001, MODULE_NAME, "Servicio no registrado: " & ServiceName
    End If

    Dim descriptor As ServiceDescriptor
    descriptor = mServices(Index)

    Select Case descriptor.Lifetime
        Case ServiceLifetime.Singleton
            ' Lazy initialization para Singletons
            If Not descriptor.IsInitialized Then
                Call InitializeServiceInstance(descriptor.instance)
                descriptor.IsInitialized = True
                mServices(Index) = descriptor
                LogDebug MODULE_NAME, "Singleton inicializado: " & ServiceName
            End If

            Set GetService = descriptor.instance

        Case ServiceLifetime.Transient
            ' Crear nueva instancia para Transient
            Dim newInstance As Object
            Set newInstance = CreateInstance(descriptor.factoryFunction)

            If newInstance Is Nothing Then
                Err.Raise 5009, MODULE_NAME, "Factory retorno Nothing para: " & ServiceName
            End If

            Call InitializeServiceInstance(newInstance)
            LogDebug MODULE_NAME, "Transient creado: " & ServiceName

            Set GetService = newInstance

        Case Else
            Err.Raise 5010, MODULE_NAME, "Lifetime desconocido para: " & ServiceName
    End Select

    Exit Function

ErrHandler:
    LogError MODULE_NAME, "Error obteniendo servicio: " & ServiceName, Err.Number, Err.Description
    Err.Raise Err.Number, MODULE_NAME, Err.Description
End Function

'@Description: Verifica si un servicio est� registrado
Public Function IsRegistered(ByVal ServiceName As String) As Boolean
    IsRegistered = (FindServiceIndex(ServiceName) >= 0)
End Function

'@Description: Devuelve el n�mero de servicios registrados
Public Function GetServiceCount() As Long
    GetServiceCount = mServiceCount
End Function

'=============================================================================
' GESTI�N DE CICLO DE VIDA
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

'@Description: Libera todos los servicios en orden inverso al registro
Public Sub DisposeAll()
    On Error Resume Next

    If mServiceCount = 0 Then
        LogInfo MODULE_NAME, "No hay servicios para liberar"
        Exit Sub
    End If

    LogInfo MODULE_NAME, "Liberando " & mServiceCount & " servicios..."

    ' Liberar en orden inverso (LIFO) para respetar dependencias
    Dim i As Long
    For i = mServiceCount - 1 To 0 Step -1
        With mServices(i)
            If .Lifetime = ServiceLifetime.Singleton Then
                If Not .instance Is Nothing Then
                    LogDebug MODULE_NAME, "Liberando: " & .ServiceName
                    Call DisposeService(.instance)
                    Set .instance = Nothing
                End If
            End If
        End With
    Next i

    mServiceCount = 0
    Erase mServices
    ReDim mServices(0 To 9)

    LogInfo MODULE_NAME, "Todos los servicios liberados correctamente"
End Sub

'@Description: Devuelve lista de nombres de servicios registrados (para diagnostico)
Public Function GetRegisteredServiceNames() As String
    If mServiceCount = 0 Then
        GetRegisteredServiceNames = "(ninguno)"
        Exit Function
    End If

    Dim result As String
    Dim i As Long

    For i = 0 To mServiceCount - 1
        If i > 0 Then result = result & ", "
        result = result & mServices(i).ServiceName
        If mServices(i).IsInitialized Then
            result = result & " [init]"
        End If
    Next i

    GetRegisteredServiceNames = result
End Function

'@Description: Libera un servicio espec�fico
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

'@Description: Crea una instancia usando una funci�n factory
Private Function CreateInstance(ByVal factoryFunction As String) As Object
    On Error GoTo ErrHandler
    
    ' Llamar a funci�n factory p�blica
    Set CreateInstance = Application.Run(factoryFunction)
    Exit Function
    
ErrHandler:
    Err.Raise 5002, MODULE_NAME, _
        "Error creando instancia con factory: " & factoryFunction & vbCrLf & Err.Description
End Function


