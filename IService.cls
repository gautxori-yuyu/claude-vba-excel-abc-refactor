VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "3.Interfaces"
'@Interface
Option Explicit

'=============================================================================
' IService - Interfaz base para todos los servicios
'=============================================================================
' Propósito: Estandarizar el ciclo de vida de servicios
'
' Todos los servicios deben implementar esta interfaz con:
'   Implements IService
'
' Ejemplo de implementación:
'   Private mIsInitialized As Boolean
'
'   Private Sub IService_Initialize(ByVal dependencies As Object)
'       Dim svcMgr As clsServiceManager
'       Set svcMgr = dependencies
'
'       ' Obtener otras dependencias si necesitas
'       ' Set mConfig = svcMgr.GetService("IConfiguration")
'
'       mIsInitialized = True
'   End Sub
'
'   Private Sub IService_Dispose()
'       ' Liberar recursos
'       mIsInitialized = False
'   End Sub
'
'   Private Property Get IService_IsInitialized() As Boolean
'       IService_IsInitialized = mIsInitialized
'   End Property
'
'   Private Property Get IService_ServiceName() As String
'       IService_ServiceName = "MiServicio"
'   End Property
'=============================================================================

'@Description: Inicializa el servicio con dependencias
'@Param dependencies: ServiceManager desde donde obtener otros servicios
Public Sub Initialize(ByVal dependencies As Object)
    Err.Raise 5, "IService", "Initialize debe implementarse en clase concreta"
End Sub

'@Description: Libera recursos del servicio
Public Sub Dispose()
    Err.Raise 5, "IService", "Dispose debe implementarse en clase concreta"
End Sub

'@Description: Indica si el servicio está inicializado
Public Property Get IsInitialized() As Boolean
    Err.Raise 5, "IService", "IsInitialized debe implementarse en clase concreta"
End Property

'@Description: Nombre único del servicio (para logging/debug)
Public Property Get ServiceName() As String
    Err.Raise 5, "IService", "ServiceName debe implementarse en clase concreta"
End Property

