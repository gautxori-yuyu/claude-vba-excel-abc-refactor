VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("1-Application")
Option Explicit

' ===============================================================================
' CLASE: clsApplication - Composition Root
' ===============================================================================
' PROPOSITO: Crear, configurar y coordinar ciclo de vida de todos los servicios
' MIGRADO DESDE: main:clsAplicacion.cls (479 lineas)
' PATRÓN: Composition Root + Facade
' ===============================================================================

Private Const MODULE_NAME As String = "clsApplication"

Private mIsInitialized As Boolean

' TODO Sprint 2: Aquí se declararán los servicios cuando se implementen
' Private mConfig As clsApplicationConfiguration
' Private mState As clsApplicationState
' etc.

Public Property Get IsInitialized() As Boolean
    IsInitialized = mIsInitialized
End Property

' ==========================================
' INICIALIZACION (COMPOSITION ROOT)
' ==========================================

Public Sub Initialize()
    On Error GoTo ErrHandler

    If mIsInitialized Then
        LogWarning MODULE_NAME, "Initialize llamado múltiples veces - ignorando"
        Exit Sub
    End If

    LogInfo MODULE_NAME, "Iniciando aplicación..."

    ' TODO Sprint 2: Crear e inicializar servicios aquí
    ' ORDEN: Sin dependencias → Con dependencias
    ' Ejemplo:
    ' Set mConfig = New clsApplicationConfiguration
    ' mConfig.LoadFromRegistry
    ' Set mExcelContext = New clsExcelExecutionContext
    ' mExcelContext.Initialize
    ' etc.

    mIsInitialized = True
    LogInfo MODULE_NAME, "Aplicación iniciada (sin servicios por ahora)"

    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "Error en Initialize", Err.Number, Err.Description
    mIsInitialized = False
    Err.Raise Err.Number, "clsApplication.Initialize", Err.Description
End Sub

' ==========================================
' TERMINACION (CLEANUP)
' ==========================================

Public Sub Terminate()
    On Error Resume Next ' Asegurar limpieza completa incluso con errores

    LogInfo MODULE_NAME, "Terminando aplicación..."

    ' TODO Sprint 2: Liberar servicios en orden inverso
    ' Ejemplo:
    ' If Not mFSMonitor Is Nothing Then
    '     mFSMonitor.StopMonitoring
    '     Set mFSMonitor = Nothing
    ' End If
    ' If Not mExcelContext Is Nothing Then
    '     mExcelContext.Terminate
    '     Set mExcelContext = Nothing
    ' End If
    ' If Not mConfig Is Nothing Then
    '     mConfig.SaveToRegistry
    '     Set mConfig = Nothing
    ' End If
    ' etc.

    mIsInitialized = False
    LogInfo MODULE_NAME, "Aplicación terminada (sin servicios por ahora)"

    On Error GoTo 0
End Sub

' ==========================================
' FINALIZER (SEGURIDAD)
' ==========================================

Private Sub Class_Terminate()
    ' Backup de seguridad si se olvida llamar Terminate()
    If mIsInitialized Then
        LogWarning MODULE_NAME, "Class_Terminate: Terminate() no fue llamado explícitamente"
        Terminate
    End If
End Sub
