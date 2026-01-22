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

' ==========================================
' SERVICIOS (SOLO CICLO DE VIDA - SIN EVENTOS)
' ==========================================

Private mConfig As clsApplicationConfiguration
Private mState As clsApplicationState
Private mExcelContext As clsExcelExecutionContext
Private mFileManager As clsFileManager
Private mOpportunityMgr As clsOpportunityManager
Private mRibbonMgr As clsRibbonManager
Private mFSMonitor As clsFileSystemMonitor

Private mIsInitialized As Boolean

' ==========================================
' PROPIEDADES (FACADE)
' ==========================================

Public Property Get Configuration() As clsApplicationConfiguration
    Set Configuration = mConfig
End Property

Public Property Get State() As clsApplicationState
    Set State = mState
End Property

Public Property Get FileManager() As clsFileManager
    Set FileManager = mFileManager
End Property

Public Property Get OpportunityManager() As clsOpportunityManager
    Set OpportunityManager = mOpportunityMgr
End Property

Public Property Get RibbonManager() As clsRibbonManager
    Set RibbonManager = mRibbonMgr
End Property

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
    
    ' ORDEN CRÍTICO: Sin dependencias → Con dependencias
    
    ' 1. Configuración (sin dependencias)
    Set mConfig = New clsApplicationConfiguration
    mConfig.LoadFromRegistry
    LogInfo MODULE_NAME, "- Configuración cargada"
    
    ' 2. Estado global (sin dependencias)
    Set mState = New clsApplicationState
    LogInfo MODULE_NAME, "- Estado inicializado"
    
    ' 3. Contexto Excel (sin dependencias, solo suscribe eventos)
    Set mExcelContext = New clsExcelExecutionContext
    mExcelContext.Initialize
    LogInfo MODULE_NAME, "- Contexto Excel inicializado"
    
    ' 4. FileManager (necesita ExcelContext)
    Set mFileManager = New clsFileManager
    ' TODO Sprint 2: mFileManager.Initialize mExcelContext
    LogInfo MODULE_NAME, "- FileManager inicializado"
    
    ' 5. OpportunityManager (necesita Config)
    Set mOpportunityMgr = New clsOpportunityManager
    ' TODO Sprint 2: mOpportunityMgr.SetBaseFolder mConfig.OpportunitiesPath
    LogInfo MODULE_NAME, "- OpportunityManager inicializado"
    
    ' 6. RibbonManager (necesita State)
    Set mRibbonMgr = New clsRibbonManager
    ' TODO Sprint 2: mRibbonMgr.Initialize mState
    LogInfo MODULE_NAME, "- RibbonManager inicializado"
    
    ' 7. FileSystemMonitor (necesita Config)
    Set mFSMonitor = New clsFileSystemMonitor
    ' TODO Sprint 3: mFSMonitor.StartMonitoring mConfig.FoldersToWatch
    LogInfo MODULE_NAME, "- FileSystemMonitor inicializado"
    
    mIsInitialized = True
    LogInfo MODULE_NAME, "Aplicación iniciada correctamente"
    
    Exit Sub
    
ErrHandler:
    LogError MODULE_NAME, "Error en Initialize", Err.Number, Err.Description
    mIsInitialized = False
    ' Re-lanzar error para que ThisWorkbook lo gestione
    Err.Raise Err.Number, "clsApplication.Initialize", Err.Description
End Sub

' ==========================================
' TERMINACION (CLEANUP)
' ==========================================

Public Sub Terminate()
    On Error Resume Next ' Asegurar limpieza completa incluso con errores
    
    LogInfo MODULE_NAME, "Terminando aplicación..."
    
    ' Orden inverso a inicialización
    
    ' 7. FileSystemMonitor
    If Not mFSMonitor Is Nothing Then
        ' TODO Sprint 3: mFSMonitor.StopMonitoring
        Set mFSMonitor = Nothing
        LogInfo MODULE_NAME, "- FileSystemMonitor terminado"
    End If
    
    ' 6. RibbonManager
    If Not mRibbonMgr Is Nothing Then
        ' TODO Sprint 2: mRibbonMgr.Terminate
        Set mRibbonMgr = Nothing
        LogInfo MODULE_NAME, "- RibbonManager terminado"
    End If
    
    ' 5. OpportunityManager
    If Not mOpportunityMgr Is Nothing Then
        Set mOpportunityMgr = Nothing
        LogInfo MODULE_NAME, "- OpportunityManager terminado"
    End If
    
    ' 4. FileManager
    If Not mFileManager Is Nothing Then
        Set mFileManager = Nothing
        LogInfo MODULE_NAME, "- FileManager terminado"
    End If
    
    ' 3. Contexto Excel
    If Not mExcelContext Is Nothing Then
        mExcelContext.Terminate
        Set mExcelContext = Nothing
        LogInfo MODULE_NAME, "- Contexto Excel terminado"
    End If
    
    ' 2. Estado
    Set mState = Nothing
    
    ' 1. Configuración
    If Not mConfig Is Nothing Then
        mConfig.SaveToRegistry
        Set mConfig = Nothing
        LogInfo MODULE_NAME, "- Configuración guardada"
    End If
    
    mIsInitialized = False
    LogInfo MODULE_NAME, "Aplicación terminada correctamente"
    
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
