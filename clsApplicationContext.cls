VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsApplicationContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "2.Infrastructure"
Option Explicit

'=============================================================================
' clsApplicationContext - Estado compartido de la aplicacion
'=============================================================================
' Responsabilidades:
'   1. Agregar/componer sub-estados (RibbonState, ChartState, ExecutionContext)
'   2. Proporcionar acceso compartido al estado de la aplicacion
'   3. Servir como punto central de acceso al estado (no a servicios)
'=============================================================================
' NOTA: Esta clase NO es un servicio. NO implementa IService.
'       Es un objeto de estado puro, creado y gestionado por clsAplicacion.
'       Los servicios acceden a ella via ServiceManager.AppContext
'=============================================================================

Private Const MODULE_NAME As String = "clsApplicationContext"

'=============================================================================
' SUB-ESTADOS ESPECIALIZADOS (Composicion)
'=============================================================================
Private mRibbonState As clsRibbonState
Private mExecutionContext As clsExecutionContext
Private mChartState As clsChartState

'=============================================================================
' ESTADO DE DOMINIO
'=============================================================================
Private m_CurrentOpportunity As Object     ' clsOpportunity
Private m_CurrentFile As Object            ' clsExcelFile
Private m_ApplicationState As Long         ' Enum: Running, Initializing, ShuttingDown

'=============================================================================
' INICIALIZACION
'=============================================================================

Public Sub Initialize(ByVal RibbonState As clsRibbonState, _
                      ByVal execContext As clsExecutionContext, _
                      ByVal ChartState As clsChartState)
    Set mRibbonState = RibbonState
    Set mExecutionContext = execContext
    Set mChartState = ChartState
End Sub

'=============================================================================
' PROPIEDADES - Delegacion a sub-estados
'=============================================================================

Public Property Get RibbonState() As clsRibbonState
    Set RibbonState = mRibbonState
End Property

Public Property Get ExecutionContext() As clsExecutionContext
    Set ExecutionContext = mExecutionContext
End Property

Public Property Get ChartState() As clsChartState
    Set ChartState = mChartState
End Property

'=============================================================================
' PROPIEDADES - Estado de dominio
'=============================================================================
'@Description: Oportunidad actualmente seleccionada
Public Property Get CurrentOpportunity() As Object
    Set CurrentOpportunity = m_CurrentOpportunity
End Property

Public Property Set CurrentOpportunity(ByVal Value As Object)
    Set m_CurrentOpportunity = Value
End Property

'@Description: Archivo de Excel actualmente activo (wrapper clsExcelFile)
Public Property Get CurrentFile() As Object
    Set CurrentFile = m_CurrentFile
End Property

Public Property Set CurrentFile(ByVal Value As Object)
    Set m_CurrentFile = Value
End Property

'@Description: Estado de la aplicaci�n (inicializando, ejecutando, cerrando)
Public Property Get ApplicationState() As Long
    ApplicationState = m_ApplicationState
End Property

Public Property Let ApplicationState(ByVal Value As Long)
    m_ApplicationState = Value
End Property

'=============================================================================
' PROPIEDADES CONVENIENTES - Acceso directo a sub-estados
'=============================================================================

'@Description: Indica si hay un gr�fico activo en este momento
Public Property Get IsChartActive() As Boolean
    IsChartActive = mChartState.IsActive
End Property

Public Property Get CanInvertAxes() As Boolean
    CanInvertAxes = mChartState.CanInvertAxes
End Property

Public Property Let CanInvertAxes(ByVal Value As Boolean)
    mChartState.CanInvertAxes = Value
End Property

Public Property Get CanGenerateChart() As Boolean
    CanGenerateChart = mChartState.CanGenerateChart
End Property

Public Property Let CanGenerateChart(ByVal Value As Boolean)
    mChartState.CanGenerateChart = Value
End Property

' Execution
Public Property Get CurrentWorkbook() As Workbook
    Set CurrentWorkbook = mExecutionContext.Workbook
End Property

Public Property Get CurrentWorksheet() As Worksheet
    Set CurrentWorksheet = mExecutionContext.Worksheet
End Property

Public Property Get CurrentChart() As Chart
    Set CurrentChart = mExecutionContext.Chart
End Property

' Ribbon
Public Property Get RibbonMode() As eRibbonMode
    RibbonMode = mRibbonState.Modo
End Property

'=============================================================================
' M�TODOS HELPER
'=============================================================================

'@Description: Resetea todo el estado a valores iniciales
Public Sub Reset()
    Set m_CurrentOpportunity = Nothing
    Set m_CurrentFile = Nothing
    m_ApplicationState = 0
    
    ' Los sub-estados se resetean a si mismos si hace falta
    If Not mChartState Is Nothing Then mChartState.Reset
End Sub

'@Description: Informaci�n de debug del estado actual
Public Function GetDebugInfo() As String
    Dim info As String
    
    info = "=== APPLICATION CONTEXT ===" & vbCrLf
    info = info & "CurrentOpportunity: " & IIf(m_CurrentOpportunity Is Nothing, "None", "Set") & vbCrLf
    info = info & "CurrentFile: " & IIf(m_CurrentFile Is Nothing, "None", "Set") & vbCrLf
    info = info & "ApplicationState: " & m_ApplicationState & vbCrLf
    info = info & vbCrLf
    
    If Not mExecutionContext Is Nothing Then
        info = info & mExecutionContext.Diagnostics & vbCrLf
    End If
    
    If Not mChartState Is Nothing Then
        info = info & mChartState.GetDebugInfo & vbCrLf
    End If
    
    If Not mRibbonState Is Nothing Then
        info = info & "RibbonMode: " & mRibbonState.RibbonStateDescription
    End If
    
    GetDebugInfo = info
End Function


