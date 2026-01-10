VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChartState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "2.Infrastructure.State"
Option Explicit

'=============================================================================
' clsChartState - Estado relacionado con graficos
'=============================================================================

Private mIsActive As Boolean
Private mCurrentChart As Chart
Private mCanInvertAxes As Boolean
Private mCanGenerateChart As Boolean

'=============================================================================
' PROPIEDADES
'=============================================================================

Public Property Get IsActive() As Boolean
    IsActive = mIsActive
End Property

Public Property Let IsActive(ByVal Value As Boolean)
    mIsActive = Value
End Property

Public Property Get CurrentChart() As Chart
    Set CurrentChart = mCurrentChart
End Property

Public Property Set CurrentChart(ByVal Value As Chart)
    Set mCurrentChart = Value
    mIsActive = Not (Value Is Nothing)
End Property

Public Property Get CanInvertAxes() As Boolean
    CanInvertAxes = mCanInvertAxes
End Property

Public Property Let CanInvertAxes(ByVal Value As Boolean)
    mCanInvertAxes = Value
End Property

Public Property Get CanGenerateChart() As Boolean
    CanGenerateChart = mCanGenerateChart
End Property

Public Property Let CanGenerateChart(ByVal Value As Boolean)
    mCanGenerateChart = Value
End Property

'=============================================================================
' METODOS
'=============================================================================

Public Sub Reset()
    mIsActive = False
    Set mCurrentChart = Nothing
    mCanInvertAxes = False
    mCanGenerateChart = False
End Sub

Public Function GetDebugInfo() As String
    Dim info As String
    info = "ChartState:" & vbCrLf
    info = info & "  IsActive: " & mIsActive & vbCrLf
    info = info & "  CurrentChart: " & IIf(mCurrentChart Is Nothing, "None", mCurrentChart.Name) & vbCrLf
    info = info & "  CanInvertAxes: " & mCanInvertAxes & vbCrLf
    info = info & "  CanGenerateChart: " & mCanGenerateChart
    GetDebugInfo = info
End Function
