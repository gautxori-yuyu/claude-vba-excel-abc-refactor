VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRibbonUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "2-Servicios.Excel.Ribbon"
Option Explicit

'=============================================================================
' IMPLEMENTACION IService
'=============================================================================
Implements IService

Private mIsInitialized As Boolean

'=============================================================================
' RIBBON UI
'=============================================================================

Private Const MODULE_NAME As String = "clsRibbonUI"

Private mRibbonUI As IRibbonUI
Private mIsRecovering As Boolean              ' Flag para evitar recursion durante recuperacion
Private mWasEverInitialized As Boolean        ' Flag para distinguir "nunca inicializado" de "perdido"

'-----------------------------------------------------------------------------
' IService Implementation
'-----------------------------------------------------------------------------

Private Sub IService_Initialize(ByVal dependencies As Object)
    If mIsInitialized Then Exit Sub
    LogInfo MODULE_NAME, "[IService_Initialize]"
    mWasEverInitialized = False
    mIsRecovering = False
    mIsInitialized = True
End Sub

Private Sub IService_Dispose()
    LogInfo MODULE_NAME, "[IService_Dispose]"
    Call StopEvents
    mIsInitialized = False
End Sub

Private Property Get IService_IsInitialized() As Boolean
    IService_IsInitialized = mIsInitialized
End Property

Private Property Get IService_ServiceName() As String
    IService_ServiceName = "RibbonUI"
End Property

'=============================================================================
' RIBBON UI MANAGEMENT
'=============================================================================

'@Description: Obtiene el puntero al Ribbon (puede ser Nothing)
Public Property Get RibbonUI() As IRibbonUI
    Set RibbonUI = mRibbonUI
End Property


Public Sub Init(ByRef ribbonObj As IRibbonUI)
    LogDebug MODULE_NAME, "[Init] - Recibiendo puntero IRibbonUI"
    Set mRibbonUI = ribbonObj

    If ribbonObj Is Nothing Then
        LogWarning MODULE_NAME, "[Init] - Se paso Nothing como puntero"
    Else
        LogDebug MODULE_NAME, "[Init] - Puntero establecido correctamente"
    End If
    mWasEverInitialized = True
End Sub

Public Sub StopEvents()
    LogDebug MODULE_NAME, "[StopEvents] - Liberando puntero"
    Set mRibbonUI = Nothing
End Sub

'@Description: Invalida todo el Ribbon, forzando recarga de callbacks
'@Note: Intenta recuperacion automatica si el Ribbon esta perdido
Public Sub InvalidarRibbon()
    On Error GoTo ErrHandler

    ' Verificar si el Ribbon esta disponible
    If Not IsRibbonUIAvailable() Then
        ' IMPORTANTE: Solo intentar recuperacion si el ribbon FUE inicializado alguna vez
        ' Esto evita que durante Class_Initialize se intente recuperar un ribbon que
        ' aun no ha sido configurado (el callback RibbonOnLoad aun no ha terminado)
        If Not mWasEverInitialized Then
            LogDebug MODULE_NAME, "[InvalidarRibbon] Ribbon aun no inicializado, omitiendo"
            Exit Sub
        End If

        ' Intentar recuperacion automatica (solo si no estamos ya recuperando)
        If Not mIsRecovering Then
            LogWarning MODULE_NAME, "[InvalidarRibbon] Ribbon perdido, intentando recuperacion..."
            If TryAutoRecover() Then
                LogDebug MODULE_NAME, "[InvalidarRibbon] Recuperacion exitosa"
            Else
                LogError MODULE_NAME, "[InvalidarRibbon] Recuperacion fallida"
                Exit Sub
            End If
        Else
            LogDebug MODULE_NAME, "[InvalidarRibbon] Recuperacion en curso, omitiendo"
            Exit Sub
        End If
    End If

    ' Invalidar el Ribbon
    mRibbonUI.Invalidate
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[InvalidarRibbon] Error", Err.Description
    ' No propagar el error, solo loguear
End Sub

'@Description: Invalida un control especifico del Ribbon
'@Note: Intenta recuperacion automatica si el Ribbon esta perdido
Public Sub InvalidarControl(idControl As String)
    On Error GoTo ErrHandler

    ' Verificar si el Ribbon esta disponible
    If Not IsRibbonUIAvailable() Then
        ' IMPORTANTE: Solo intentar recuperacion si el ribbon FUE inicializado alguna vez
        If Not mWasEverInitialized Then
            LogDebug MODULE_NAME, "[InvalidarControl] Ribbon aun no inicializado, omitiendo: " & idControl
            Exit Sub
        End If

        ' Intentar recuperacion automatica
        If Not mIsRecovering Then
            LogWarning MODULE_NAME, "[InvalidarControl] Ribbon perdido, intentando recuperacion..."
            If Not TryAutoRecover() Then
                LogError MODULE_NAME, "[InvalidarControl] Recuperacion fallida para control: " & idControl
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If

    ' Invalidar el control
    mRibbonUI.InvalidateControl idControl
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[InvalidarControl] Error en control " & idControl, , Err.Description
End Sub


'@Description: Activa una pestaña del Ribbon por id
'@Scope: Friend (uso exclusivo desde clsAplicacion)
'@ArgumentDescriptions: tabId | id del tab a activar
'@Category: UI / Ribbon
Friend Sub ActivarTab(tabId As String)
    On Error Resume Next
    If IsRibbonUIAvailable() Then
        mRibbonUI.ActivateTab tabId
    End If
End Sub

'@Description: Verifica si el objeto ribbonUI esta disponible y funcional
Private Function IsRibbonUIAvailable() As Boolean
    If mRibbonUI Is Nothing Then IsRibbonUIAvailable = False: Exit Function
    
    On Error Resume Next
    Dim testType As String
    testType = TypeName(mRibbonUI)
    If Err.Number <> 0 Then
        IsRibbonUIAvailable = False
        Err.Clear
    ElseIf testType = "Nothing" Or testType = "Empty" Then
        IsRibbonUIAvailable = False
    Else
        IsRibbonUIAvailable = True
    End If
End Function

'@Description: Intenta recuperar el Ribbon automaticamente
Private Function TryAutoRecover() As Boolean
    On Error GoTo ErrHandler

    mIsRecovering = True
    LogInfo MODULE_NAME, "[TryAutoRecover] - Iniciando"

    ' Usar el modulo de recuperacion
    TryAutoRecover = TryRecoverRibbon()

    mIsRecovering = False
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "[TryAutoRecover] - Error", Err.Number, Err.Description
    mIsRecovering = False
    TryAutoRecover = False
End Function


'@Description: Indica si el Ribbon esta en proceso de recuperacion
Public Property Get IsRecovering() As Boolean
    IsRecovering = mIsRecovering
End Property


'@Description: Diagnostico rapido del estado del Ribbon
Public Function GetQuickDiagnostics() As String
    Dim info As String
    info = "RibbonUI: "

    If Not IsRibbonUIAvailable() Then
        info = info & "Nothing"
    Else
        On Error Resume Next
        info = info & TypeName(mRibbonUI)
        If Err.Number <> 0 Then
            info = info & "Corrupted"
            Err.Clear
        End If
        On Error GoTo 0
    End If

    info = info & " | WasInit: " & mWasEverInitialized
    info = info & " | Recovering: " & mIsRecovering

    GetQuickDiagnostics = info
End Function

'=============================================================================
' ACCIONES DEL RIBBON - Implementacion de logica de negocio
'=============================================================================

Public Sub GenerarGraficosDesdeCurvasRto()
    On Error GoTo ErrHandler
    
    LogInfo MODULE_NAME, "[GenerarGraficosDesdeCurvasRto] Iniciando..."
    
    ' TODO: Implementar logica de generacion de graficos
    ' Cuando exista el servicio IChartGenerator:
    ' Dim chartGen As Object
    ' Set chartGen = App.Services.GetService("IChartGenerator")
    ' chartGen.GenerarGraficoSensibilidad
    
    MsgBox "Generar graficos - TODO: Implementar", vbInformation
    
    Exit Sub
    
ErrHandler:
    LogError MODULE_NAME, "[GenerarGraficosDesdeCurvasRto] Error", Err.Number, Err.Description
    MsgBox "Error al generar graficos: " & Err.Description, vbCritical
End Sub

Public Sub InvertirEjes()
    On Error GoTo ErrHandler
    
    LogInfo MODULE_NAME, "[InvertirEjes] Iniciando..."
    
    ' Llamar a la funcion existente
    Call InvertirEjesGraficoActivo
    
    Exit Sub
    
ErrHandler:
    LogError MODULE_NAME, "[InvertirEjes] Error", Err.Number, Err.Description
    MsgBox "Error al invertir ejes: " & Err.Description, vbCritical
End Sub

Public Sub FormatearCGASING()
    On Error GoTo ErrHandler
    
    LogInfo MODULE_NAME, "[FormatearCGASING] Iniciando..."
    
    ' TODO: Implementar con IFormatter cuando exista
    ' Dim formatter As IFormatter
    ' Set formatter = App.Services.GetService("ICGASINGFormatter")
    ' If formatter.CanFormat(ActiveWorkbook) Then
    '     formatter.Format ActiveWorkbook
    ' End If
    
    ' Por ahora llamar a la funcion existente
    Call FormatCGASINGSheet
    
    Exit Sub
    
ErrHandler:
    LogError MODULE_NAME, "[FormatearCGASING] Error", Err.Number, Err.Description
    MsgBox "Error al formatear CGASING: " & Err.Description, vbCritical
End Sub

Public Sub MostrarConfigurador()
    On Error GoTo ErrHandler
    
    LogInfo MODULE_NAME, "[MostrarConfigurador] Mostrando formulario..."
    
    ' Mostrar formulario de configuracion
    frmConfiguracion.Show
    
    Exit Sub
    
ErrHandler:
    LogError MODULE_NAME, "[MostrarConfigurador] Error", Err.Number, Err.Description
    MsgBox "Error al mostrar configurador: " & Err.Description, vbCritical
End Sub

Public Sub ReplaceWithNamesInValidations()
    On Error GoTo ErrHandler
    
    LogInfo MODULE_NAME, "[ReplaceWithNamesInValidations] Iniciando..."
    
    ' TODO: Implementar logica
    MsgBox "Replace with names - TODO: Implementar", vbInformation
    
    Exit Sub
    
ErrHandler:
    LogError MODULE_NAME, "[ReplaceWithNamesInValidations] Error", Err.Number, Err.Description
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Public Sub CrearBackupCodigoVBA()
    On Error GoTo ErrHandler
    
    LogInfo MODULE_NAME, "[CrearBackupCodigoVBA] Iniciando..."
    
    ' TODO: Implementar backup de codigo VBA
    ' Exportar todos los modulos a una carpeta con timestamp
    
    MsgBox "Backup VBA - TODO: Implementar", vbInformation
    
    Exit Sub
    
ErrHandler:
    LogError MODULE_NAME, "[CrearBackupCodigoVBA] Error", Err.Number, Err.Description
    MsgBox "Error al crear backup: " & Err.Description, vbCritical
End Sub

'=============================================================================
' CICLO DE VIDA
'=============================================================================

Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
    mWasEverInitialized = False
    mIsRecovering = False
End Sub

Private Sub Class_Terminate()
    Call IService_Dispose
End Sub

