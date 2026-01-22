VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRibbonManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("3-Services.RibbonManagement")
Option Explicit

' ===============================================================================
' CLASE: clsRibbonManager
' ===============================================================================
' PROPOSITO:
'   Servicio que gestiona el Ribbon de Excel:
'     - Almacena puntero IRibbonUI
'     - Invalida controles cuando cambia estado
'     - Recupera Ribbon perdido automáticamente
'     - Evalúa estado enabled de botones
'
'   NO dispara eventos custom (reemplazados por llamadas directas).
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   USA:
'     - IRibbonUI (COM de Excel)
'     - clsApplicationState.RibbonState (lee estado)
'     - clsApplicationState.ChartState (lee estado para botones)
'     - clsApplicationState.FileState (lee estado para validaciones)
'
'   USADO POR:
'     - clsApplication (llama InvalidateRibbon)
'     - clsChartEventManager (llama RefreshChartButtons)
'     - modRibbonCallbacks (llama GetControlEnabled)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - PULL: Lee estado de clsApplicationState
'   - PUSH: Llama IRibbonUI.Invalidate/InvalidateControl
'   - NO eventos custom (eliminados)
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' ===============================================================================
' Archivos origen: clsRibbonEvents.cls + clsRibbonUI.cls (fusionados)
'
' F-001: Almacenar puntero IRibbonUI
'   - Legacy (clsRibbonEvents): Private mRibbonUI As IRibbonUI (línea 18)
'   - Legacy (clsRibbonEvents): Init(ribbonObj) (línea 56)
'   - Migra: Public Sub Initialize(ribbonUI As IRibbonUI)
'
' F-002: Invalidar Ribbon completo
'   - Legacy (clsRibbonEvents): InvalidarRibbon() (línea 156)
'   - Migra: Public Sub InvalidateRibbon()
'   - Llama _ribbonUI.Invalidate()
'
' F-003: Invalidar control específico
'   - Legacy (clsRibbonEvents): InvalidarControl(idControl) (línea 189)
'   - Migra: Public Sub InvalidateControl(controlId As String)
'   - Llama _ribbonUI.InvalidateControl(controlId)
'
' F-004: Evaluar estado enabled de controles
'   - Legacy (clsRibbonEvents): GetRibbonControlEnabled(control) (línea 119)
'   - Migra: Public Function GetControlEnabled(controlId As String) As Boolean
'   - Consulta estado en clsApplicationState (NO lógica inline)
'
' F-005: Activar tab del Ribbon
'   - Legacy (clsRibbonEvents): ActivarTab(tabId) (línea 220)
'   - Migra: Public Sub ActivateTab(tabId As String)
'   - Llama _ribbonUI.ActivateTab(tabId)
'
' F-006: Recuperación automática de Ribbon perdido
'   - Legacy (clsRibbonEvents): TryAutoRecover() (línea ~250)
'   - Migra: Public Function TryRecover() As Boolean
'   - Llama modMacroAppLifecycle.TryRecoverRibbon()
'
' F-007: Validación de disponibilidad
'   - Legacy (clsRibbonEvents): IsRibbonUIAvailable() (línea ~230)
'   - Migra: Public Function IsAvailable() As Boolean
'   - Verifica que _ribbonUI no sea Nothing y sea válido
'
' F-008: Detener eventos (liberación)
'   - Legacy (clsRibbonEvents): StopEvents() (línea 78)
'   - Migra: Public Sub Dispose()
'   - Set _ribbonUI = Nothing
'
' FUSIONA funcionalidad de:
'   - clsRibbonEvents (eventos y puntero IRibbonUI)
'   - clsRibbonUI (si existía separado en algún branch)
'
' ELIMINACIONES:
'   - ❌ 6 eventos custom (GenerarGraficosDesdeCurvasRto, InvertirEjes, etc)
'   - ❌ 6 métodos On* que solo hacían RaiseEvent
'   - Reemplazados por: llamadas directas desde modRibbonCallbacks a clsApplication
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 2: Private _ribbonUI As IRibbonUI
' TODO Sprint 2: Private _appState As clsApplicationState
' TODO Sprint 2: Public Sub Initialize(ribbonUI As IRibbonUI, appState)
' TODO Sprint 2: Public Sub InvalidateRibbon()
' TODO Sprint 2: Public Sub InvalidateControl(controlId As String)
' TODO Sprint 2: Public Function GetControlEnabled(controlId As String) As Boolean
' TODO Sprint 2: Public Sub ActivateTab(tabId As String)
' TODO Sprint 3: Public Function TryRecover() As Boolean
' TODO Sprint 3: Public Function IsAvailable() As Boolean
' TODO Sprint 3: Public Sub Dispose()
