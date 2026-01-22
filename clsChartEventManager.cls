VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChartEventManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("3-Services.ChartManagement")
Option Explicit

' ===============================================================================
' CLASE: clsChartEventManager
' ===============================================================================
' PROPOSITO:
'   Servicio que gestiona eventos de gráficos en Excel.
'   Monitorea todos los gráficos (ChartObjects) en la hoja activa y
'   captura eventos de activación/desactivación mediante clsChartEventAdapter.
'
'   Actualiza clsChartState cuando cambia el gráfico activo.
'   Notifica a clsRibbonManager para actualizar botones.
'
'   NO dispara eventos custom (reemplazados por llamadas directas).
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   USA:
'     - clsChartEventAdapter (crea instancias para cada gráfico)
'     - clsApplicationState.ChartState (escribe estado)
'     - clsRibbonManager (llama RefreshChartButtons)
'
'   USADO POR:
'     - clsApplication (llama WatchSheet cuando cambia hoja activa)
'
'   RECIBE callbacks desde:
'     - clsChartEventAdapter.OnChartActivated(chart)
'     - clsChartEventAdapter.OnChartDeactivated(chart)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - PULL: Lee estado de clsApplicationState.ChartState
'   - PUSH: Escribe estado en clsApplicationState.ChartState
'   - PUSH: Llama clsRibbonManager.RefreshChartButtons()
'   - CALLBACKS: Recibe de clsChartEventAdapter
'   - NO eventos custom (eliminados)
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' ===============================================================================
' Archivo origen: clsChartEventsManager.cls (líneas 1-148)
'
' F-001: Monitoreo de hoja activa
'   - Legacy: WatchSheet(sh As Worksheet) (línea ~30)
'   - Migra: Public Sub WatchSheet(sh As Worksheet)
'   - Crea clsChartEventAdapter para cada ChartObject en hoja
'
' F-002: Detención de monitoreo
'   - Legacy: StopWatching() (línea ~60)
'   - Migra: Public Sub StopWatching()
'   - Libera todos los adaptadores de gráficos
'
' F-003: Callback de activación de gráfico
'   - Legacy: mChart_Activated (línea ~80) → RaiseEvent ChartActivated
'   - CAMBIA: Public Sub OnChartActivated(chart As Excel.Chart)
'   - Actualiza clsChartState.ActiveChart
'   - Llama clsRibbonManager.RefreshChartButtons()
'
' F-004: Callback de desactivación de gráfico
'   - Legacy: mChart_Deactivated (línea ~100) → RaiseEvent ChartDeactivated
'   - CAMBIA: Public Sub OnChartDeactivated(chart As Excel.Chart)
'   - Actualiza clsChartState.ActiveChart = Nothing
'   - Llama clsRibbonManager.RefreshChartButtons()
'
' F-005: Colección de adaptadores activos
'   - Legacy: Private mActiveCharts As Collection (línea ~15)
'   - Migra: Private _chartAdapters As Collection
'
' F-006: Referencia a hoja monitoreada
'   - Legacy: Private mWatchingSheet As Worksheet (línea ~16)
'   - Migra: Private _watchedSheet As Worksheet
'
' ELIMINACIONES:
'   - ❌ Public Event ChartActivated (línea ~12)
'   - ❌ Public Event ChartDeactivated (línea ~13)
'   - ❌ Public Event HojaConGraficosCambiada (línea ~14)
'   - Reemplazados por: métodos públicos + llamadas directas
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 2: Private _chartAdapters As Collection
' TODO Sprint 2: Private _watchedSheet As Worksheet
' TODO Sprint 2: Private _appState As clsApplicationState
' TODO Sprint 2: Private _ribbonManager As clsRibbonManager
' TODO Sprint 2: Public Sub WatchSheet(sh As Worksheet)
' TODO Sprint 2: Public Sub StopWatching()
' TODO Sprint 2: Public Sub OnChartActivated(chart As Excel.Chart)
' TODO Sprint 2: Public Sub OnChartDeactivated(chart As Excel.Chart)
