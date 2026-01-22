VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChartEventAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("2-Infrastructure.ExcelIntegration")
Option Explicit

' ===============================================================================
' CLASE: clsChartEventAdapter
' ===============================================================================
' PROPOSITO:
'   Adaptador de eventos COM de Excel.Chart individual.
'   Captura eventos de UN gráfico (Activate, Deactivate, MouseDown, etc)
'   y ejecuta callbacks directos en clsChartEventManager.
'
'   Instanciado dinámicamente por clsChartEventManager (uno por gráfico).
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   ESCUCHA (WithEvents):
'     - Excel.Chart (eventos COM de UN gráfico)
'
'   LLAMA (callbacks directos):
'     - clsChartEventManager.OnChartActivated(chart)
'     - clsChartEventManager.OnChartDeactivated(chart)
'
'   CREADO POR:
'     - clsChartEventManager (crea una instancia por gráfico en hoja activa)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - EVENTS: WithEvents Chart (único permitido)
'   - CALLBACKS: Llamadas directas a clsChartEventManager
'   - NO re-emite RaiseEvent
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' -------------------------------------------------------------------------------
' Archivo origen: clsChartEvents.cls (líneas 1-97)
'
' F-001: Wrapper de eventos Chart
'   - Legacy: Private WithEvents mChart As Chart (línea ~10)
'   - Migra: Mantiene WithEvents (necesario para COM)
'
' F-002: Eventos capturados
'   - Legacy: mChart_Activate (línea 23)
'   - Legacy: mChart_Deactivate (línea 31)
'   - CAMBIA: RaiseEvent → callback directo
'
' F-003: Asignación de gráfico a monitorear
'   - Legacy: Property Set Chart(cht As Chart) (línea ~50)
'   - Migra: Property Set MonitoredChart(cht As Chart)
'
' ELIMINACIONES:
'   - ❌ Public Event ChartActivated (línea ~12)
'   - ❌ Public Event ChartDeactivated (línea ~13)
'   - Reemplazados por: callbacks directos
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 2: Private WithEvents _chart As Excel.Chart
' TODO Sprint 2: Private _callbackHandler As clsChartEventManager
' TODO Sprint 2: Public Sub Initialize(handler, chart)
' TODO Sprint 2: Implementar _chart_Activate con callback directo
' TODO Sprint 2: Implementar _chart_Deactivate con callback directo
