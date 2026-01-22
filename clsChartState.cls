VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChartState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("2-Infrastructure.State")
Option Explicit

' ===============================================================================
' CLASE: clsChartState
' ===============================================================================
' PROPOSITO:
'   Almacena estado actual de gráficos en Excel:
'     - Gráfico activo (ActiveChart)
'     - Si se puede invertir ejes (CanInvertAxes)
'     - Hoja con gráficos activa
'
'   Provee acceso mediante properties. NO dispara eventos.
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   USADO POR:
'     - clsChartEventManager (escribe estado)
'     - clsRibbonManager (lee estado para habilitar botones)
'     - modRibbonCallbacks (lee estado para validar)
'
'   CONTENIDO EN:
'     - clsApplicationState
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - PULL: Properties Get/Set para acceso directo
'   - Sin eventos (eliminamos eventos que solo actualizaban flags)
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' -------------------------------------------------------------------------------
' Archivo origen: clsChartState.cls (líneas 1-12) + disperso en clsAplicacion
'
' F-001: Estado de gráfico activo
'   - Legacy: clsAplicacion.bChartActive (línea ~400)
'   - Migra: Property ActiveChart As Excel.Chart
'
' F-002: Estado de invertibilidad
'   - Legacy: clsAplicacion.bCanInvertAxes (línea ~401)
'   - Migra: Property CanInvertAxes As Boolean
'
' F-003: Hoja con gráficos
'   - Legacy: clsChartEventsManager.mWatchingSheet
'   - Migra: Property ActiveSheetWithCharts As Worksheet
'
' ELIMINACIONES:
'   - ❌ Eventos ChartActivated/Deactivated (reemplazados por property setters)
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 2: Property ActiveChart As Excel.Chart
' TODO Sprint 2: Property CanInvertAxes As Boolean
' TODO Sprint 2: Property ActiveSheetWithCharts As Worksheet
