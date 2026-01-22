VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsApplicationState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("2-Infrastructure.State")
Option Explicit

' ===============================================================================
' CLASE: clsApplicationState
' ===============================================================================
' PROPOSITO:
'   Contenedor centralizado de ESTADO GLOBAL de la aplicación.
'   Consolida los estados de ribbon, chart, file y execution que en legacy
'   estaban dispersos en múltiples clases.
'
'   Provee acceso compartido a estado mediante properties (PULL).
'   NO contiene lógica, solo almacenamiento y acceso.
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   CONTIENE:
'     - clsRibbonState (estado del ribbon)
'     - clsChartState (estado de gráficos)
'     - clsFileState (estado de archivos)
'     - clsApplicationConfiguration (configuración)
'
'   USADO POR:
'     - TODOS los servicios (lectura/escritura de estado)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - PULL: Servicios leen/escriben estado mediante properties
'   - Sin eventos (patrón pull puro)
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' -------------------------------------------------------------------------------
' NO EXISTE EXACTAMENTE EN MAIN
'
' Consolida estado disperso en:
'   - clsAplicacion.bChartActive línea ~400
'   - clsAplicacion.bCanInvertAxes línea ~401
'   - clsRibbonState (ya existía)
'   - clsFileManager.p_currentFile
'   - Variables globales sin encapsular
'
' BENEFICIO: Estado centralizado, fácil de depurar y testear
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 1: Properties para acceder a cada sub-estado
' TODO Sprint 1: Property Get RibbonState() As clsRibbonState
' TODO Sprint 1: Property Get ChartState() As clsChartState
' TODO Sprint 1: Property Get FileState() As clsFileState
' TODO Sprint 1: Property Get Configuration() As clsApplicationConfiguration
