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
' CLASE: clsApplication
' ===============================================================================
' PROPOSITO:
'   Composition Root de la aplicación. Crea, configura y conecta TODOS los
'   componentes del sistema en el orden correcto. Es el único punto de entrada
'   para la inicialización y terminación de la aplicación.
'
'   NO contiene lógica de negocio, solo orchestración de ciclo de vida.
'   NO escucha eventos de servicios (elimina los 6 WithEvents de legacy).
'   SÍ recibe callbacks directos desde adaptadores de infraestructura.
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   CREA (composición):
'     - clsServiceContainer (DI container)
'     - clsApplicationState (estado global)
'     - clsApplicationConfiguration (configuración)
'     - clsExcelExecutionContext (adaptador Excel)
'     - clsChartEventAdapter (adaptador Chart)
'     - clsFileSystemMonitor (monitor FS)
'     - clsRibbonManager (gestor ribbon)
'     - clsChartEventManager (gestor eventos chart)
'     - clsFileManager (gestor archivos)
'     - clsOpportunityManager (gestor oportunidades)
'
'   RECIBE (callbacks directos desde adaptadores):
'     - clsExcelExecutionContext.OnWorkbookOpened(wb)
'     - clsExcelExecutionContext.OnSheetActivated(sh)
'     - clsFileSystemMonitor eventos de dominio (WithEvents - único permitido)
'
'   LLAMA (métodos públicos):
'     - Todos los servicios mediante llamadas directas (NO eventos)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - PULL: Lee estado de clsApplicationState cuando necesita
'   - PUSH: Llama directamente a métodos públicos de servicios
'   - EVENTS: Solo escucha eventos de DOMINIO de clsFileSystemMonitor
'   - CALLBACKS: Recibe callbacks directos desde adaptadores de infraestructura
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' -------------------------------------------------------------------------------
' Archivo origen: clsAplicacion.cls (líneas 1-479)
'
' F-001: Inicialización de aplicación
'   - Legacy: clsAplicacion.Class_Initialize() líneas 45-122
'   - Migra: Creación de TODOS los servicios en orden correcto
'   - Migra: Suscripción a eventos (reducida de 6 WithEvents a 1)
'   - Migra: Configuración de callbacks directos
'
' F-002: Terminación de aplicación
'   - Legacy: clsAplicacion.Class_Terminate() líneas 123-155
'   - Migra: Liberación de servicios en orden inverso
'   - Migra: Limpieza de estado
'
' F-003: Facade de acceso a servicios
'   - Legacy: Properties Get líneas 156-211
'   - Migra: Propiedades para acceder a servicios principales
'   - ELIMINA: RibbonHandler (se mueve a clsRibbonManager)
'
' F-004: Coordinación de eventos de Excel
'   - Legacy: ctx_WorkbookActivated línea 233
'   - Legacy: ctx_SheetActivated línea 242
'   - Legacy: ctx_SheetDeactivated línea 258
'   - CAMBIA: De eventos WithEvents a callbacks directos
'
' F-005: Coordinación de eventos de FileSystem
'   - Legacy: mFSMonitoringCoord_OpportunityCreated línea 302
'   - Legacy: mFSMonitoringCoord_OpportunityDeleted línea 307
'   - Legacy: mFSMonitoringCoord_TemplateCreated línea 328
'   - MANTIENE: WithEvents (son eventos de DOMINIO 1-a-N)
'
' F-006: Comandos desde Ribbon
'   - Legacy: evRibbon_GenerarGraficosDesdeCurvasRto línea 398
'   - Legacy: evRibbon_InvertirEjes línea 409
'   - Legacy: evRibbon_FormatearCGASING línea 418
'   - Legacy: evRibbon_Configurador línea 433
'   - CAMBIA: De eventos WithEvents a métodos públicos llamados directamente
'
' ELIMINACIONES (antes eran WithEvents, ahora llamadas directas):
'   - ❌ ctx_* (ExecutionContext): Ahora callbacks directos
'   - ❌ mOpportunities_* (OpportunitiesMgr): Ahora llamadas directas
'   - ❌ mChartManager_* (ChartEventsManager): Ahora llamadas directas
'   - ❌ evRibbon_* (RibbonEvents): Ahora métodos públicos
'   - ❌ mRibbonState_StateChanged: Ahora property setter
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 1: Implementar Initialize/Terminate
' TODO Sprint 1: Crear servicios y configurar dependencias
' TODO Sprint 2: Implementar callbacks desde adaptadores
' TODO Sprint 3: Implementar métodos públicos para comandos Ribbon
