

# Análisis Detallado de la Carpeta `main-mirror`

## Descripción General

Esta carpeta contiene una versión del complemento XLAM para Excel con una arquitectura más tradicional y monolítica. A continuación se detalla el análisis de cada componente según la plantilla de análisis proporcionada.

## Sección 1: Inventario de Componentes

### 1.1. Clases (.cls)

#### 📦 clsAplicacion

**Ubicación:** `clsAplicacion.cls` (líneas 1-456)

**Propósito:**
Coordinador principal de la aplicación. Crea todos los servicios, suscribe eventos, y expone facade de acceso.

**Responsabilidades:**
- Creación e inicialización de todos los servicios
- Suscripción centralizada a eventos (WithEvents)
- Exposición de facade para acceso a servicios
- Coordinación de respuestas a eventos
- Gestión del ciclo de vida de la aplicación

**Atributos privados:**

```vba
Private m_bChartActive As Boolean
Private mOpportunities As clsOpportunitiesMgr
Private mChartManager As clsChartEventsManager
Private mFSMonitoringCoord As clsFSMonitoringCoord
Private mRibbonState As clsRibbonState
Private evRibbon As clsRibbonEvents
Private mConfiguration As clsConfiguration
Private mFileMgr As clsFileManager
Private ctx As clsExecutionContext

```

**WithEvents declarados:**



```
Private WithEvents mOpportunities As clsOpportunitiesMgr
Private WithEvents mChartManager As clsChartEventsManager
Private WithEvents mFSMonitoringCoord As clsFSMonitoringCoord
Private WithEvents mRibbonState As clsRibbonState
Private WithEvents evRibbon As clsRibbonEvents
Private WithEvents ctx As clsExecutionContext

```

**Eventos que escucha:**




| Emisor | Evento | Manejador | Línea |
| --- | --- | --- | --- |
| clsExecutionContext | WorkbookActivated | ctx\_WorkbookActivated | 156 |
| clsExecutionContext | SheetActivated | ctx\_SheetActivated | 167 |
| clsExecutionContext | SheetDeactivated | ctx\_SheetDeactivated | 178 |
| clsOpportunitiesMgr | OpportunityCollectionUpdate | mOpportunities\_OpportunityCollectionUpdate | 201 |
| clsOpportunitiesMgr | currOpportunityChanged | mOpportunities\_currOpportunityChanged | 212 |
| clsChartEventsManager | ChartActivated | mChartManager\_ChartActivated | 256 |
| clsChartEventsManager | ChartDeactivated | mChartManager\_ChartDeactivated | 267 |
| clsFSMonitoringCoord | OpportunityCreated | mFSMonitoringCoord\_OpportunityCreated | 278 |
| clsFSMonitoringCoord | OpportunityDeleted | mFSMonitoringCoord\_OpportunityDeleted | 289 |
| clsFSMonitoringCoord | OpportunityRenamed | mFSMonitoringCoord\_OpportunityRenamed | 295 |
| clsFSMonitoringCoord | OpportunityItemDeleted | mFSMonitoringCoord\_OpportunityItemDeleted | 301 |
| clsFSMonitoringCoord | OpportunityItemRenamed | mFSMonitoringCoord\_OpportunityItemRenamed | 307 |
| clsFSMonitoringCoord | TemplateCreated | mFSMonitoringCoord\_TemplateCreated | 313 |
| clsFSMonitoringCoord | TemplateChanged | mFSMonitoringCoord\_TemplateChanged | 319 |
| clsFSMonitoringCoord | GasFileCreated | mFSMonitoringCoord\_GasFileCreated | 325 |
| clsFSMonitoringCoord | GasFileChanged | mFSMonitoringCoord\_GasFileChanged | 331 |
| clsFSMonitoringCoord | MonitoringError | mFSMonitoringCoord\_MonitoringError | 337 |
| clsFSMonitoringCoord | MonitoringReconnected | mFSMonitoringCoord\_MonitoringReconnected | 343 |
| clsFSMonitoringCoord | MonitoringFailed | mFSMonitoringCoord\_MonitoringFailed | 349 |
| clsRibbonEvents | GenerarGraficosDesdeCurvasRto | evRibbon\_GenerarGraficosDesdeCurvasRto | 365 |
| clsRibbonEvents | InvertirEjes | evRibbon\_InvertirEjes | 371 |
| clsRibbonEvents | FormatearCGASING | evRibbon\_FormatearCGASING | 377 |
| clsRibbonEvents | Configurador | evRibbon\_Configurador | 383 |
| clsRibbonEvents | NuevaOportunidad | evRibbon\_NuevaOportunidad | 389 |
| clsRibbonEvents | ReplaceWithNamesInValidations | evRibbon\_ReplaceWithNamesInValidations | 395 |
| clsRibbonState | StateChanged | mRibbonState\_StateChanged | 145 |


**Eventos que dispara:**  

Ninguno (es consumidor final)


**Métodos públicos:**



```
Public Property Get bChartActive() As Boolean                           ' Línea 45
Public Property Get Configuration() As clsConfiguration                 ' Línea 51
Public Property Get executionContext() As clsExecutionContext           ' Línea 57
Public Property Get FileMgr() As clsFileManager                        ' Línea 63
Public Property Get OpportunitiesMgr() As clsOpportunitiesMgr          ' Línea 69
Public Property Get ChartManager() As clsChartEventsManager            ' Línea 75
Public Property Get Ribbon() As clsRibbonEvents                        ' Línea 81
Public Property Get RibbonState() As clsRibbonState                    ' Línea 87
Public Property Let RibbonHandler(xlRibbon As IRibbonUI)               ' Línea 93
Public Sub Initialize()                                                 ' Línea 108
Public Sub Terminate()                                                  ' Línea 135
Public Sub InitFileManager()                                            ' Línea 189
Public Sub ToggleRibbonMode()                                           ' Línea 225
Public Property Get bCanInvertAxes() As Boolean                        ' Línea 425

```

**Métodos privados:**  

20+ métodos privados (líneas 108-450)


**Dependencias:**



```
graph LR
	clsAplicacion --> clsConfiguration
	clsAplicacion --> clsExecutionContext
	clsAplicacion --> clsFileManager
	clsAplicacion --> clsOpportunitiesMgr
	clsAplicacion --> clsChartEventsManager
	clsAplicacion --> clsFSMonitoringCoord
	clsAplicacion --> clsRibbonEvents
	clsAplicacion --> clsRibbonState

```

**Líneas de código:** 456


**Complejidad estimada:** Alta (God Object - múltiples responsabilidades)


#### 📦 clsExecutionContext


**Ubicación:** `clsExecutionContext.cls` (líneas 1-250)


**Propósito:**  

Provee acceso seguro y explícito al contexto de ejecución (Workbook, Worksheet, Chart, Selection). Centraliza la suscripción a eventos de Application y propaga cambios a otras clases.


**Responsabilidades:**


* Suscripción a eventos de Application
* Proporcionar acceso seguro al contexto actual
* Cachear referencias para evitar problemas de puntero
* Propagar eventos a otros componentes


**Atributos privados:**



```
Private m_xlApp As Application
Private m_lastWorkbookObjKey As Double
Private m_lastWorksheetObjKey As Double
Private m_lastChartObjKey As Double
Private m_lastSelectionObjKey As Double
Private m_cachedChartInfo As T_CachedChartInfo

```

**WithEvents declarados:**



```
Private WithEvents m_xlApp As Application

```

**Eventos que escucha:**




| Emisor | Evento | Manejador | Línea |
| --- | --- | --- | --- |
| Application | WorkbookOpen | m\_xlApp\_WorkbookOpen | 65 |
| Application | WorkbookActivate | m\_xlApp\_WorkbookActivate | 71 |
| Application | WorkbookBeforeClose | m\_xlApp\_WorkbookBeforeClose | 77 |
| Application | SheetActivate | m\_xlApp\_SheetActivate | 83 |
| Application | SheetDeactivate | m\_xlApp\_SheetDeactivate | 92 |


**Eventos que dispara:**


* `WorkbookOpened`
* `WorkbookActivated`
* `WorkbookBeforeClose`
* `WorksheetActivated`
* `WorksheetDeactivated`
* `SheetActivated`
* `SheetDeactivated`
* `SelectionChanged`


**Métodos públicos:**



```
Public Sub Initialize()                                    ' Línea 55
Public Property Get Workbook() As Workbook                ' Línea 79
Public Property Get Worksheet() As Worksheet              ' Línea 95
Public Property Get Selection() As Object                 ' Línea 109
Public Property Get Application() As Application          ' Línea 118
Public Property Get Chart() As Chart                      ' Línea 125
Public Property Get HasWorkbook() As Boolean              ' Línea 200
Public Property Get HasWorksheet() As Boolean             ' Línea 205
Public Property Get HasSelection() As Boolean             ' Línea 210
Public Property Get HasChart() As Boolean                 ' Línea 215
Public Function GetSelectedRange() As Range               ' Línea 220
Public Function Diagnostics() As String                   ' Línea 230

```

**Dependencias:**



```
graph LR
	clsExecutionContext --> Application

```

**Líneas de código:** 250


**Complejidad estimada:** Media


#### 📦 clsConfiguration


**Ubicación:** `clsConfiguration.cls` (líneas 1-200)


**Propósito:**  

Gestión de la configuración de la aplicación, almacenando rutas y parámetros en el registro de Windows.


**Responsabilidades:**


* Almacenar y recuperar configuración del registro
* Proporcionar rutas de carpetas configurables
* Mantener parámetros de configuración


**Atributos privados:**



```
Private mRutaOportunidades As String
Private mRutaPlantillas As String
Private mRutaOfergas As String
Private mRutaGasVBNet As String
Private mRutaExcelCalcTempl As String
Private mArrComprImgs As Variant
Private mArrComprDrawPIDs As Variant
Private mSAM As Integer

```

**WithEvents declarados:**  

Ninguno


**Eventos que escucha:**  

Ninguno


**Eventos que dispara:**  

Ninguno (es consumidor final)


**Métodos públicos:**



```
Public Property Get RutaOportunidades() As String         ' Línea 15
Public Property Let RutaOportunidades(newRuta As String) ' Línea 19
Public Property Get RutaPlantillas() As String            ' Línea 24
Public Property Let RutaPlantillas(newRuta As String)     ' Línea 28
Public Property Get RutaOfergas() As String               ' Línea 33
Public Property Let RutaOfergas(newRuta As String)        ' Línea 37
Public Property Get RutaGasVBNet() As String              ' Línea 42
Public Property Let RutaGasVBNet(newRuta As String)       ' Línea 46
Public Property Get RutaExcelCalcTempl() As String        ' Línea 51
Public Property Let RutaExcelCalcTempl(newRuta As String) ' Línea 55
Public Property Get ListComprImgs() As Variant            ' Línea 60
Public Property Let ListComprImgs(arrRutas As Variant)    ' Línea 64
Public Property Get ListComprDrawPIDs() As Variant        ' Línea 69
Public Property Let ListComprDrawPIDs(arrRutas As Variant)' Línea 73
Public Property Get SAM() As Integer                      ' Línea 78
Public Property Let SAM(newSAM As Integer)                ' Línea 82
Public Property Get oDicFoldersToWatch() As Object        ' Línea 95

```

**Dependencias:**



```
graph LR
	clsConfiguration --> WScript.Shell
	clsConfiguration --> scripting.dictionary

```

**Líneas de código:** 200


**Complejidad estimada:** Media


#### 📦 clsFileManager


**Ubicación:** `clsFileManager.cls` (líneas 1-350)


**Propósito:**  

Gestor genérico de archivos que supervisa cualquier tipo de archivo relacionado con la gestión de la aplicación.


**Responsabilidades:**


* Supervisar archivos Excel, PDF, Word, etc.
* Mantener un índice de archivos supervisados
* Mantener sincronizado el archivo de Excel activo
* Proveer análisis de archivos sin duplicar lógica


**Atributos privados:**



```
Private p_trackedFiles As Object
Private p_currExcelFile As clsExcelFile
Private ctx As clsExecutionContext

```

**WithEvents declarados:**



```
Private WithEvents ctx As clsExecutionContext

```

**Eventos que escucha:**




| Emisor | Evento | Manejador | Línea |
| --- | --- | --- | --- |
| clsExecutionContext | WorkbookActivated | ctx\_WorkbookActivated | 285 |
| clsExecutionContext | WorkbookOpened | ctx\_WorkbookOpened | 295 |
| clsExecutionContext | WorkbookBeforeClose | ctx\_WorkbookBeforeClose | 305 |


**Eventos que dispara:**  

Ninguno (es consumidor final)


**Métodos públicos:**



```
Public Sub Initialize(ByVal executionContext As clsExecutionContext) ' Línea 45
Public Property Get ActiveWb() As clsExcelFile                      ' Línea 65
Friend Property Set ActiveWb(f As clsExcelFile)                     ' Línea 76
Public Property Get TrackedCount() As Long                          ' Línea 95
Public Function GetOrTrackWorkbook(wb As Workbook) As clsExcelFile  ' Línea 105
Public Sub UntrackWorkbook(wb As Workbook)                          ' Línea 125
Public Sub TrackFile(f As Object)                                   ' Línea 145
Public Sub UntrackFile(f As Object)                                 ' Línea 165
Public Function AnalizarArchivo(fich As Object) As T_InfoArchivo    ' Línea 185
Public Function AnalizarArchivoActivo() As T_InfoArchivo            ' Línea 215
Public Function GetTrackedFilesInfo() As String                     ' Línea 325

```

**Dependencias:**



```
graph LR
	clsFileManager --> clsExecutionContext
	clsFileManager --> clsExcelFile
	clsFileManager --> IFile

```

**Líneas de código:** 350


**Complejidad estimada:** Media-Alta


#### 📦 clsOpportunitiesMgr


**Ubicación:** `clsOpportunitiesMgr.cls` (líneas 1-300)


**Propósito:**  

Gestiona la lista de “Oportunidades” (subcarpetas) de un directorio base configurado en el sistema.


**Responsabilidades:**


* Refrescar, enumerar y cambiar oportunidad actual
* Disparar eventos para notificar cambio de oportunidad
* Detectar y procesar cambios en carpetas de oportunidades


**Atributos privados:**



```
Private strOportunitiesBaseFolder As String
Private p_ColOpportunities As Collection
Private p_CurrOpportunity As Long
Private p_bEnabled As Boolean
Private ctx As clsExecutionContext

```

**WithEvents declarados:**



```
Private WithEvents ctx As clsExecutionContext

```

**Eventos que escucha:**  

Ninguno (método no implementado)


**Eventos que dispara:**


* `currOpportunityChanged`
* `OpportunityCollectionUpdate`


**Métodos públicos:**



```
Public Sub SetBaseFolder(ByVal ruta As String)                      ' Línea 65
Public Function actualizarColeccionOportunidades()                  ' Línea 85
Public Sub ProcesarCambiosEnOportunidades(ByVal subfolderName As String) ' Línea 145
Public Sub ProcesarCambiosEnItemsOportunidad(ByVal cambios As String) ' Línea 175
Public Function numOpportunities() As Variant                       ' Línea 185
Public Property Get OportunityLabel(Index As Integer) As String     ' Línea 195
Public Property Get OportunityPath(Index As Long) As String         ' Línea 205
Public Property Let CurrOpportunity(Index As Long)                  ' Línea 215
Public Property Get CurrOpportunity() As Long                       ' Línea 225
Public Sub CreaOportunidad()                                        ' Línea 255

```

**Dependencias:**



```
graph LR
	clsOpportunitiesMgr --> clsExecutionContext
	clsOpportunitiesMgr --> Scripting.FileSystemObject
	clsOpportunitiesMgr --> VBScript.RegExp
	clsOpportunitiesMgr --> App.Configuration

```

**Líneas de código:** 300


**Complejidad estimada:** Media


#### 📦 clsChartEventsManager


**Ubicación:** `clsChartEventsManager.cls` (líneas 1-150)


**Propósito:**  

Gestor centralizado de eventos de gráficos (orquestador).


**Responsabilidades:**


* Vigilar gráficos en hojas de Excel
* Notificar activación/desactivación de gráficos
* Coordinar eventos de gráficos


**Atributos privados:**



```
Private mActiveCharts As Collection
Private mWatchingSheet As Object

```

**WithEvents declarados:**  

Ninguno


**Eventos que escucha:**  

Ninguno


**Eventos que dispara:**


* `ChartActivated`
* `ChartDeactivated`
* `HojaConGraficosCambiada`


**Métodos públicos:**



```
Public Sub WatchSheet(sh As Object)                                ' Línea 45
Public Sub StopWatching()                                          ' Línea 85
Public Sub RefreshCurrentSheet()                                   ' Línea 115
Friend Sub NotifyChartActivated(cht As Chart)                     ' Línea 135
Friend Sub NotifyChartDeactivated(cht As Chart)                   ' Línea 140

```

**Dependencias:**



```
graph LR
	clsChartEventsManager --> clsChartEvents
	clsChartEventsManager --> ChartObject

```

**Líneas de código:** 150


**Complejidad estimada:** Media


#### 📦 clsFSMonitoringCoord


**Ubicación:** `clsFSMonitoringCoord.cls` (líneas 1-500)


**Propósito:**  

Coordinador de monitoreo del sistema de archivos.


**Responsabilidades:**


* Configurar y gestionar el monitoreo de carpetas
* Procesar eventos de cambio en el sistema de archivos
* Disparar eventos específicos según tipo de archivo/carpeta


**Atributos privados:**



```
Private mFolderWatcher As clsFSWatcher
Private m_rutaOportunidades As String
Private m_rutaPlantillas As String
Private m_rutaGasVBNet As String

```

**WithEvents declarados:**



```
Private WithEvents mFolderWatcher As clsFSWatcher

```

**Eventos que escucha:**




| Emisor | Evento | Manejador | Línea |
| --- | --- | --- | --- |
| clsFSWatcher | SubfolderCreated | mFolderWatcher\_SubfolderCreated | 150 |
| clsFSWatcher | SubfolderDeleted | mFolderWatcher\_SubfolderDeleted | 160 |
| clsFSWatcher | SubfolderRenamed | mFolderWatcher\_SubfolderRenamed | 170 |
| clsFSWatcher | FileCreated | mFolderWatcher\_FileCreated | 180 |
| clsFSWatcher | FileDeleted | mFolderWatcher\_FileDeleted | 190 |
| clsFSWatcher | FileChanged | mFolderWatcher\_FileChanged | 200 |
| clsFSWatcher | FileRenamed | mFolderWatcher\_FileRenamed | 210 |
| clsFSWatcher | ErrorOccurred | mFolderWatcher\_ErrorOccurred | 230 |
| clsFSWatcher | WatcherReconnected | mFolderWatcher\_WatcherReconnected | 250 |
| clsFSWatcher | WatcherReconnectionFailed | mFolderWatcher\_WatcherReconnectionFailed | 260 |


**Eventos que dispara:**


* `OpportunityCreated`
* `OpportunityDeleted`
* `OpportunityRenamed`
* `OpportunityItemDeleted`
* `OpportunityItemRenamed`
* `TemplateCreated`
* `TemplateChanged`
* `GasFileCreated`
* `GasFileChanged`
* `MonitoringError`
* `MonitoringReconnected`
* `MonitoringFailed`


**Métodos públicos:**



```
Public Property Get FolderWatcher() As clsFSWatcher                 ' Línea 45
Friend Sub IniciarMonitoreo(ByVal oDicFolders As Object)           ' Línea 75
Public Sub ConfigurarMonitoreoOportunidades(ByVal rutaBase As String) ' Línea 350
Public Sub ConfigurarMonitoreoPlantillas(ByVal rutaBase As String) ' Línea 375
Public Sub ConfigurarMonitoreoGasVBNet(ByVal rutaBase As String)   ' Línea 400
Public Sub VerEstadisticasMonitoreo()                             ' Línea 425
Public Sub VerHistorialMonitoreo()                                ' Línea 450
Public Sub LimpiarHistorialMonitoreo()                            ' Línea 475
Public Sub VerConfiguracionWatcher()                              ' Línea 485

```

**Dependencias:**



```
graph LR
	clsFSMonitoringCoord --> clsFSWatcher

```

**Líneas de código:** 500


**Complejidad estimada:** Alta


#### 📦 clsRibbonEvents


**Ubicación:** `clsRibbonEvents.cls` (líneas 1-200)


**Propósito:**  

Gestión de eventos del Ribbon, envuelve el objeto IRibbonUI y gestiona su ciclo de vida con protección y logging.


**Responsabilidades:**


* Gestionar puntero IRibbonUI
* Proporcionar métodos de invalidación segura
* Disparar eventos de acciones del usuario en el Ribbon


**Atributos privados:**



```
Private mribbonUI As IRibbonUI
Private mIsRecovering As Boolean
Private mWasEverInitialized As Boolean

```

**WithEvents declarados:**  

Ninguno


**Eventos que escucha:**  

Ninguno


**Eventos que dispara:**


* `GenerarGraficosDesdeCurvasRto`
* `InvertirEjes`
* `FormatearCGASING`
* `Configurador`
* `NuevaOportunidad`
* `ReplaceWithNamesInValidations`


**Métodos públicos:**



```
Public Property Get ribbonUI() As IRibbonUI                        ' Línea 35
Public Sub Init(ByRef ribbonObj As IRibbonUI)                     ' Línea 55
Public Sub StopEvents()                                           ' Línea 65
Public Sub OnGenerarGraficosDesdeCurvasRto()                      ' Línea 70
Public Sub OnInvertirEjes()                                       ' Línea 75
Public Sub OnFormatearCGASING()                                   ' Línea 80
Public Sub OnConfigurador()                                       ' Línea 85
Public Sub OnNuevaOportunidad()                                   ' Línea 90
Public Sub OnReplaceWithNamesInValidations()                      ' Línea 95
Friend Sub ActivarTab(tabId As String)                            ' Línea 105
Public Function GetRibbonControlEnabled(control As IRibbonControl) As Boolean ' Línea 115
Public Sub InvalidarRibbon()                                      ' Línea 125
Public Sub InvalidarControl(idControl As String)                  ' Línea 155
Public Function GetQuickDiagnostics() As String                   ' Línea 190

```

**Dependencias:**



```
graph LR
	clsRibbonEvents --> IRibbonUI

```

**Líneas de código:** 200


**Complejidad estimada:** Media


#### 📦 clsRibbonState


**Ubicación:** `clsRibbonState.cls` (líneas 1-80)


**Propósito:**  

Representa el estado lógico del Ribbon.


**Responsabilidades:**


* Mantener el modo actual del Ribbon
* Proporcionar métodos para cambiar el estado
* Disparar eventos cuando cambia el estado


**Atributos privados:**



```
Private mModoRibbon As eRibbonMode
Private mVisible As Boolean

```

**WithEvents declarados:**  

Ninguno


**Eventos que escucha:**  

Ninguno


**Eventos que dispara:**


* `StateChanged`


**Métodos públicos:**



```
Public Property Get Modo() As eRibbonMode                         ' Línea 15
Public Property Let Modo(Value As eRibbonMode)                    ' Línea 19
Public Sub ToggleModo()                                           ' Línea 30
Public Function RibbonStateDescription() As String                ' Línea 45
Public Function IsRibbonTabVisible() As Boolean                   ' Línea 65
Public Function IsAdminGroupVisible() As Boolean                  ' Línea 75

```

**Dependencias:**  

Ninguna


**Líneas de código:** 80


**Complejidad estimada:** Baja


### 1.2. Módulos (.bas)


#### 📄 modCALLBACKSRibbon


**Ubicación:** `modCALLBACKSRibbon.bas` (líneas 1-300)


**Propósito:**  

Módulo de integración con la Ribbon que gestiona visibilidad y ejecución de macros para gráficos de sensibilidad.


**Funciones públicas:**



```
Sub RibbonOnLoad(xlRibbon As IRibbonUI)                           ' Línea 15
Sub OnCompararHojas(control As IRibbonControl)                    ' Línea 35
Sub OnDirtyRecalc(control As IRibbonControl)                      ' Línea 40
Sub OnEvalUDFs(control As IRibbonControl)                         ' Línea 45
Public Sub OnChangeAlturaFilas(control As IRibbonControl)         ' Línea 50
Public Sub OnMakeEditableBook(control As IRibbonControl)          ' Línea 58
Public Sub OnFitForPrint(control As IRibbonControl)               ' Línea 63
Public Sub OnVBAExport(control As IRibbonControl)                 ' Línea 68
Public Sub OnVBAImport(control As IRibbonControl)                 ' Línea 73
Public Sub OnOpenLog(control As IRibbonControl)                   ' Línea 78
Public Sub OnVBABackup(control As IRibbonControl)                 ' Línea 83
Public Sub OnProcMetadataSync(control As IRibbonControl)          ' Línea 89
Public Sub OnToggleXLAMVisibility(control As IRibbonControl)      ' Línea 94
Public Sub OnGenerarGraficosDesdeCurvasRto(control As IRibbonControl) ' Línea 107
Public Sub OnInvertirEjes(control As IRibbonControl)              ' Línea 112
Public Sub OnFormatearCGASING(control As IRibbonControl)          ' Línea 117
Public Sub OnNuevaOportunidad(control As IRibbonControl)          ' Línea 122
Public Sub OnReplaceWithNamesInValidations(control As IRibbonControl) ' Línea 127
Sub OnConfigurador(control As IRibbonControl)                     ' Línea 135
Public Sub CallbackRefrescarOportunidades(control As IRibbonControl) ' Línea 148
Sub GetOportunidadesCount(control As IRibbonControl, ByRef returnedVal) ' Línea 162
Sub GetOportunidadesLabel(control As IRibbonControl, Index As Integer, ByRef label) ' Línea 172
Sub OnOportunidadesSeleccionada(control As IRibbonControl, id As String, Index As Integer) ' Línea 182
Sub GetSelectedOportunidadIndex(control As IRibbonControl, ByRef Index) ' Línea 192
Sub GetSupertipRutaBaseOportunidades(control As IRibbonControl, ByRef returnedVal) ' Línea 198
Sub GetSupertipRutaBasePlantillas(control As IRibbonControl, ByRef returnedVal) ' Línea 202
Sub GetSupertipRutaBaseOfergas(control As IRibbonControl, ByRef returnedVal) ' Línea 206
Sub GetSupertipRutaBaseGasVBNet(control As IRibbonControl, ByRef returnedVal) ' Línea 210
Sub GetSupertipRutaBaseCalcTmpl(control As IRibbonControl, ByRef returnedVal) ' Línea 214
Function GetSupertipRutaBase(ByVal strSettingRuta As String)       ' Línea 220
Public Sub GetGraficoEnabled(control As IRibbonControl, ByRef enabled) ' Línea 228
Public Sub GetInvertirEjesEnabled(control As IRibbonControl, ByRef enabled) ' Línea 233
Public Sub GetCGASINGEnabled(control As IRibbonControl, ByRef enabled) ' Línea 238
Public Sub GetNuevaOportunidadEnabled(control As IRibbonControl, ByRef enabled) ' Línea 243
Public Sub GetOfertaFullEnabled(control As IRibbonControl, ByRef enabled) ' Línea 248
Public Sub GetOpenLogEnabled(control As IRibbonControl, ByRef enabled) ' Línea 252
Public Sub GetMenuEnabled(control As IRibbonControl, ByRef enabled) ' Línea 257
Public Sub GetLabelToggleXLAM(control As IRibbonControl, ByRef returnedVal) ' Línea 264
Public Sub GetLabelGrpConfiguracion(control As IRibbonControl, ByRef returnedVal) ' Línea 270
Public Sub GetTabABCVisible(control As IRibbonControl, ByRef Visible) ' Línea 276
Public Sub GetGrpDeveloperAdminVisible(control As IRibbonControl, ByRef Visible) ' Línea 282

```

**Funciones privadas (solo cantidad):**  

1 función privada


**Dependencias / Patrón de delegación:**


* Llama a: `App` → `RibbonHandler`, `Ribbon`, `OpportunitiesMgr`, `Configuration`
* Patrón: Callback XML → Delegación a clsAplicacion


**Líneas de código:** 300


#### 📄 mod\_Logger


**Ubicación:** `mod_Logger.bas` (líneas 1-150)


**Propósito:**  

Módulo de logging centralizado que proporciona funciones de logging consistentes para toda la aplicación.


**Funciones públicas:**



```
Public Sub AbrirLog()                                             ' Línea 25
Public Sub InitLogger(Optional ByVal minLevel As LogLevel = LOG_DEBUG, Optional ByVal logToFile As Boolean = False, Optional ByVal logFilePath As String = "") ' Línea 45
Public Sub LogDebug(ByVal source As String, ByVal message As String) ' Línea 65
Public Sub LogInfo(ByVal source As String, ByVal message As String) ' Línea 72
Public Sub LogWarning(ByVal source As String, ByVal message As String) ' Línea 79
Public Sub LogError(ByVal source As String, ByVal message As String, Optional ByVal errNumber As Long = 0, Optional ByVal errDescription As String = "") ' Línea 86
Public Sub LogCritical(ByVal source As String, ByVal message As String, Optional ByVal errNumber As Long = 0, Optional ByVal errDescription As String = "") ' Línea 100
Public Sub LogCurrentError(ByVal source As String, Optional ByVal additionalInfo As String = "") ' Línea 114
Public Function GetLevelName(ByVal level As LogLevel) As String    ' Línea 135
Public Sub ClearLogFile()                                         ' Línea 145
Public Function GetLogFilePath() As String                        ' Línea 150

```

**Funciones privadas (solo cantidad):**  

4 funciones privadas


**Dependencias / Patrón de delegación:**


* Llama a: `Debug.Print`, `File System`
* Patrón: Singleton de logging con niveles


**Líneas de código:** 150


#### 📄 mod\_ConstantsGlobals


**Ubicación:** `mod_ConstantsGlobals.bas` (líneas 1-200)


**Propósito:**  

Módulo que contiene constantes y tipos globales utilizados en toda la aplicación.


**Funciones públicas:**  

Ninguna


**Funciones privadas (solo cantidad):**  

Ninguna


**Dependencias / Patrón de delegación:**


* Define: `Enums`, `Constants`, `Types`
* Patrón: Contenedor de constantes globales


**Líneas de código:** 200


### 1.3. Formularios (.frm)


No se han encontrado formularios en esta revisión inicial. Se deben revisar los archivos `.frm` y `.frx` para completar esta sección.


### 1.4. Tabla de Eventos (Quién dispara → Quién escucha)




| Clase Emisora | Evento | Clase(s) Receptora(s) | Manejador | Línea |
| --- | --- | --- | --- | --- |
| Application | WorkbookOpen | clsExecutionContext | m\_xlApp\_WorkbookOpen | 65 |
| Application | SheetActivate | clsExecutionContext | m\_xlApp\_SheetActivate | 83 |
| clsExecutionContext | WorkbookActivated | clsFileManager | ctx\_WorkbookActivated | 285 |
| clsExecutionContext | SheetActivated | clsFileManager | ctx\_SheetActivated | 295 |
| clsOpportunitiesMgr | currOpportunityChanged | clsAplicacion | mOpportunities\_currOpportunityChanged | 212 |
| clsChartEventsManager | ChartActivated | clsAplicacion | mChartManager\_ChartActivated | 256 |
| clsRibbonEvents | GenerarGraficosDesdeCurvasRto | clsAplicacion | evRibbon\_GenerarGraficosDesdeCurvasRto | 365 |
| clsFSMonitoringCoord | OpportunityCreated | clsAplicacion | mFSMonitoringCoord\_OpportunityCreated | 278 |
| clsRibbonState | StateChanged | clsAplicacion | mRibbonState\_StateChanged | 145 |


### 1.5. UDFs (User Defined Functions)


No se han identificado UDFs en los módulos revisados en esta sección. Se deben revisar los módulos específicos de UDFs para completar esta sección.


### 1.6. Macros de Excel


Se han identificado varias macros en el módulo `modCALLBACKSRibbon` que son ejecutables desde el Ribbon, botones o atajos de teclado.


## Sección 2: Grafos de Dependencias


### 2.1. DIAGRAMAS ESTRUCTURALES


#### 2.1.1. Diagrama UML de Clases



```
classDiagram
	class clsAplicacion {
		-m_bChartActive Boolean
		-mOpportunities clsOpportunitiesMgr
		-mChartManager clsChartEventsManager
		-mFSMonitoringCoord clsFSMonitoringCoord
		-mRibbonState clsRibbonState
		-evRibbon clsRibbonEvents
		-mConfiguration clsConfiguration
		-mFileMgr clsFileManager
		-ctx clsExecutionContext
		+bChartActive() Boolean
		+Configuration() clsConfiguration
		+executionContext() clsExecutionContext
		+FileMgr() clsFileManager
		+OpportunitiesMgr() clsOpportunitiesMgr
		+ChartManager() clsChartEventsManager
		+Ribbon() clsRibbonEvents
		+RibbonState() clsRibbonState
		+RibbonHandler(xlRibbon As IRibbonUI) 
		+Initialize() 
		+Terminate() 
		+InitFileManager() 
		+ToggleRibbonMode() 
		+bCanInvertAxes() Boolean
	}
	class clsExecutionContext {
		-m_xlApp Application
		-m_lastWorkbookObjKey Double
		-m_lastWorksheetObjKey Double
		-m_lastChartObjKey Double
		-m_lastSelectionObjKey Double
		-m_cachedChartInfo T_CachedChartInfo
		+Initialize() 
		+Workbook() Workbook
		+Worksheet() Worksheet
		+Selection() Object
		+Application() Application
		+Chart() Chart
		+HasWorkbook() Boolean
		+HasWorksheet() Boolean
		+HasSelection() Boolean
		+HasChart() Boolean
		+GetSelectedRange() Range
		+Diagnostics() String
	}
	class clsConfiguration {
		-mRutaOportunidades String
		-mRutaPlantillas String
		-mRutaOfergas String
		-mRutaGasVBNet String
		-mRutaExcelCalcTempl String
		-mArrComprImgs Variant
		-mArrComprDrawPIDs Variant
		-mSAM Integer
		+RutaOportunidades() String
		+RutaOportunidades(newRuta As String) 
		+RutaPlantillas() String
		+RutaPlantillas(newRuta As String) 
		+RutaOfergas() String
		+RutaOfergas(newRuta As String) 
		+RutaGasVBNet() String
		+RutaGasVBNet(newRuta As String) 
		+RutaExcelCalcTempl() String
		+RutaExcelCalcTempl(newRuta As String) 
		+ListComprImgs() Variant
		+ListComprImgs(arrRutas As Variant) 
		+ListComprDrawPIDs() Variant
		+ListComprDrawPIDs(arrRutas As Variant) 
		+SAM() Integer
		+SAM(newSAM As Integer) 
		+oDicFoldersToWatch() Object
	}
	class clsFileManager {
		-p_trackedFiles Object
		-p_currExcelFile clsExcelFile
		-ctx clsExecutionContext
		+Initialize(executionContext As clsExecutionContext) 
		+ActiveWb() clsExcelFile
		+ActiveWb(f As clsExcelFile) 
		+TrackedCount() Long
		+GetOrTrackWorkbook(wb As Workbook) clsExcelFile
		+UntrackWorkbook(wb As Workbook) 
		+TrackFile(f As Object) 
		+UntrackFile(f As Object) 
		+AnalizarArchivo(fich As Object) T_InfoArchivo
		+AnalizarArchivoActivo() T_InfoArchivo
		+GetTrackedFilesInfo() String
	}
	class clsOpportunitiesMgr {
		-strOportunitiesBaseFolder String
		-p_ColOpportunities Collection
		-p_CurrOpportunity Long
		-p_bEnabled Boolean
		-ctx clsExecutionContext
		+SetBaseFolder(ruta As String) 
		+actualizarColeccionOportunidades() 
		+ProcesarCambiosEnOportunidades(subfolderName As String) 
		+ProcesarCambiosEnItemsOportunidad(cambios As String) 
		+numOpportunities() Variant
		+OportunityLabel(Index As Integer) String
		+OportunityPath(Index As Long) String
		+CurrOpportunity(Index As Long) 
		+CurrOpportunity() Long
		+CreaOportunidad() 
	}
	class clsChartEventsManager {
		-mActiveCharts Collection
		-mWatchingSheet Object
		+WatchSheet(sh As Object) 
		+StopWatching() 
		+RefreshCurrentSheet() 
		+NotifyChartActivated(cht As Chart) 
		+NotifyChartDeactivated(cht As Chart) 
	}
	class clsFSMonitoringCoord {
		-mFolderWatcher clsFSWatcher
		-m_rutaOportunidades String
		-m_rutaPlantillas String
		-m_rutaGasVBNet String
		+FolderWatcher() clsFSWatcher
		+IniciarMonitoreo(oDicFolders As Object) 
		+ConfigurarMonitoreoOportunidades(rutaBase As String) 
		+ConfigurarMonitoreoPlantillas(rutaBase As String) 
		+ConfigurarMonitoreoGasVBNet(rutaBase As String) 
		+VerEstadisticasMonitoreo() 
		+VerHistorialMonitoreo() 
		+LimpiarHistorialMonitoreo() 
		+VerConfiguracionWatcher() 
	}
	class clsRibbonEvents {
		-mribbonUI IRibbonUI
		-mIsRecovering Boolean
		-mWasEverInitialized Boolean
		+ribbonUI() IRibbonUI
		+Init(ribbonObj As IRibbonUI) 
		+StopEvents() 
		+OnGenerarGraficosDesdeCurvasRto() 
		+OnInvertirEjes() 
		+OnFormatearCGASING() 
		+OnConfigurador() 
		+OnNuevaOportunidad() 
		+OnReplaceWithNamesInValidations() 
		+ActivarTab(tabId As String) 
		+GetRibbonControlEnabled(control As IRibbonControl) Boolean
		+InvalidarRibbon() 
		+InvalidarControl(idControl As String) 
		+GetQuickDiagnostics() String
	}
	class clsRibbonState {
		-mModoRibbon eRibbonMode
		-mVisible Boolean
		+Modo() eRibbonMode
		+Modo(Value As eRibbonMode) 
		+ToggleModo() 
		+RibbonStateDescription() String
		+IsRibbonTabVisible() Boolean
		+IsAdminGroupVisible() Boolean
	}

	clsAplicacion --> clsConfiguration : usa
	clsAplicacion --> clsExecutionContext : usa
	clsAplicacion --> clsFileManager : usa
	clsAplicacion --> clsOpportunitiesMgr : usa
	clsAplicacion --> clsChartEventsManager : usa
	clsAplicacion --> clsFSMonitoringCoord : usa
	clsAplicacion --> clsRibbonEvents : usa
	clsAplicacion --> clsRibbonState : usa
	clsFileManager --> clsExecutionContext : usa
	clsFileManager --> clsExcelFile : usa
	clsOpportunitiesMgr --> clsExecutionContext : usa
	clsOpportunitiesMgr --> App.Configuration : usa
	clsFSMonitoringCoord --> clsFSWatcher : usa
	clsAplicacion ..> clsExecutionContext : WithEvents
	clsAplicacion ..> clsOpportunitiesMgr : WithEvents
	clsAplicacion ..> clsChartEventsManager : WithEvents
	clsAplicacion ..> clsFSMonitoringCoord : WithEvents
	clsAplicacion ..> clsRibbonEvents : WithEvents
	clsAplicacion ..> clsRibbonState : WithEvents
	clsFileManager ..> clsExecutionContext : WithEvents
	clsExecutionContext ..> Application : WithEvents
	clsFSMonitoringCoord ..> clsFSWatcher : WithEvents

```

#### 2.1.2. Diagrama de Componentes por Nivel



```
graph TD
	subgraph "Nivel 0 - Entry Point"
		TW[ThisWorkbook]
	end

	subgraph "Nivel 1 - Coordinador"
		APP[clsAplicacion<br/>⚠️ God Object<br/>20+ manejadores]
	end

	subgraph "Nivel 2 - Servicios Core"
		CFG[clsConfiguration]
		EXEC[clsExecutionContext<br/>7 eventos]
		FILEMGR[clsFileManager]
	end

	subgraph "Nivel 3 - Servicios Dominio"
		OPP[clsOpportunitiesMgr<br/>2 eventos]
		CHART[clsChartEventsManager<br/>3 eventos]
		FS[clsFSMonitoringCoord<br/>8 eventos]
	end

	subgraph "Nivel 4 - UI"
		RIBBONEV[clsRibbonEvents<br/>6 eventos<br/>⚠️ 2 responsabilidades]
		RIBBONST[clsRibbonState<br/>1 evento]
	end

	subgraph "Nivel 5 - Callbacks"
		CALLBACKS[modCALLBACKSRibbon<br/>12 callbacks]
	end

	TW --> APP
	APP --> CFG
	APP ..> EXEC
	APP --> FILEMGR
	APP ..> OPP
	APP ..> CHART
	APP ..> FS
	APP ..> RIBBONEV
	APP ..> RIBBONST

	RIBBONEV --> RIBBONST

	CALLBACKS --> APP

	style APP fill:#ff6b6b
	style RIBBONEV fill:#ffa500

```

#### 2.1.3. Matriz de Dependencias (Tabla de Acoplamiento)




|  | clsConfig | clsExecCtx | clsFileMgr | clsOppMgr | clsChartMgr | clsFSMon | clsRibbonEv | clsRibbonSt |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| **clsAplicacion** | ✓ | WithEvents | ✓ | WithEvents | WithEvents | WithEvents | WithEvents | WithEvents |
| **clsFileManager** |  | WithEvents |  |  |  |  |  |  |
| **clsRibbonEvents** |  |  |  |  |  |  |  | ✓ |
| **clsOpportunitiesMgr** | ✓ |  | ✓ |  |  |  |  |  |


**Análisis de acoplamiento:**


* ⚠️ **clsAplicacion:** Acoplado a 8 clases (alto acoplamiento aferente - God Object)
* ⚠️ **clsOpportunitiesMgr** → **App.Configuration**: Acoplamiento directo a través de dependencia
* ✅ **clsChartEventsManager**: Bajo acoplamiento (solo 2 dependencias)


### 2.2. DIAGRAMAS DE COMPORTAMIENTO


#### 2.2.1. Diagramas de Secuencia para Análisis de Arquitectura (Flujos Críticos)


**Escenario 1: Diagrama maestro de Inicialización/Carga de la Aplicación**



```
sequenceDiagram
	participant Excel
	participant TW as ThisWorkbook
	participant App as clsAplicacion
	participant Svc as Servicios
	participant Ribbon as clsRibbonEvents

	Excel->>TW: Workbook_Open()
	TW->>App: App.Initialize()
	App->>App: Crear servicios
	loop Para cada servicio
		App->>Svc: New clsServicio()
		App->>Svc: Initialize()
	end
	Note over App: Suscribir WithEvents
	App->>Svc: Set mService = servicio
	Excel->>Ribbon: RibbonOnLoad(ribbon)
	Ribbon->>App: App.RibbonUI.Init(ribbon)
	App-->>TW: Inicialización completa

```

**Escenario 2: Diagrama de Cierre/Gestión de Persistencia**



```
sequenceDiagram
	participant Excel
	participant TW as ThisWorkbook
	participant App as clsAplicacion
	participant Svc as Servicios

	Excel->>TW: Workbook_BeforeClose()
	TW->>App: TerminateApp()
	App->>Svc: Dispose/limpieza
	Note over App: Liberar recursos COM
	App->>Svc: Set objeto = Nothing
	App-->>TW: Limpieza completada
	TW-->>Excel: Continuar cierre

```

**Escenario 3: Control COM supervisor de ficheros detecta cambios en carpeta supervisada**



```
sequenceDiagram
	participant FSWatcher as clsFSWatcher
	participant Coord as clsFSMonitoringCoord
	participant App as clsAplicacion
	participant OppMgr as clsOpportunitiesMgr

	FSWatcher->>Coord: SubfolderCreated(parentFolder, subfolderName)
	Coord->>App: RaiseEvent OpportunityCreated
	App->>OppMgr: ProcesarCambiosEnOportunidades(subfolderName)
	OppMgr->>OppMgr: actualizarColeccionOportunidades()
	App->>Ribbon: InvalidarControl("ddlOportunidades")

```

**Escenario 4: Ejecución de una Macro/Comando Genérico**



```
sequenceDiagram
	participant User as Usuario
	participant XML as Ribbon XML
	participant Callback as modCALLBACKSRibbon
	participant RibbonEv as clsRibbonEvents
	participant App as clsAplicacion
	participant Svc as Servicio

	User->>XML: Click "Generar Gráficos"
	XML->>Callback: OnGenerarGraficos(control)
	Callback->>RibbonEv: OnGenerarGraficosDesdeCurvasRto()
	RibbonEv->>App: Event GenerarGraficosDesdeCurvasRto
	App->>App: evRibbon_GenerarGraficos...()
	App->>Svc: GenerarGraficos()
	Svc-->>User: Gráficos creados

```

#### 2.2.2. Diagrama de Máquina de Estados para componentes de interfaz (Ribbon)


##### 2.2.2.1. Diagrama de Máquina de Estados para el ribbon



```
stateDiagram-v2
	[*] --> OpportunityOnly
	OpportunityOnly --> User : Ctrl+Shift+R
	User --> Admin : Ctrl+Shift+R
	Admin --> Hidden : Ctrl+Shift+R
	Hidden --> OpportunityOnly : Ctrl+Shift+R

	note right of OpportunityOnly
		Tab visible solo si
		EsFicheroOportunidad()
	end note

	note right of Admin
		Grupo Admin visible
	end note

```

## Sección 3: Inventario de Funcionalidad


### 3.1. Tabla de Funcionalidades Esperadas




| ID | Funcionalidad | Actor/Trigger | Resultado Esperado |
| --- | --- | --- | --- |
| **F001** | Generar gráficos de sensibilidad | Usuario hace clic en botón “Generar Gráficos” del ribbon | Se crean gráficos de sensibilidad en hoja activa basados en curvas de rendimiento |
| **F002** | Invertir ejes de gráfico activo | Usuario hace clic en “Invertir Ejes” con gráfico seleccionado | Los ejes X e Y del gráfico se intercambian |
| **F003** | Formatear hoja CGASING | Usuario hace clic en “Formatear CGASING” estando en hoja CGASING | Aplica formato estándar a la hoja (colores, anchos, etc.) |
| **F004** | Abrir configuración | Usuario hace clic en “Configurador” | Se abre formulario frmConfiguracion con rutas y parámetros |
| **F005** | Crear nueva oportunidad | Usuario hace clic en “Nueva Oportunidad” | Se crea carpeta de oportunidad con estructura de plantilla |
| **F006** | Seleccionar oportunidad | Usuario selecciona item en dropdown “Oportunidades” | Cambia la oportunidad activa del sistema |
| **F007** | Cambiar modo ribbon | Usuario presiona Ctrl+Shift+R | Ribbon cambia entre modos: Admin → Hidden → User → OpportunityOnly → Admin |
| **F008** | Mostrar/ocultar tab ribbon según modo | Cambio de modo ribbon | Tab “Ofertas Especial” visible/oculto según modo y contexto |
| **F009** | Mostrar/ocultar grupo Admin | Cambio de modo ribbon | Grupo “Admin” visible solo en modo Admin |
| **F010** | Habilitar/deshabilitar botón “Generar Gráficos” | Cambio de hoja activa | Botón habilitado solo si es fichero oportunidad Y hoja válida |
| **F011** | Habilitar/deshabilitar botón “Invertir Ejes” | Activar/desactivar gráfico | Botón habilitado solo si hay gráfico activo Y es invertible |
| **F101** | Detectar nueva carpeta de oportunidad | Sistema de archivos crea carpeta en ruta monitoreada | Lista de oportunidades se actualiza automáticamente |
| **F102** | Detectar eliminación de oportunidad | Sistema de archivos elimina carpeta monitoreada | Lista de oportunidades se actualiza |
| **F103** | Detectar cambio en plantilla | Sistema de archivos modifica archivo de plantilla | (Evento capturado, acción TBD) |
| **F104** | Detectar cambio en archivo Gas | Sistema de archivos modifica archivo .gas | (Evento capturado, acción TBD) |
| **F201** | Convertir unidades | Usuario usa fórmula `=ConvertUnits(valor, "from", "to")` | Devuelve valor convertido |
| **F202** | Propiedades termodinámicas | Usuario usa fórmula `=PropsSI("P", "T", 300, "Q", 1, "Water")` | Devuelve propiedad de CoolProp |
| **F203** | Cálculos CGASING | Usuario usa fórmulas `=CGASING_*` | Devuelve cálculos específicos de dominio |


**Total funcionalidades documentadas:** 17


### 3.2. Implementación Actual de Cada Funcionalidad


#### Implementación de F001 - Generar gráficos de sensibilidad


**Flujo técnico paso a paso:**


1. Ribbon XML define: `<button id="btnGenerarGraficos" onAction="OnGenerarGraficos"/>`
2. Excel invoca callback: `modCALLBACKSRibbon.OnGenerarGraficos(control)`
3. Callback delega: `App.RibbonEvents.OnGenerarGraficosDesdeCurvasRto()`
4. clsRibbonEvents dispara evento: `RaiseEvent GenerarGraficosDesdeCurvasRto`
5. clsAplicacion maneja evento: `evRibbon_GenerarGraficosDesdeCurvasRto()`
6. clsAplicacion ejecuta lógica: `Call EjecutarGraficoEnLibroActivo`


**Diagrama de secuencia:**



```
sequenceDiagram
	Usuario->>RibbonXML: Clic botón
	RibbonXML->>modCALLBACKSRibbon: OnGenerarGraficos()
	modCALLBACKSRibbon->>clsRibbonEvents: OnGenerarGraficosDesdeCurvasRto()
	clsRibbonEvents->>clsAplicacion: Event GenerarGraficosDesdeCurvasRto
	clsAplicacion->>moduloGraficos: EjecutarGraficoEnLibroActivo()
	moduloGraficos-->>Usuario: Gráficos creados

```

**Archivos involucrados:**


* `modCALLBACKSRibbon.bas` línea 107
* `clsRibbonEvents.cls` línea 70
* `clsAplicacion.cls` línea 365
* `modMACROGraficoSensibilidad.bas` línea X


#### Implementación de F002 - Invertir ejes de gráfico activo


**Flujo técnico paso a paso:**


1. Ribbon XML define: `<button id="btnInvertirSeries" onAction="OnInvertirEjes"/>`
2. Excel invoca callback: `modCALLBACKSRibbon.OnInvertirEjes(control)`
3. Callback delega: `App.RibbonEvents.OnInvertirEjes()`
4. clsRibbonEvents dispara evento: `RaiseEvent InvertirEjes`
5. clsAplicacion maneja evento: `evRibbon_InvertirEjes()`
6. clsAplicacion ejecuta lógica: `Call InvertirEjesDelGraficoActivo`


**Diagrama de secuencia:**



```
sequenceDiagram
	Usuario->>RibbonXML: Clic botón
	RibbonXML->>modCALLBACKSRibbon: OnInvertirEjes()
	modCALLBACKSRibbon->>clsRibbonEvents: OnInvertirEjes()
	clsRibbonEvents->>clsAplicacion: Event InvertirEjes
	clsAplicacion->>moduloGraficos: InvertirEjesDelGraficoActivo()
	moduloGraficos-->>Usuario: Ejes invertidos

```

**Archivos involucrados:**


* `modCALLBACKSRibbon.bas` línea 112
* `clsRibbonEvents.cls` línea 75
* `clsAplicacion.cls` línea 371
* `modMACROGraficoSensibilidad.bas` línea X


#### Implementación de F101 - Detectar nueva carpeta de oportunidad


**Flujo técnico paso a paso:**


1. clsFSWatcher detecta creación de subcarpeta
2. clsFSWatcher dispara evento: `SubfolderCreated(parentFolder, subfolderName)`
3. clsFSMonitoringCoord maneja evento: `mFolderWatcher_SubfolderCreated()`
4. clsFSMonitoringCoord verifica si es carpeta de oportunidades
5. clsFSMonitoringCoord dispara evento: `OpportunityCreated(parentFolder, subfolderName)`
6. clsAplicacion maneja evento: `mFSMonitoringCoord_OpportunityCreated()`
7. clsAplicacion delega a: `clsOpportunitiesMgr.ProcesarCambiosEnOportunidades()`


**Diagrama de secuencia:**



```
sequenceDiagram
	SistemaArchivos->>clsFSWatcher: Crear carpeta
	clsFSWatcher->>clsFSMonitoringCoord: Event SubfolderCreated
	clsFSMonitoringCoord->>clsFSMonitoringCoord: Verificar tipo carpeta
	clsFSMonitoringCoord->>clsAplicacion: Event OpportunityCreated
	clsAplicacion->>clsOpportunitiesMgr: ProcesarCambiosEnOportunidades
	clsOpportunitiesMgr-->>Usuario: Oportunidad añadida

```

**Archivos involucrados:**


* `clsFSWatcher.cls` (externo)
* `clsFSMonitoringCoord.cls` línea 150
* `clsAplicacion.cls` línea 278
* `clsOpportunitiesMgr.cls` línea 145


## Sección 4: Patrones y Anti-Patrones


### 4.1. Patrones Identificados


1. **Patrón Fachada (Facade)**: `clsAplicacion` expone una interfaz simplificada a los servicios
2. **Patrón Observador (Observer/Observable)**: Uso extensivo de `WithEvents` para suscribirse a eventos
3. **Patrón Adaptador**: `clsExecutionContext` adapta el acceso al contexto de Excel
4. **Patrón Singleton**: `App()` en `ThisWorkbook` como punto de acceso global
5. **Patrón Estratégia**: Diferentes modos de Ribbon implementados como estrategias


### 4.2. Anti-Patrones Identificados


1. **Dios (God Object)**: `clsAplicacion` concentra demasiadas responsabilidades
2. **Código Espagueti**: Excesiva interconexión entre componentes
3. **Acoplamiento Estrecho**: Muchas clases dependen directamente de otras
4. **Singleton Global**: Uso de `App()` como acceso global a la aplicación
5. **Event Handler Prolífico**: Muchos manejadores de eventos en una sola clase


## Sección 5: Reglas y Restricciones


### 5.1. Reglas de Arquitectura


1. **Regla de Inicialización**: Todos los servicios deben inicializarse en orden correcto
2. **Regla de Eventos**: Los eventos deben propagarse de forma consistente
3. **Regla de Recursos**: Los recursos COM deben liberarse adecuadamente
4. **Regla de Configuración**: La configuración debe persistirse y cargarse del registro


### 5.2. Restricciones Técnicas


1. **Restricción de Memoria**: Limitación de objetos COM no liberados
2. **Restricción de Contexto**: Acceso seguro al contexto de Excel
3. **Restricción de Seguridad**: Acceso restringido al sistema de archivos
4. **Restricción de Interfaz**: Ribbon debe mantenerse funcional ante desconexiones


## Sección 6: Cómo Usar Este Documento


Este documento sirve como guía de referencia para:


1. **Entender la arquitectura actual** del sistema
2. **Identificar puntos de mejora** en la estructura del código
3. **Facilitar la incorporación** de nuevos desarrolladores
4. **Apoyar decisiones de refactorización** y mantenimiento
5. **Documentar el comportamiento** del sistema para futuras referencias




