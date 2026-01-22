VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileSystemMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("2-Infrastructure.FileSystem")
Option Explicit

' ===============================================================================
' CLASE: clsFileSystemMonitor
' ===============================================================================
' PROPOSITO:
'   Coordinador e intérprete de eventos de sistema de archivos.
'   Suscribe a clsFileSystemWatcher (WithEvents), interpreta eventos
'   según contexto de negocio, y dispara EVENTOS DE DOMINIO.
'
'   Clasifica eventos en:
'     - Oportunidades (carpetas en ruta oportunidades)
'     - Plantillas (archivos .xlsx en ruta plantillas)
'     - Gas (archivos .xlsx en ruta gas)
'
'   Único caso válido de WithEvents + RaiseEvent dentro de arquitectura limpia
'   (porque son eventos de DOMINIO, potencial 1-a-N).
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   ESCUCHA (WithEvents):
'     - clsFileSystemWatcher (eventos de FS)
'
'   USA:
'     - clsApplicationConfiguration (para obtener rutas)
'
'   DISPARA (RaiseEvent - EVENTOS DE DOMINIO):
'     - OpportunityCreated(parentFolder, subfolderName)
'     - OpportunityDeleted(parentFolder, subfolderName)
'     - OpportunityRenamed(parentFolder, oldName, newName)
'     - OpportunityItemDeleted(folder, fileName)
'     - OpportunityItemRenamed(folder, oldName, newName)
'     - TemplateCreated(folder, fileName)
'     - TemplateChanged(folder, fileName)
'     - GasFileCreated(folder, fileName)
'     - GasFileChanged(folder, fileName)
'     - MonitoringError(folder, errorMessage)
'     - MonitoringReconnected(folder, attempts)
'     - MonitoringFailed(folder, reason)
'
'   ESCUCHADO POR:
'     - clsApplication (único suscriptor actual)
'     - Potencialmente otros (justifica RaiseEvent)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - EVENTS: WithEvents clsFileSystemWatcher
'   - EVENTS: RaiseEvent (eventos de DOMINIO, 1-a-N)
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' ===============================================================================
' Archivo origen: clsFSMonitoringCoord.cls (líneas 1-647)
'
' F-001: Coordinación de monitoreo
'   - Legacy: Private WithEvents mFolderWatcher As clsFSWatcher (línea 15)
'   - Legacy: Sub IniciarMonitoreo(oDicFolders) (línea 72)
'   - Migra: Mantiene WithEvents + IniciarMonitoreo
'
' F-002: Almacenamiento de rutas configuradas
'   - Legacy: m_rutaOportunidades, m_rutaPlantillas, m_rutaGasVBNet (líneas 19-21)
'   - Migra: Private _configuredPaths As clsApplicationConfiguration
'
' F-003: Eventos de subcarpetas → Oportunidades
'   - Legacy: mFolderWatcher_SubfolderCreated (línea 146)
'   - Legacy: mFolderWatcher_SubfolderDeleted (línea 157)
'   - Legacy: mFolderWatcher_SubfolderRenamed (línea 167)
'   - Migra: Clasifica por ruta → RaiseEvent OpportunityCreated/Deleted/Renamed
'
' F-004: Eventos de archivos → Plantillas / Gas
'   - Legacy: mFolderWatcher_FileCreated (línea 177)
'   - Legacy: mFolderWatcher_FileChanged (línea 202)
'   - Legacy: mFolderWatcher_FileDeleted (línea 192)
'   - Migra: Clasifica por ruta → RaiseEvent TemplateCreated/GasFileCreated
'
' F-005: Eventos de error
'   - Legacy: mFolderWatcher_ErrorOccurred (línea 233)
'   - Legacy: mFolderWatcher_WatcherReconnected (línea 246)
'   - Legacy: mFolderWatcher_WatcherReconnectionFailed (línea 255)
'   - Migra: RaiseEvent MonitoringError/Reconnected/Failed
'
' F-006: Funciones de configuración rápida
'   - Legacy: ConfigurarMonitoreoOportunidades (línea 292)
'   - Legacy: ConfigurarMonitoreoPlantillas (línea 320)
'   - Legacy: ConfigurarMonitoreoGasVBNet (línea 345)
'   - Migra: Métodos de configuración especializada
'
' F-007: Funciones de diagnóstico
'   - Legacy: VerEstadisticasMonitoreo (línea 374)
'   - Legacy: VerHistorialMonitoreo (línea 427)
'   - Legacy: VerConfiguracionWatcher (línea 495)
'   - Migra: Métodos de diagnóstico (delegar a clsFileSystemWatcher)
'
' F-008: Helpers de comparación de rutas
'   - Legacy: EsRutaOportunidades, EsRutaPlantillas, EsRutaGasVBNet (líneas 271-284)
'   - Migra: Private helpers para clasificar eventos
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 3: Private WithEvents _watcher As clsFileSystemWatcher
' TODO Sprint 3: Private _config As clsApplicationConfiguration
' TODO Sprint 3: Public Sub StartMonitoring(config As clsApplicationConfiguration)
' TODO Sprint 3: Implementar 10 manejadores _watcher_* con clasificación
' TODO Sprint 3: Implementar helpers de clasificación de rutas
' TODO Sprint 3: Métodos de configuración rápida
