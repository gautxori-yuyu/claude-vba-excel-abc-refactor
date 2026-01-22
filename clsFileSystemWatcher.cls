VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileSystemWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("2-Infrastructure.FileSystem")
Option Explicit

' ===============================================================================
' CLASE: clsFileSystemWatcher
' ===============================================================================
' PROPOSITO:
'   Adaptador de eventos COM del componente externo FolderWatcher.
'   Captura eventos de cambios en sistema de archivos (FileCreated,
'   SubfolderCreated, etc) y los re-emite como eventos VBA.
'
'   Único WithEvents autorizado para FolderWatcher COM.
'   Único caso válido de RaiseEvent (adaptador de eventos externos).
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   ESCUCHA (WithEvents):
'     - FolderWatcher (componente COM externo)
'
'   DISPARA (RaiseEvent - VÁLIDO aquí):
'     - SubfolderCreated(parentPath, folderName)
'     - SubfolderDeleted(parentPath, folderName)
'     - SubfolderRenamed(parentPath, oldName, newName)
'     - FileCreated(folder, fileName)
'     - FileDeleted(folder, fileName)
'     - FileChanged(folder, fileName)
'     - FileRenamed(folder, oldName, newName)
'     - ErrorOccurred(folder, errorMsg)
'     - WatcherReconnected(folder, attempts)
'     - WatcherReconnectionFailed(folder, reason)
'
'   ESCUCHADO POR:
'     - clsFileSystemMonitor (único suscriptor)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - EVENTS: WithEvents FolderWatcher COM
'   - EVENTS: RaiseEvent (adaptador, justificado)
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' ===============================================================================
' Archivo origen: clsFSWatcher.cls (líneas 1-713)
'
' F-001: Wrapper de componente FolderWatcher COM
'   - Legacy: Private WithEvents fw As FolderWatcher (línea ~20)
'   - Migra: Mantiene WithEvents (necesario para COM externo)
'
' F-002: Inicio de monitoreo de carpetas
'   - Legacy: Sub IniciarMonitoreo (línea ~100)
'   - Migra: Sub StartMonitoring(folderPath, includeSubdirs, filterPattern, etc)
'
' F-003: Detención de monitoreo
'   - Legacy: Sub Dispose (línea ~200)
'   - Migra: Sub StopMonitoring()
'
' F-004: Eventos de subcarpetas
'   - Legacy: fw_SubfolderCreated (línea ~300)
'   - Legacy: fw_SubfolderDeleted (línea ~320)
'   - Legacy: fw_SubfolderRenamed (línea ~340)
'   - Migra: RaiseEvent SubfolderCreated/Deleted/Renamed
'
' F-005: Eventos de archivos
'   - Legacy: fw_FileCreated (línea ~400)
'   - Legacy: fw_FileDeleted (línea ~420)
'   - Legacy: fw_FileChanged (línea ~440)
'   - Legacy: fw_FileRenamed (línea ~460)
'   - Migra: RaiseEvent FileCreated/Deleted/Changed/Renamed
'
' F-006: Eventos de error y reconexión
'   - Legacy: fw_ErrorOccurred (línea ~500)
'   - Legacy: fw_WatcherReconnected (línea ~520)
'   - Migra: RaiseEvent ErrorOccurred/WatcherReconnected
'
' F-007: Configuraciones especiales
'   - Legacy: ConfigurarFiltroSoloCarpetas (línea ~550)
'   - Legacy: ConfigurarFiltroTamaño (línea ~570)
'   - Legacy: ConfigurarAccionMover (línea ~600)
'   - Migra: Métodos de configuración
'
' F-008: Diagnóstico
'   - Legacy: CarpetasActivas, ObtenerEstadisticas, ObtenerHistorial (líneas ~620-650)
'   - Migra: Métodos de diagnóstico
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 3: Private WithEvents _watcher As FolderWatcher
' TODO Sprint 3: Public Sub StartMonitoring(...)
' TODO Sprint 3: Public Sub StopMonitoring()
' TODO Sprint 3: Implementar 10 manejadores _watcher_* con RaiseEvent
' TODO Sprint 3: Métodos de configuración y diagnóstico
