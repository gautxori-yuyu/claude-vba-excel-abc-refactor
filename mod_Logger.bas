Attribute VB_Name = "mod_Logger"
'@Folder("2-Infrastructure.Logging")
Option Explicit

' ===============================================================================
' MODULO: mod_Logger
' ===============================================================================
' PROPOSITO:
'   Sistema de logging centralizado para toda la aplicación.
'   Escribe logs en archivo de texto con niveles de severidad:
'     - DEBUG: Información detallada para desarrollo
'     - INFO: Mensajes informativos normales
'     - WARNING: Advertencias no críticas
'     - ERROR: Errores críticos
'
'   Gestiona rotación de logs antiguos (elimina logs con más de X días).
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES/MODULOS:
' -------------------------------------------------------------------------------
'   USADO POR:
'     - TODAS las clases/módulos del sistema
'
'   USA:
'     - FileSystemObject (para escritura de archivos)
'     - mod_Constants (para constantes LOG_*)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - PULL: Clientes llaman LogDebug(), LogInfo(), LogWarning(), LogError()
'   - Sin eventos
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' ===============================================================================
' Archivo origen: mod_Logger.bas (líneas 1-223)
'
' F-001: Inicialización del logger
'   - Legacy: InitLogger(logLevel, enabled, logPath) (línea ~20)
'   - Migra: Public Sub InitLogger(...)
'   - Crea archivo de log con timestamp en nombre
'
' F-002: Funciones de logging por nivel
'   - Legacy: LogDebug(module, message) (línea ~50)
'   - Legacy: LogInfo(module, message) (línea ~60)
'   - Legacy: LogWarning(module, message) (línea ~70)
'   - Legacy: LogError(module, procName, errNum, errDesc) (línea ~80)
'   - Migra: Mantiene las 4 funciones públicas
'
' F-003: Rotación de logs antiguos
'   - Legacy: removeOldLogs() (línea ~120)
'   - Migra: Public Sub RemoveOldLogs(daysToKeep)
'   - Elimina archivos .log con más de N días
'
' F-004: Deshabilitación temporal de logging
'   - Legacy: Variable global gLogEnabled (línea ~10)
'   - Migra: Public Property Let Enabled(value As Boolean)
'
' F-005: Configuración de nivel de log
'   - Legacy: Variable global gLogLevel (línea ~11)
'   - Migra: Public Property Let LogLevel(level As LogLevelEnum)
'
' F-006: Helpers privados
'   - Legacy: GetLogFilePath(), FormatLogEntry() (líneas ~150-180)
'   - Migra: Private helpers para formateo
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 1: Public Sub InitLogger(logLevel, enabled, logPath)
' TODO Sprint 1: Public Sub LogDebug/Info/Warning/Error(...)
' TODO Sprint 1: Public Sub RemoveOldLogs(daysToKeep)
' TODO Sprint 1: Properties Enabled, LogLevel
' TODO Sprint 2: Private helpers de formateo y escritura
