Attribute VB_Name = "mod_Constants"
'@Folder("2-Infrastructure.Logging")
Option Explicit

' ===============================================================================
' MODULO: mod_Constants
' ===============================================================================
' PROPOSITO:
'   Constantes globales y enumeraciones usadas en toda la aplicación.
'   Incluye:
'     - Niveles de log (LogLevelEnum)
'     - Claves de registro (CFG_*)
'     - Modos de ribbon (RibbonModeEnum)
'     - Constantes de archivo (FILE_*)
'     - Timeouts y límites
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES/MODULOS:
' -------------------------------------------------------------------------------
'   USADO POR:
'     - TODAS las clases/módulos del sistema
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' ===============================================================================
' Archivo origen: mod_ConstantsGlobals.bas (líneas 1-302)
'
' F-001: Enumeración de niveles de log
'   - Legacy: Public Enum LogLevelEnum (línea ~10)
'   - Migra: LogLevelEnum: LOG_DEBUG=0, LOG_INFO=1, LOG_WARNING=2, LOG_ERROR=3
'
' F-002: Constantes de configuración (Registry)
'   - Legacy: CFG_RUTA_OPORTUNIDADES, CFG_RUTA_PLANTILLAS, etc (líneas ~30-50)
'   - Migra: Public Const CFG_RUTA_* As String
'
' F-003: Enumeración de modos de ribbon
'   - Legacy: Public Enum RibbonModeEnum (línea ~60)
'   - Migra: rmOpportunityOnly=0, rmUser=1, rmAdmin=2, rmHidden=3
'
' F-004: Constantes de archivo
'   - Legacy: FILE_EXTENSION_XLSM, FILE_EXTENSION_XLSX (líneas ~80-90)
'   - Migra: Public Const FILE_* As String
'
' F-005: Tipos personalizados (UDTs)
'   - Legacy: Type FileInfo, Type ChartInfo (líneas ~100-150)
'   - Migra: Mantener UDTs necesarios
'
' F-006: Constantes de módulo/clase
'   - Legacy: MODULE_NAME disperso en cada clase
'   - Migra: Cada clase define su propio MODULE_NAME (no global)
'
' F-007: Timeouts y límites
'   - Legacy: TIMEOUT_FS_MONITOR, MAX_LOG_DAYS (líneas ~200-250)
'   - Migra: Public Const TIMEOUT_*, MAX_*
'
' ELIMINACIONES:
'   - Variables globales de estado (mover a clsApplicationState)
'   - Variables de instancias globales (mover a clsApplication)
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 1: Public Enum LogLevelEnum
' TODO Sprint 1: Public Enum RibbonModeEnum
' TODO Sprint 1: Public Const CFG_RUTA_* As String
' TODO Sprint 1: Public Const FILE_* As String
' TODO Sprint 2: Types necesarios
' TODO Sprint 2: Timeouts y límites
