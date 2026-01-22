VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsApplicationConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("2-Infrastructure.Configuration")
Option Explicit

' ===============================================================================
' CLASE: clsApplicationConfiguration
' ===============================================================================
' PROPOSITO:
'   Gestiona configuración persistente de la aplicación almacenada en
'   el registro de Windows (HKEY_CURRENT_USER).
'
'   Lee y escribe rutas de carpetas monitoreadas:
'     - Ruta de oportunidades
'     - Ruta de plantillas
'     - Ruta de gas VBNet
'     - Ruta de plantilla Excel Calc
'
'   Expone Dictionary de carpetas a monitorear para clsFileSystemMonitor.
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   USADA POR:
'     - clsApplication (lectura en Initialize)
'     - clsFileSystemMonitor (obtiene rutas para clasificar eventos)
'     - clsOpportunityManager (obtiene ruta base de oportunidades)
'     - frmConfiguracion (UI para editar configuración)
'
'   USA:
'     - API Registry de Windows (SaveSetting, GetSetting)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - PULL: Clientes leen properties (RutaOportunidades, etc)
'   - Sin eventos
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' ===============================================================================
' Archivo origen: clsConfiguration.cls (líneas 1-254)
'
' F-001: Propiedades de rutas
'   - Legacy: RutaOportunidades (líneas 45-65)
'   - Legacy: RutaPlantillas (líneas 67-87)
'   - Legacy: RutaGasVBNet (líneas 89-102)
'   - Legacy: RutaExcelCalcTempl (líneas 104-117)
'   - Migra: Properties Get/Let para cada ruta
'
' F-002: Dictionary de carpetas a monitorear
'   - Legacy: oDicFoldersToWatch (línea 119)
'   - Migra: Property Get FoldersToWatch() As Object
'   - Devuelve Dictionary con claves CFG_RUTA_* → valores = rutas
'
' F-003: Carga desde registro
'   - Legacy: CargarDesdeRegistro() (línea 134)
'   - Migra: Public Sub LoadFromRegistry()
'   - Lee: HKEY_CURRENT_USER\Software\ABC\Ofertas\RutaBase, RutaPlantillas, etc
'
' F-004: Guardado en registro
'   - Legacy: GuardarEnRegistro() (línea 178)
'   - Migra: Public Sub SaveToRegistry()
'   - Escribe: HKEY_CURRENT_USER\Software\ABC\Ofertas\...
'
' F-005: Constantes de claves de registro
'   - Legacy: CFG_RUTA_OPORTUNIDADES, CFG_RUTA_PLANTILLAS, etc (en mod_Constants)
'   - Migra: Constantes privadas en clase o en mod_Constants
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 1: Properties para cada ruta (Get/Let)
' TODO Sprint 1: Public Sub LoadFromRegistry()
' TODO Sprint 1: Public Sub SaveToRegistry()
' TODO Sprint 1: Property Get FoldersToWatch() As Object
' TODO Sprint 2: Validación de rutas (RutaExiste helper)
