VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("1-Application")
Option Explicit

' ===============================================================================
' CLASE: ThisWorkbook
' ===============================================================================
' PROPOSITO:
'   Entry point de la aplicación Excel. Captura eventos de ciclo de vida del
'   workbook (Open, BeforeClose) y delega TODA la lógica a clsApplication.
'
'   Contiene CERO lógica de negocio. Solo crea y destruye clsApplication.
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   CREA:
'     - clsApplication (única instancia global)
'
'   EXPONE:
'     - Function App() As clsApplication (para acceso global)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - PUSH: Llama a clsApplication.Initialize() en Workbook_Open
'   - PUSH: Llama a clsApplication.Terminate() en Workbook_BeforeClose
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' -------------------------------------------------------------------------------
' Archivo origen: ThisWorkbook.cls (líneas 1-172)
'
' F-001: Inicialización en apertura del workbook
'   - Legacy: Workbook_Open() líneas 17-58
'   - Migra: removeOldLogs() - mover a mod_Logger
'   - Migra: InitLogger() - mover a mod_Logger
'   - Migra: AutoInstalador() - mover a modMacroAppLifecycle
'   - Migra: AutoRegistrarTodasLasUDFs() - mover a modAppUDFsRegistration
'   - Migra: Application.OnKey "^+R" - atajo teclado ribbon
'   - Migra: Creación de mApp (clsApplication)
'   - SIMPLIFICA: Solo 3 líneas aquí, resto en servicios
'
' F-002: Limpieza en cierre del workbook
'   - Legacy: Workbook_BeforeClose() líneas 60-71
'   - Migra: App.Terminate()
'   - Migra: Set mApp = Nothing
'
' F-003: Función global de acceso App()
'   - Legacy: App() líneas 80-96
'   - Migra: Mantiene patrón Singleton de acceso global
'
' ELIMINACIONES (mover a otros lugares):
'   - InstallAddin() líneas 105-111 → modMacroAppLifecycle
'   - IsXLAMFile() líneas 113-119 → modUtils
'   - Test_*() líneas 121-172 → Eliminar (tests de desarrollo)
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 1: Implementar Workbook_Open con inicialización mínima
' TODO Sprint 1: Implementar Workbook_BeforeClose con limpieza
' TODO Sprint 1: Implementar función App() como Singleton
