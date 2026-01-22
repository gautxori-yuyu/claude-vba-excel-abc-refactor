Attribute VB_Name = "modRibbonCallbacks"
'@Folder("6-Presentation.RibbonCallbacks")
Option Explicit

' ===============================================================================
' MODULO: modRibbonCallbacks
' ===============================================================================
' PROPOSITO:
'   Callbacks del Ribbon XML - punto de entrada para acciones de usuario.
'   Recibe eventos desde la UI del Ribbon y delega DIRECTAMENTE a clsApplication.
'
'   NO contiene lógica de negocio, solo delegación.
'   Patrón: Ribbon XML → Callback → clsApplication.MetodoPublico()
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES/MODULOS:
' -------------------------------------------------------------------------------
'   LLAMA:
'     - clsApplication (mediante función global App())
'     - clsRibbonManager (para GetControlEnabled)
'
'   LLAMADO POR:
'     - Ribbon XML (IRibbonUI)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - PUSH: Llama directamente a métodos públicos de clsApplication
'   - PULL: Consulta estado para callbacks GetEnabled, GetVisible
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' ===============================================================================
' Archivo origen: modCALLBACKSRibbon.bas (líneas 1-332)
'
' F-001: Inicialización del Ribbon
'   - Legacy: RibbonOnLoad(xlRibbon As IRibbonUI) (línea 18)
'   - Migra: Public Sub RibbonOnLoad(xlRibbon As IRibbonUI)
'   - Pasa puntero IRibbonUI a App.RibbonManager
'
' F-002: Callbacks de acciones (botones)
'   - Legacy: OnGenerarGraficosDesdeCurvasRto(control) (línea 104)
'   - Legacy: OnInvertirEjes(control) (línea 109)
'   - Legacy: OnFormatearCGASING(control) (línea 114)
'   - Legacy: OnNuevaOportunidad(control) (línea 119)
'   - Legacy: OnReplaceWithNamesInValidations(control) (línea 124)
'   - Legacy: OnConfigurador(control) (línea 129)
'   - Migra: Mantiene las 6 funciones con DELEGACION DIRECTA
'   - CAMBIA: De App.Ribbon.OnXXX (evento) → App.ComandoXXX (método directo)
'
' F-003: Callback de visibilidad de tab
'   - Legacy: GetTabVisible(control, ByRef visible) (línea 145)
'   - Migra: Public Sub GetTabVisible(control, ByRef visible)
'   - Consulta App.RibbonManager.GetTabVisible(control.id)
'
' F-004: Callback de visibilidad de grupo
'   - Legacy: GetGroupVisible(control, ByRef visible) (línea 167)
'   - Migra: Public Sub GetGroupVisible(control, ByRef visible)
'   - Consulta App.RibbonManager.GetGroupVisible(control.id)
'
' F-005: Callback de estado enabled de controles
'   - Legacy: GetControlEnabled(control, ByRef enabled) (línea 189)
'   - Migra: Public Sub GetControlEnabled(control, ByRef enabled)
'   - Consulta App.RibbonManager.GetControlEnabled(control.id)
'   - IMPORTANTE: Mueve lógica de negocio de callback a RibbonManager
'
' SIMPLIFICACION:
'   - Antes: Callback → RibbonEvents.OnXXX → RaiseEvent → clsAplicacion → Lógica
'   - Ahora: Callback → clsApplication.MetodoPublico → Lógica
'   - Elimina 2 niveles de indirección
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 1: Public Sub RibbonOnLoad(xlRibbon As IRibbonUI)
' TODO Sprint 2: Callbacks de 6 acciones (delegar a App.XXX)
' TODO Sprint 2: Callbacks de visibilidad (delegar a RibbonManager)
' TODO Sprint 2: Callback GetControlEnabled (delegar a RibbonManager)
' TODO Sprint 3: Documentar mapping Ribbon XML ↔ Callbacks
