VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRibbonState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("2-Infrastructure.State")
Option Explicit

' ===============================================================================
' CLASE: clsRibbonState
' ===============================================================================
' PROPOSITO:
'   Almacena estado lógico del Ribbon (modo actual, visibilidad).
'   Determina qué tabs/controles están visibles y habilitados.
'
'   4 modos:
'     0 = OpportunityOnly (tab visible solo si es fichero oportunidad)
'     1 = User (tab siempre visible, grupo Admin oculto)
'     2 = Admin (tab visible, grupo Admin visible)
'     3 = Hidden (tab completamente oculto)
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   USADO POR:
'     - clsRibbonManager (lee estado para invalidar controles)
'     - modRibbonCallbacks (lee estado para GetTabVisible, GetControlEnabled)
'     - clsApplication (cambia modo con ToggleModo)
'
'   CONTENIDO EN:
'     - clsApplicationState
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - PULL: Properties Get/Set
'   - ELIMINA: Evento StateChanged (ahora property setter notifica directamente)
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' -------------------------------------------------------------------------------
' Archivo origen: clsRibbonState.cls (líneas 1-100)
'
' F-001: Gestión de modo ribbon
'   - Legacy: mModoRibbon As RibbonModeEnum línea ~10
'   - Legacy: ToggleModo() línea ~45
'   - Migra: Property ModoActual As Long (Get/Set)
'   - Migra: Sub ToggleModo()
'
' F-002: Visibilidad de tab
'   - Legacy: TabVisible() As Boolean línea ~67
'   - Migra: Property TabVisible As Boolean (calculado)
'
' F-003: Visibilidad de grupo Admin
'   - Legacy: GrupoAdminVisible() As Boolean línea ~89
'   - Migra: Property AdminGroupVisible As Boolean (calculado)
'
' ELIMINACIONES:
'   - ❌ Public Event StateChanged (línea ~12)
'   - Reemplazado por: Property setter llama directamente a RibbonManager.InvalidateRibbon()
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 2: Property ModoActual As Long
' TODO Sprint 2: Property TabVisible As Boolean
' TODO Sprint 2: Property AdminGroupVisible As Boolean
' TODO Sprint 2: Sub ToggleModo()
