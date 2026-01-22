VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelExecutionContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("2-Infrastructure.ExcelIntegration")
Option Explicit

' ===============================================================================
' CLASE: clsExcelExecutionContext
' ===============================================================================
' PROPOSITO:
'   Adaptador de eventos COM de Excel.Application.
'   ÚNICO WithEvents autorizado para Application.
'
'   Captura eventos de Excel (WorkbookOpen, SheetActivated, etc) y
'   ejecuta CALLBACKS DIRECTOS en clsApplication (NO re-emite eventos custom).
'
'   Patrón: WithEvents COM → Callback directo (elimina cascada de eventos)
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   ESCUCHA (WithEvents):
'     - Excel.Application (eventos COM)
'
'   LLAMA (callbacks directos):
'     - clsApplication.OnWorkbookOpened(wb)
'     - clsApplication.OnWorkbookActivated(wb)
'     - clsApplication.OnSheetActivated(sh)
'     - clsApplication.OnSheetDeactivated(sh)
'     - clsApplication.OnSelectionChanged(rng)
'
'   CONFIGURADO POR:
'     - clsApplication (inyecta callback handler en Initialize)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - EVENTS: WithEvents Application (único permitido)
'   - CALLBACKS: Llamadas directas a clsApplication
'   - NO re-emite RaiseEvent (elimina cascada)
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' -------------------------------------------------------------------------------
' Archivo origen: clsExecutionContext.cls (líneas 1-289)
'
' F-001: Wrapper de eventos Excel.Application
'   - Legacy: Private WithEvents m_xlApp As Application (línea 16)
'   - Migra: Mantiene WithEvents (necesario para capturar eventos COM)
'
' F-002: Eventos capturados y re-emitidos
'   - Legacy: 6 eventos re-emitidos con RaiseEvent (líneas 45-81)
'   - CAMBIA: 6 eventos → callbacks directos (elimina RaiseEvent)
'
' Eventos migrados:
'   - m_xlApp_WorkbookOpen → OnWorkbookOpened(wb)
'   - m_xlApp_WorkbookActivate → OnWorkbookActivated(wb)
'   - m_xlApp_WorkbookBeforeClose → OnWorkbookBeforeClose(wb, cancel)
'   - m_xlApp_SheetActivate → OnSheetActivated(sh)
'   - m_xlApp_SheetDeactivate → OnSheetDeactivated(sh)
'   - m_xlApp_SheetSelectionChange → OnSelectionChanged(rng)
'
' ELIMINACIONES:
'   - ❌ Public Event WorkbookOpened (línea 23)
'   - ❌ Public Event WorkbookActivated (línea 24)
'   - ❌ Public Event SheetActivated (línea 27)
'   - ❌ 3 eventos más (total 6 eventos eliminados)
'   - Reemplazados por: callbacks directos
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 1: Private WithEvents _app As Application
' TODO Sprint 1: Private _callbackHandler As clsApplication
' TODO Sprint 1: Public Sub Initialize(handler As clsApplication)
' TODO Sprint 2: Implementar 6 manejadores _app_* con callbacks directos
