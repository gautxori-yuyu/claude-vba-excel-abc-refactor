VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRibbonEventAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("2-Infrastructure.ExcelIntegration")
Option Explicit

' ===============================================================================
' CLASE: clsRibbonEventAdapter
' ===============================================================================
' PROPOSITO:
'   Adaptador OPCIONAL de eventos COM de IRibbonUI.
'
'   NOTA: IRibbonUI generalmente NO tiene eventos útiles para WithEvents.
'   Esta clase se mantiene por completitud arquitectónica, pero puede
'   estar VACÍA en la práctica.
'
'   Si hay eventos útiles de IRibbonUI, se capturan aquí y se envían
'   como callbacks a clsRibbonManager.
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   ESCUCHA (WithEvents - si aplica):
'     - IRibbonUI (eventos COM - RARAMENTE usados)
'
'   LLAMA (callbacks directos):
'     - clsRibbonManager (si hay eventos relevantes)
'
'   USADO POR:
'     - clsRibbonManager
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - EVENTS: WithEvents IRibbonUI (si hay eventos relevantes)
'   - CALLBACKS: Llamadas directas a clsRibbonManager
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' -------------------------------------------------------------------------------
' NO APLICA - IRibbonUI no tiene eventos relevantes en VBA
'
' En legacy, clsRibbonEvents NO tenía WithEvents de IRibbonUI.
' Solo almacenaba el puntero mRibbonUI y disparaba eventos custom.
'
' Esta clase se mantiene por arquitectura, pero probablemente quedará vacía.
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 2: Evaluar si IRibbonUI tiene eventos útiles
' TODO Sprint 2: Si no → marcar clase como PLACEHOLDER
' TODO Sprint 2: Si sí → implementar WithEvents + callbacks
