VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("2-Infrastructure.DependencyInjection")
Option Explicit

' ===============================================================================
' INTERFAZ: IService
' ===============================================================================
' PROPOSITO:
'   Interfaz (contrato) que TODOS los servicios deben implementar.
'   Define métodos de ciclo de vida estándar: Initialize() y Dispose().
'
'   Permite a clsServiceContainer gestionar servicios de forma polimórfica.
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   IMPLEMENTADA POR:
'     - clsChartEventManager
'     - clsFileManager
'     - clsOpportunityManager
'     - clsRibbonManager
'     - clsFileSystemMonitor
'     - (todos los servicios)
'
'   USADA POR:
'     - clsServiceContainer (para gestión polimórfica)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - Interfaz pura (sin implementación)
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' -------------------------------------------------------------------------------
' NO EXISTE EN MAIN - NUEVA FUNCIONALIDAD
'
' Estandariza el ciclo de vida de servicios que en legacy tenían
' métodos inconsistentes (algunos Initialize, otros Start, otros nada).
'
' ===============================================================================

' --- CONTRATO DE INTERFAZ ---
' TODO Sprint 1: Public Sub Initialize()
' TODO Sprint 1: Public Sub Dispose()
