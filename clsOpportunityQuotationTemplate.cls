VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpportunityQuotationTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("4-Domain.Templates")
Option Explicit

' ===============================================================================
' CLASE: clsOpportunityQuotationTemplate
' ===============================================================================
' PROPOSITO:
'   Template de dominio para generar cotizaciones de oportunidades.
'   Conoce la estructura del documento de cotización y cómo poblarlo
'   con datos de oportunidad y ofertas.
'
'   DOMINIO PURO: Lógica de generación de cotización sin dependencias
'   de infraestructura.
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   USA:
'     - clsOpportunity (datos de oportunidad)
'     - clsOffer (datos de ofertas)
'
'   USADA POR:
'     - clsOpportunityManager (genera cotizaciones)
'     - Commands de aplicación (CmdGenerateQuotation - si existe)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - PULL: Recibe datos de oportunidad y ofertas
'   - PUSH: Devuelve estructura de cotización generada
'   - Sin eventos
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' ===============================================================================
' Archivo origen: clsOpportunityOfferQuotationTpl.cls (líneas 1-12)
'
' NOTA: Clase legacy muy simple (12 líneas) - probablemente placeholder.
'
' F-001: Generación de cotización
'   - Legacy: Clase mínima sin implementación real
'   - NUEVO: Public Function Generate(opp As clsOpportunity, offers As Collection) As Variant
'
' F-002: Validación de datos
'   - NUEVO: Public Function Validate(opp, offers) As Boolean
'
' EXPANSION NECESARIA en Sprint 2 tras analizar uso en legacy.
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 2: Analizar uso real de templates en legacy
' TODO Sprint 2: Public Function Generate(opp, offers) As Variant
' TODO Sprint 3: Public Function Validate(opp, offers) As Boolean
