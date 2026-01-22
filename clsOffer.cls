VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("4-Domain.Entities")
Option Explicit

' ===============================================================================
' CLASE: clsOffer
' ===============================================================================
' PROPOSITO:
'   Entidad de dominio que representa una oferta de compresor.
'   Atributos básicos de oferta (tipo, precio, especificaciones).
'
'   DOMINIO PURO: Sin dependencias de infraestructura.
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   USADA POR:
'     - clsOfferRepository (crea y persiste)
'     - clsOpportunityManager (asocia con oportunidades)
'
'   USA:
'     - modOfferTypes (tipos de ofertas - enumeraciones)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - PULL: Clientes leen/escriben properties
'   - Sin eventos
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' ===============================================================================
' Archivo origen: clsOferta.cls (líneas 1-51)
'
' F-001: Atributos de oferta
'   - Legacy: Private m* As Variant (líneas ~10-20)
'   - Migra: Private _offerType As OfferTypeEnum
'   - Migra: Private _price As Double
'   - Migra: Private _specifications As String
'   - (otros atributos según análisis detallado de legacy)
'
' F-002: Propiedades de acceso
'   - Legacy: Property Get/Let para cada atributo (líneas ~22-48)
'   - Migra: Public Property Get/Let para cada atributo
'
' F-003: Validación de oferta
'   - Legacy: NO EXISTE
'   - NUEVO: Public Function IsValid() As Boolean
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 2: Analizar atributos completos desde clsOferta legacy
' TODO Sprint 2: Private _* As Tipo para cada atributo
' TODO Sprint 2: Property Get/Let para cada atributo
' TODO Sprint 3: Public Function IsValid() As Boolean
