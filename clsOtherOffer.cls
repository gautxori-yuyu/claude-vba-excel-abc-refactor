VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOtherOffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("4-Domain.Entities")
Option Explicit

' ===============================================================================
' CLASE: clsOtherOffer
' ===============================================================================
' PROPOSITO:
'   Entidad de dominio que representa una oferta de otro tipo
'   (no compresor estándar).
'
'   Especialización o tipo alternativo de oferta.
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
'     - modOfferTypes (tipos de ofertas)
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
' Archivo origen: clsOfertaOtro.cls (líneas 1-37)
'
' F-001: Atributos específicos de oferta otro tipo
'   - Legacy: Private m* As Variant (líneas ~10-15)
'   - Migra: Atributos específicos (diferencia con clsOffer estándar)
'
' F-002: Propiedades de acceso
'   - Legacy: Property Get/Let (líneas ~17-35)
'   - Migra: Public Property Get/Let para cada atributo
'
' NOTA: Evaluar si esta clase puede fusionarse con clsOffer
'       usando un atributo "Type" en lugar de clase separada.
'       Decisión en Sprint 2 tras analizar uso en legacy.
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 2: Analizar diferencias con clsOffer
' TODO Sprint 2: Decidir si fusionar o mantener separado
' TODO Sprint 2: Implementar atributos y properties
