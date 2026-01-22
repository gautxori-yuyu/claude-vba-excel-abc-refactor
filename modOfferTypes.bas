Attribute VB_Name = "modOfferTypes"
'@Folder("4-Domain.ValueObjects")
Option Explicit

' ===============================================================================
' MODULO: modOfferTypes
' ===============================================================================
' PROPOSITO:
'   Value Objects de dominio: enumeraciones y constantes relacionadas
'   con tipos de ofertas.
'
'   Define los tipos válidos de ofertas en el sistema.
'   DOMINIO PURO: Sin dependencias de infraestructura.
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   USADO POR:
'     - clsOffer (para definir tipo de oferta)
'     - clsOtherOffer (si aplica)
'     - clsOfferRepository (para filtrar por tipo)
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' ===============================================================================
' Archivo origen: modOfertaTypes.bas (líneas 1-13)
'
' F-001: Enumeración de tipos de ofertas
'   - Legacy: Public Enum OfferTypeEnum (líneas ~5-10)
'   - Migra: Public Enum OfferTypeEnum
'   - Valores posibles (a confirmar en Sprint 2):
'     - otStandard = 0 (oferta estándar)
'     - otOther = 1 (oferta otro tipo)
'     - otCustom = 2 (oferta personalizada)
'     - (otros según análisis)
'
' F-002: Constantes de tipos
'   - Legacy: Public Const OFFER_TYPE_* (si existen)
'   - Migra: Public Const para cada tipo
'
' F-003: Helper de validación
'   - NUEVO: Public Function IsValidOfferType(value As Long) As Boolean
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 2: Analizar tipos reales de ofertas en legacy
' TODO Sprint 2: Public Enum OfferTypeEnum con valores correctos
' TODO Sprint 2: Public Const para cada tipo
' TODO Sprint 3: Public Function IsValidOfferType(value) As Boolean
