VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOfferRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("5-Data")
Option Explicit

' ===============================================================================
' CLASE: clsOfferRepository
' ===============================================================================
' PROPOSITO:
'   Repositorio que gestiona persistencia de ofertas (CRUD).
'   Abstrae la fuente de datos (Excel, base de datos, archivos).
'
'   Implementa patrón Repository para desacoplar dominio de infraestructura.
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   USA:
'     - clsOffer (entidad de dominio)
'     - clsOtherOffer (entidad de dominio)
'     - clsDBContext (contexto de base de datos si aplica)
'     - Excel.Worksheet (acceso directo a hojas)
'
'   USADA POR:
'     - clsOpportunityManager (para leer/escribir ofertas)
'     - Commands de aplicación (CmdCreateOffer, etc)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - PULL: Clientes llaman GetOfferById, GetAllOffers
'   - PUSH: Clientes llaman SaveOffer, DeleteOffer
'   - Sin eventos
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' ===============================================================================
' Archivo origen: clsOfertaRepository.cls (líneas 1-149)
'
' F-001: Obtener oferta por ID
'   - Legacy: GetOfferById(id) As clsOferta (línea ~20)
'   - Migra: Public Function GetOfferById(id As String) As clsOffer
'
' F-002: Obtener todas las ofertas
'   - Legacy: GetAllOffers() As Collection (línea ~40)
'   - Migra: Public Function GetAllOffers() As Collection
'
' F-003: Guardar oferta
'   - Legacy: SaveOffer(offer As clsOferta) (línea ~60)
'   - Migra: Public Sub SaveOffer(offer As clsOffer)
'   - Persiste a Excel o base de datos
'
' F-004: Eliminar oferta
'   - Legacy: DeleteOffer(id As String) (línea ~80)
'   - Migra: Public Sub DeleteOffer(id As String)
'
' F-005: Buscar ofertas por criterios
'   - Legacy: FindOffersByType(offerType) As Collection (línea ~100)
'   - Migra: Public Function FindOffersByType(offerType As OfferTypeEnum) As Collection
'
' F-006: Crear nueva oferta
'   - Legacy: CreateNewOffer() As clsOferta (línea ~120)
'   - Migra: Public Function CreateNewOffer() As clsOffer
'   - Factory method para crear instancia vacía
'
' F-007: Validar existencia
'   - Legacy: OfferExists(id) As Boolean (línea ~140)
'   - Migra: Public Function OfferExists(id As String) As Boolean
'
' NOTA: Implementación depende de si se usa Excel como BD o base de datos externa.
'       Evaluar en Sprint 2.
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 2: Private _dbContext As clsDBContext (si aplica)
' TODO Sprint 2: Public Function GetOfferById(id) As clsOffer
' TODO Sprint 2: Public Function GetAllOffers() As Collection
' TODO Sprint 2: Public Sub SaveOffer(offer As clsOffer)
' TODO Sprint 2: Public Sub DeleteOffer(id As String)
' TODO Sprint 3: Public Function FindOffersByType(offerType) As Collection
' TODO Sprint 3: Public Function CreateNewOffer() As clsOffer
' TODO Sprint 3: Public Function OfferExists(id) As Boolean
