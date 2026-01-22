VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpportunityManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("3-Services.OpportunityManagement")
Option Explicit

' ===============================================================================
' CLASE: clsOpportunityManager
' ===============================================================================
' PROPOSITO:
'   Servicio que gestiona la colección de oportunidades comerciales.
'   Lee carpetas desde ruta base configurada, mantiene índice de oportunidades,
'   y gestiona la oportunidad actual seleccionada.
'
'   Provee operaciones CRUD sobre oportunidades.
'   NO dispara eventos custom (reemplazados por llamadas directas).
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   USA:
'     - clsOpportunity (modelo de dominio)
'     - clsApplicationConfiguration (para obtener ruta base)
'     - FileSystemObject (para leer carpetas)
'
'   USADO POR:
'     - clsApplication (llama RefreshList, SetCurrentOpportunity)
'     - clsFileManager (para asociar archivos con oportunidades)
'     - modRibbonCallbacks (para obtener lista de oportunidades)
'
'   RECIBE callbacks desde:
'     - clsApplication cuando FileSystemMonitor detecta cambios
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - PULL: Clientes llaman GetOpportunityByIndex, CurrentOpportunity
'   - PUSH: Clientes llaman RefreshList, SetCurrentOpportunity
'   - NO WithEvents (eliminado)
'   - NO eventos custom (eliminados)
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' ===============================================================================
' Archivo origen: clsOpportunitiesMgr.cls (líneas 1-347)
'
' F-001: Configurar carpeta base
'   - Legacy: SetBaseFolder(path) (línea 45)
'   - Migra: Public Sub SetBaseFolder(path As String)
'
' F-002: Refrescar lista de oportunidades
'   - Legacy: RefreshList() (línea 67)
'   - Migra: Public Sub RefreshList()
'   - Lee carpetas de ruta base, crea clsOpportunity por cada una
'
' F-003: Obtener colección de oportunidades
'   - Legacy: Property Get Opportunities As Collection (línea 89)
'   - Migra: Public Function GetOpportunities() As Collection
'
' F-004: Obtener oportunidad actual
'   - Legacy: Property Get CurrentOpportunity As clsOpportunity (línea 101)
'   - Migra: Public Function GetCurrentOpportunity() As clsOpportunity
'
' F-005: Establecer oportunidad actual
'   - Legacy: Property Let CurrentIndex(value) (línea 119)
'   - CAMBIA: Public Sub SetCurrentOpportunity(index As Long)
'   - Notifica a clsApplication con llamada directa (no evento)
'
' F-006: Obtener oportunidad por índice
'   - Legacy: GetOpportunityByIndex(idx) (línea 145)
'   - Migra: Public Function GetOpportunityByIndex(idx As Long) As clsOpportunity
'
' F-007: Buscar oportunidad por nombre
'   - Legacy: NO EXISTE
'   - NUEVO: Public Function FindOpportunityByName(name) As clsOpportunity
'
' F-008: Colección privada
'   - Legacy: Private p_ColOpportunities As Collection (línea ~15)
'   - Migra: Private _opportunities As Collection
'
' F-009: Índice actual
'   - Legacy: Private p_CurrentIndex As Long (línea ~16)
'   - Migra: Private _currentIndex As Long
'
' F-010: Ruta base
'   - Legacy: Private p_BaseFolder As String (línea ~17)
'   - Migra: Private _baseFolder As String
'
' ELIMINACIONES:
'   - ❌ Public Event currOpportunityChanged (línea ~12)
'   - ❌ Public Event OpportunityCollectionUpdate (línea ~13)
'   - Reemplazados por: llamadas directas a clsApplication cuando cambia
'
' F-011: Actualizar colección desde eventos FileSystem
'   - Legacy: NO EXISTE explícitamente
'   - NUEVO: Public Sub OnOpportunityCreated(folderName)
'   - NUEVO: Public Sub OnOpportunityDeleted(folderName)
'   - NUEVO: Public Sub OnOpportunityRenamed(oldName, newName)
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 2: Private _opportunities As Collection
' TODO Sprint 2: Private _currentIndex As Long
' TODO Sprint 2: Private _baseFolder As String
' TODO Sprint 2: Private _config As clsApplicationConfiguration
' TODO Sprint 2: Public Sub SetBaseFolder(path)
' TODO Sprint 2: Public Sub RefreshList()
' TODO Sprint 2: Public Function GetOpportunities() As Collection
' TODO Sprint 2: Public Function GetCurrentOpportunity() As clsOpportunity
' TODO Sprint 2: Public Sub SetCurrentOpportunity(index As Long)
' TODO Sprint 2: Public Function GetOpportunityByIndex(idx) As clsOpportunity
' TODO Sprint 3: Public Sub OnOpportunityCreated/Deleted/Renamed
