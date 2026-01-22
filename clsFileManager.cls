VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("3-Services.FileManagement")
Option Explicit

' ===============================================================================
' CLASE: clsFileManager
' ===============================================================================
' PROPOSITO:
'   Servicio que gestiona tracking de archivos Excel abiertos.
'   Mantiene índice de archivos supervisados y sincroniza con el archivo activo.
'
'   Provee análisis de archivos (¿es fichero oportunidad?, asociar con oportunidad).
'   Actualiza clsFileState con archivo activo.
'
'   NO escucha eventos con WithEvents (recibe callbacks directos).
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   USA:
'     - clsApplicationState.FileState (escribe estado)
'     - clsExcelFile (modelo de archivo)
'     - clsOpportunityManager (para asociar archivos)
'
'   USADO POR:
'     - clsApplication (llama OnWorkbookOpened, OnWorkbookClosed)
'
'   RECIBE callbacks desde:
'     - clsApplication.OnWorkbookOpened(wb)
'     - clsApplication.OnWorkbookBeforeClose(wb, cancel)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - PULL: Lee estado de clsApplicationState.FileState
'   - PUSH: Escribe estado en clsApplicationState.FileState
'   - CALLBACKS: Recibe de clsApplication
'   - NO WithEvents (eliminado)
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' ===============================================================================
' Archivo origen: clsFileManager.cls (líneas 1-378)
'
' F-001: Inicialización con contexto
'   - Legacy: Initialize(execCtx As clsExecutionContext) (línea 45)
'   - Legacy: Private WithEvents ctx As clsExecutionContext (línea 18)
'   - CAMBIA: Public Sub Initialize(appState As clsApplicationState)
'   - ELIMINA: WithEvents (reemplazado por callbacks directos)
'
' F-002: Tracking de archivo
'   - Legacy: TrackFile(wb As Workbook) (línea 156)
'   - Migra: Public Sub TrackFile(wb As Workbook)
'   - Crea clsExcelFile y agrega a FileState.TrackedFiles
'
' F-003: Dejar de trackear archivo
'   - Legacy: UntrackFile(key As String) (línea 189)
'   - Migra: Public Sub UntrackFile(key As String)
'   - Elimina de FileState.TrackedFiles
'
' F-004: Obtener archivo por clave
'   - Legacy: GetFileByKey(key) As clsExcelFile (línea 134)
'   - Migra: Public Function GetFileByKey(key) As clsExcelFile
'
' F-005: Callback de WorkbookActivated
'   - Legacy: ctx_WorkbookActivated(wb) línea 89 (WithEvents)
'   - CAMBIA: Public Sub OnWorkbookActivated(wb As Workbook)
'   - Actualiza FileState.CurrentFile
'
' F-006: Callback de WorkbookBeforeClose
'   - Legacy: ctx_WorkbookBeforeClose(wb, cancel) línea 112 (WithEvents)
'   - CAMBIA: Public Sub OnWorkbookBeforeClose(wb As Workbook, cancel As Boolean)
'   - Llama UntrackFile
'
' F-007: Propiedades de acceso (eliminadas - usar FileState)
'   - Legacy: Property Get CurrentFile As clsExcelFile (línea 67)
'   - Legacy: Property Get TrackedFiles As Object (línea 78)
'   - CAMBIA: Acceso directo a FileState (no propiedades duplicadas)
'
' ELIMINACIONES:
'   - ❌ Private WithEvents ctx As clsExecutionContext (línea 18)
'   - ❌ Property Get CurrentFile/TrackedFiles (usar FileState)
'   - Reemplazados por: callbacks directos + acceso a FileState
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 2: Private _appState As clsApplicationState
' TODO Sprint 2: Private _opportunityManager As clsOpportunityManager
' TODO Sprint 2: Public Sub Initialize(appState, oppManager)
' TODO Sprint 2: Public Sub TrackFile(wb As Workbook)
' TODO Sprint 2: Public Sub UntrackFile(key As String)
' TODO Sprint 2: Public Function GetFileByKey(key) As clsExcelFile
' TODO Sprint 2: Public Sub OnWorkbookActivated(wb)
' TODO Sprint 2: Public Sub OnWorkbookBeforeClose(wb, cancel)
