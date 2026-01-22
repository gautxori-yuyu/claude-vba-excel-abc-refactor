VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("2-Infrastructure.State")
Option Explicit

' ===============================================================================
' CLASE: clsFileState
' ===============================================================================
' PROPOSITO:
'   Almacena estado de archivos Excel:
'     - Archivo Excel activo (CurrentFile)
'     - Diccionario de archivos trackeados (TrackedFiles)
'     - Si el archivo activo es de oportunidad
'
'   Provee acceso mediante properties. NO dispara eventos.
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   USADO POR:
'     - clsFileManager (escribe estado)
'     - modRibbonCallbacks (lee para validar comandos)
'     - clsOpportunityManager (lee para asociar archivos)
'
'   CONTENIDO EN:
'     - clsApplicationState
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - PULL: Properties Get/Set
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' -------------------------------------------------------------------------------
' Archivo origen: clsFileState.cls (líneas 1-12) + clsFileManager
'
' F-001: Archivo Excel activo
'   - Legacy: clsFileManager.p_currentFile As clsExcelFile (línea ~20)
'   - Migra: Property CurrentFile As clsExcelFile
'
' F-002: Archivos trackeados
'   - Legacy: clsFileManager.p_trackedFiles As Object (línea ~19)
'   - Migra: Property TrackedFiles As Object (Dictionary)
'
' F-003: Validación de archivo oportunidad
'   - Legacy: Disperso en múltiples lugares
'   - Migra: Property IsOpportunityFile As Boolean (calculado)
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 2: Property CurrentFile As clsExcelFile
' TODO Sprint 2: Property TrackedFiles As Object
' TODO Sprint 2: Property IsOpportunityFile As Boolean
