VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("4-Domain.Entities")
Option Explicit

' ===============================================================================
' CLASE: clsExcelFile
' ===============================================================================
' PROPOSITO: Modelo de dominio para archivo Excel abierto
' MIGRADO DESDE: main:clsExcelFile.cls (412 lineas)
' IMPLEMENTA: IFile
' ===============================================================================
'
' RESPONSABILIDADES:
' - Representar archivo Excel con su contexto (Workbook, Path, Opportunity)
' - Proveer analisis de tipo de archivo (es oportunidad, es plantilla, etc.)
' - Encapsular metadata del archivo (nombre, ruta, oportunidad asociada)
' - Implementar contrato IFile para archivos Excel
'
' RELACIONES:
' - USADO POR: clsFileManagementService (tracking de archivos abiertos)
' - USA: clsOpportunity (para asociar archivo con oportunidad)
' - IMPLEMENTA: IFile
'
' COMUNICACION:
' - PULL: Servicios leen propiedades (Path, IsOpportunityFile, AssociatedOpportunity)
' - PUSH: Servicios crean instancias via constructor
'
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' - clsExcelFile.cls (lineas 1-412):
'   * Property Get/Let Workbook() As Workbook
'   * Property Get Path() As String
'   * Property Get FileName() As String
'   * Property Get Extension() As String
'   * Property Get IsOpportunityFile() As Boolean
'   * Property Get/Set AssociatedOpportunity() As clsOpportunity
'   * Function IsCGASINGSheet(ws As Worksheet) As Boolean
'   * Function GetCGASINGSheets() As Collection
'   * Property Get Exists() As Boolean (de IFile)
'   * Property Get LastModified() As Date (de IFile)
'   * Implementacion completa de IFile
'
' PROPIEDADES CLAVE:
' - Workbook: Referencia al objeto Workbook de Excel
' - Path: Ruta completa del archivo
' - AssociatedOpportunity: Oportunidad comercial asociada (si aplica)
' - IsOpportunityFile: Indica si el archivo pertenece a una oportunidad
'
' EVENTOS:
' - Ninguno (modelo de dominio pasivo)
'
' TODO Sprint 2: Migrar implementacion desde main:clsExcelFile.cls
' ===============================================================================
