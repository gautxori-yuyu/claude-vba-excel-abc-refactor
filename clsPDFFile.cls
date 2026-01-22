VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPDFFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("4-Domain.Entities")
Option Explicit

' ===============================================================================
' CLASE: clsPDFFile
' ===============================================================================
' PROPOSITO: Modelo de dominio para archivo PDF
' MIGRADO DESDE: main:clsPDFFile.cls (100 lineas)
' IMPLEMENTA: IFile
' ===============================================================================
'
' RESPONSABILIDADES:
' - Representar archivo PDF con su metadata
' - Proveer operaciones especificas de archivos PDF
' - Encapsular informacion del documento (titulo, autor, paginas)
' - Implementar contrato IFile para archivos PDF
'
' RELACIONES:
' - USADO POR: Servicios que gestionan documentacion de oportunidades
' - IMPLEMENTA: IFile
'
' COMUNICACION:
' - PULL: Servicios leen propiedades (Path, Title, PageCount)
' - PUSH: Servicios crean instancias via constructor
'
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' - clsPDFFile.cls (lineas 1-100):
'   * Property Get Path() As String
'   * Property Get FileName() As String
'   * Property Get Extension() As String
'   * Property Get Title() As String (metadata PDF)
'   * Property Get Author() As String (metadata PDF)
'   * Property Get PageCount() As Long
'   * Property Get Exists() As Boolean (de IFile)
'   * Property Get LastModified() As Date (de IFile)
'   * Sub Delete() (de IFile)
'   * Function Copy(destination) (de IFile)
'   * Function Move(destination) (de IFile)
'   * Function Open() As Boolean (abre PDF con aplicacion predeterminada)
'
' PROPIEDADES CLAVE:
' - Path: Ruta completa del archivo PDF
' - Title: Titulo del documento (metadata)
' - PageCount: Numero de paginas del PDF
'
' EVENTOS:
' - Ninguno (modelo de dominio pasivo)
'
' TODO Sprint 3: Migrar implementacion desde main:clsPDFFile.cls
' ===============================================================================
