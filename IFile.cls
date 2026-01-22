VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("4-Domain.Entities")
'@Interface
Option Explicit

' ===============================================================================
' INTERFAZ: IFile
' ===============================================================================
' PROPOSITO: Contrato para clases que representan archivos
' MIGRADO DESDE: main:IFile.cls (56 lineas)
' ===============================================================================
'
' RESPONSABILIDADES:
' - Definir contrato comun para diferentes tipos de archivo
' - Garantizar interfaz uniforme para operaciones con archivos
' - Permitir polimorfismo en repositorios y servicios
'
' RELACIONES:
' - IMPLEMENTADO POR: clsExcelFile, clsPDFFile
' - USADO POR: Repositorios y servicios que trabajan con archivos abstractos
'
' COMUNICACION:
' - PULL: Consumidores leen propiedades via interfaz
' - PUSH: Consumidores invocan metodos via interfaz
'
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' - IFile.cls (lineas 1-56):
'   * Property Get Path() As String
'   * Property Get FileName() As String
'   * Property Get Extension() As String
'   * Property Get Exists() As Boolean
'   * Property Get LastModified() As Date
'   * Sub Delete()
'   * Function Copy(destination As String) As Boolean
'   * Function Move(destination As String) As Boolean
'
' IMPLEMENTADORES EN REFACTOR:
' - clsExcelFile (archivos .xlsx, .xlsm, .xls)
' - clsPDFFile (archivos .pdf)
'
' TODO Sprint 2: Migrar definicion de interfaz desde main:IFile.cls
' ===============================================================================
