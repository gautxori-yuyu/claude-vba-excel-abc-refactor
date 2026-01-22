VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVBAProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("9-Utilities.VBAIntrospection")
Option Explicit

' ===============================================================================
' CLASE: clsVBAProcedure
' ===============================================================================
' PROPOSITO: Modelo para introspeccion y analisis de procedimientos VBA
' MIGRADO DESDE: main:clsVBAProcedure.cls (430 lineas)
' ===============================================================================
'
' RESPONSABILIDADES:
' - Parsear codigo fuente de procedimientos VBA
' - Extraer metadata (nombre, tipo, parametros, linea inicio/fin)
' - Analizar firma de procedimiento (Public/Private, Sub/Function)
' - Identificar dependencias entre procedimientos
'
' RELACIONES:
' - USADO POR: modMACROProceduresToWorksheet (exportacion de procedimientos)
' - USADO POR: modMACROListarProyectosVBA (analisis de proyectos)
' - USA: modUTILSProcedureParsing (utilidades de parseo)
'
' COMUNICACION:
' - PULL: Macros de analisis leen propiedades (Name, Type, LineStart, LineEnd)
' - PUSH: Macros crean instancias parseando codigo fuente
'
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' - clsVBAProcedure.cls (lineas 1-430):
'   * Property Get Name() As String
'   * Property Get ProcedureType() As vbext_ProcKind (Sub, Function, Property Get/Let/Set)
'   * Property Get Scope() As String (Public, Private, Friend)
'   * Property Get ModuleName() As String
'   * Property Get LineStart() As Long
'   * Property Get LineEnd() As Long
'   * Property Get ParameterCount() As Long
'   * Property Get Parameters() As Collection (clsParameter)
'   * Property Get SourceCode() As String
'   * Property Get Signature() As String
'   * Function HasParameter(paramName As String) As Boolean
'   * Function CallsOtherProcedure(procName As String) As Boolean
'   * Sub ParseFromSourceCode(source As String)
'
' PROPIEDADES CLAVE:
' - Name: Nombre del procedimiento
' - ProcedureType: Sub, Function, Property
' - Scope: Public, Private, Friend
' - SourceCode: Codigo completo del procedimiento
' - Signature: Firma completa (para documentacion)
'
' EVENTOS:
' - Ninguno (modelo de utilidad)
'
' TODO Sprint 4: Migrar implementacion desde main:clsVBAProcedure.cls
' ===============================================================================
