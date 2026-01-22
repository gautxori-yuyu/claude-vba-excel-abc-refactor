VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("5-Data")
Option Explicit

' ===============================================================================
' CLASE: clsDBContext
' ===============================================================================
' PROPOSITO:
'   Contexto de acceso a base de datos (si aplica).
'   Gestiona conexiones a base de datos externa (SQL Server, Access, etc).
'
'   Si el sistema solo usa Excel como "base de datos", esta clase puede
'   ser un placeholder o eliminarseEn Sprint 2.
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   USA:
'     - ADODB.Connection (para conexión a BD)
'     - ADODB.Recordset (para queries)
'
'   USADA POR:
'     - clsOfferRepository (ejecuta queries)
'     - Otros repositorios (si existen)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - PULL: Repositorios llaman ExecuteQuery, ExecuteNonQuery
'   - Sin eventos
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' ===============================================================================
' Archivo origen: clsDBContext.cls (líneas 1-94)
'
' F-001: Conexión a base de datos
'   - Legacy: Private mConnection As ADODB.Connection (línea ~10)
'   - Legacy: OpenConnection() (línea ~20)
'   - Migra: Public Sub OpenConnection(connectionString As String)
'
' F-002: Cierre de conexión
'   - Legacy: CloseConnection() (línea ~40)
'   - Migra: Public Sub CloseConnection()
'
' F-003: Ejecutar query SELECT
'   - Legacy: ExecuteQuery(sql) As ADODB.Recordset (línea ~50)
'   - Migra: Public Function ExecuteQuery(sql As String) As ADODB.Recordset
'
' F-004: Ejecutar comando INSERT/UPDATE/DELETE
'   - Legacy: ExecuteNonQuery(sql) (línea ~70)
'   - Migra: Public Sub ExecuteNonQuery(sql As String)
'
' F-005: Validación de conexión
'   - Legacy: IsConnected() As Boolean (línea ~85)
'   - Migra: Public Function IsConnected() As Boolean
'
' NOTA: Evaluar en Sprint 2 si realmente se necesita base de datos externa
'       o si Excel es suficiente. Si Excel, esta clase puede eliminarse.
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 2: Evaluar si se necesita BD externa o solo Excel
' TODO Sprint 2: Si BD: Private _connection As ADODB.Connection
' TODO Sprint 2: Si BD: Public Sub OpenConnection(connectionString)
' TODO Sprint 2: Si BD: Public Sub CloseConnection()
' TODO Sprint 2: Si BD: Public Function ExecuteQuery(sql) As ADODB.Recordset
' TODO Sprint 2: Si BD: Public Sub ExecuteNonQuery(sql)
' TODO Sprint 2: Si Excel solo: Marcar clase como PLACEHOLDER o eliminar
