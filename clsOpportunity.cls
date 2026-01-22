VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpportunity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("4-Domain.Entities")
Option Explicit

' ===============================================================================
' CLASE: clsOpportunity
' ===============================================================================
' PROPOSITO:
'   Entidad de dominio que representa una oportunidad comercial.
'   Modelo simple con:
'     - Nombre de la oportunidad
'     - Ruta de la carpeta asociada
'     - Métodos de validación de negocio (si aplican)
'
'   DOMINIO PURO: Sin dependencias de infraestructura.
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   USADA POR:
'     - clsOpportunityManager (crea y gestiona colección)
'     - clsExcelFile (asocia archivos con oportunidades)
'
'   USA:
'     - Ninguna (entidad pura de dominio)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - PULL: Clientes leen properties (Nombre, Path)
'   - Sin eventos
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' ===============================================================================
' Archivo origen: clsOpportunity.cls (líneas 1-41)
'
' F-001: Atributos de oportunidad
'   - Legacy: Private mPath As String (línea ~10)
'   - Legacy: Private mNombre As String (línea ~11)
'   - Migra: Private _path As String
'   - Migra: Private _name As String
'
' F-002: Propiedades de acceso
'   - Legacy: Property Get/Let Path (líneas 15-25)
'   - Legacy: Property Get/Let Nombre (líneas 27-37)
'   - Migra: Public Property Get/Let Path As String
'   - Migra: Public Property Get/Let Name As String
'
' F-003: Validación de nombre
'   - Legacy: NO EXISTE
'   - NUEVO: Public Function IsValidName() As Boolean
'   - Valida que nombre no esté vacío, no contenga caracteres inválidos
'
' F-004: Comparación de oportunidades
'   - Legacy: NO EXISTE
'   - NUEVO: Public Function Equals(other As clsOpportunity) As Boolean
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 2: Private _path As String
' TODO Sprint 2: Private _name As String
' TODO Sprint 2: Property Get/Let Path As String
' TODO Sprint 2: Property Get/Let Name As String
' TODO Sprint 3: Public Function IsValidName() As Boolean
' TODO Sprint 3: Public Function Equals(other) As Boolean
