VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "3.Interfaces"
'@Interface
Option Explicit

'=============================================================================
' IFormatter - Interfaz para formateadores
'=============================================================================
' Propósito: Estandarizar formatters (CGASING, Oferta, Reporte, etc.)
'
' Todos los formatters deben implementar esta interfaz con:
'   Implements IFormatter
'   Implements IService  ' Si quieres gestión de ciclo de vida
'
' Ejemplo de implementación:
'   Implements IFormatter
'
'   Private Function IFormatter_Format(ByVal target As Object) As Boolean
'       On Error GoTo ErrHandler
'
'       ' Validar que target es del tipo esperado
'       If Not IFormatter_CanFormat(target) Then
'           IFormatter_Format = False
'           Exit Function
'       End If
'
'       ' Tu lógica de formateo aquí
'       Dim wb As Workbook
'       Set wb = target
'       ' ... formatear ...
'
'       IFormatter_Format = True
'       Exit Function
'
'   ErrHandler:
'       IFormatter_Format = False
'   End Function
'
'   Private Function IFormatter_CanFormat(ByVal target As Object) As Boolean
'       ' Verificar si el target es válido para este formatter
'       On Error Resume Next
'       Dim wb As Workbook
'       Set wb = target
'       IFormatter_CanFormat = Not wb Is Nothing
'   End Function
'
'   Private Property Get IFormatter_FormatterName() As String
'       IFormatter_FormatterName = "CGASINGFormatter"
'   End Property
'=============================================================================

'@Description: Formatea el target (workbook, worksheet, range, etc.)
'@Param target: Objeto a formatear (Workbook, Worksheet, Range, etc.)
'@Returns: True si el formateo fue exitoso, False en caso contrario
Public Function Format(ByVal target As Object) As Boolean
    Err.Raise 5, "IFormatter", "Format debe implementarse en clase concreta"
End Function

'@Description: Verifica si puede formatear el target dado
'@Param target: Objeto a validar
'@Returns: True si este formatter puede procesar el target
Public Function CanFormat(ByVal target As Object) As Boolean
    Err.Raise 5, "IFormatter", "CanFormat debe implementarse en clase concreta"
End Function

'@Description: Nombre descriptivo del formateador
Public Property Get FormatterName() As String
    Err.Raise 5, "IFormatter", "FormatterName debe implementarse en clase concreta"
End Property

