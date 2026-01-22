VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRefEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("6-Presentation.Forms.Controls")
Option Explicit

' ===============================================================================
' CLASE: CRefEdit
' ===============================================================================
' PROPOSITO: Wrapper para control RefEdit personalizado
' MIGRADO DESDE: main:CRefEdit.cls (79 lineas)
' ===============================================================================
'
' RESPONSABILIDADES:
' - Encapsular funcionalidad del control RefEdit
' - Proveer interfaz simplificada para formularios
' - Gestionar validacion de rangos seleccionados
'
' RELACIONES:
' - USADO POR: Formularios que necesitan seleccion de rangos (frmConfiguracion)
' - USA: API de RefEdit (modUTILSRefEditAPI)
'
' COMUNICACION:
' - PULL: Formularios leen propiedades (Text, Value, IsValid)
' - PUSH: Formularios configuran propiedades
'
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' - CRefEdit.cls (lineas 1-79):
'   * Inicializacion del control
'   * Validacion de rangos
'   * Gestion de eventos del control
'   * Conversion entre texto y Range
'
' EVENTOS:
' - Ninguno (control simple, sin WithEvents)
'
' TODO Sprint 2: Migrar implementacion desde main:CRefEdit.cls
' ===============================================================================
