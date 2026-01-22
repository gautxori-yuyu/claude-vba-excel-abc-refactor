VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wshUnidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("8-UserDefinedFunctions.Units")
Option Explicit

' ===============================================================================
' HOJA: wshUnidades
' ===============================================================================
' PROPOSITO: Hoja de calculo con tabla de conversiones de unidades
' MIGRADO DESDE: main:wshUnidades.cls (17 lineas)
' ===============================================================================
'
' RESPONSABILIDADES:
' - Contener tabla maestra de conversiones de unidades
' - Proveer datos a UDF ConvertirUnidad()
' - Gestionar eventos de cambio en la hoja (validaciones)
'
' RELACIONES:
' - USADO POR: UDFs_Units.bas (funcion ConvertirUnidad lee esta hoja)
' - USADO POR: modMACROUnits.bas (macros de gestion de unidades)
'
' COMUNICACION:
' - PULL: UDFs leen rangos de la hoja (factores de conversion)
' - EVENTOS: Worksheet_Change, Worksheet_Activate (validaciones, proteccion)
'
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' - wshUnidades.cls (lineas 1-17):
'   * Private Sub Worksheet_Change(ByVal Target As Range)
'     - Validar cambios en tabla de unidades
'     - Recalcular factores de conversion dependientes
'   * Private Sub Worksheet_Activate()
'     - Mostrar ayuda contextual
'   * Private Sub Worksheet_Deactivate()
'     - Guardar cambios pendientes
'
' ESTRUCTURA DE LA HOJA:
' - Columna A: Tipo de magnitud (Longitud, Masa, Temperatura, etc.)
' - Columna B: Unidad origen
' - Columna C: Unidad destino
' - Columna D: Factor de conversion
' - Columna E: Offset (para conversiones no lineales como temperatura)
'
' EVENTOS:
' - Worksheet_Change: Valida ediciones en tabla de unidades
' - Worksheet_Activate/Deactivate: Gestion UI
'
' TODO Sprint 4: Migrar codigo de eventos desde main:wshUnidades.cls
' TODO Sprint 4: Verificar estructura de tabla en hoja fisica
' ===============================================================================
