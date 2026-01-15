Attribute VB_Name = "modCALLBACKSRibbon"
' Modulo de integracion con la Ribbon: gestiona visibilidad y ejecucion de macros
' ==========================================
' MAPEO DE APIs (version anterior -> version refactorizada):
'   App.Ribbon.OnXXX          -> App.Events.RibbonEvents.OnXXX
'   App.Ribbon.InvalidarXXX   -> App.RibbonUI.InvalidarXXX
'   App.Ribbon.GetRibbonControlEnabled -> App.Events.RibbonEvents.GetRibbonControlEnabled
' ==========================================

'@Folder "2-Servicios.Excel.Ribbon"
'@IgnoreModule ProcedureNotUsed
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "modCALLBACKSRibbon"


' ==========================================
' CALLBACK: Se llama al cargar el Ribbon
' ==========================================
Sub RibbonOnLoad(xlRibbon As IRibbonUI)
Attribute RibbonOnLoad.VB_ProcData.VB_Invoke_Func = " \n0"
    LogInfo MODULE_NAME, "[callback: RibbonOnLoad] - Inicio"
    On Error GoTo ErrorHandler

    ' Inicializamos la referencia al ribbon en la aplicacion
    App.RibbonUI.Init xlRibbon

    LogInfo MODULE_NAME, "[callback: RibbonOnLoad] - ribbon cargado en la interfaz de excel"
    App.RibbonUI.InvalidarRibbon

    Exit Sub
ErrorHandler:
    LogError MODULE_NAME, "[callback: RibbonOnLoad] - Error", , Err.Description
End Sub

' ==========================================
' CALLBACKS DE MACROS
' ==========================================

Sub OnCompararHojas(control As IRibbonControl)
Attribute OnCompararHojas.VB_ProcData.VB_Invoke_Func = " \n0"
    MostrarComparador
End Sub

Sub OnDirtyRecalc(control As IRibbonControl)
Attribute OnDirtyRecalc.VB_ProcData.VB_Invoke_Func = " \n0"
    AplicarDirtyATodasLasHojasConFormulas
End Sub

Sub OnEvalUDFs(control As IRibbonControl)
Attribute OnEvalUDFs.VB_ProcData.VB_Invoke_Func = " \n0"
    ReemplazarUDFsEnFormulas
End Sub

Public Sub OnChangeAlturaFilas(control As IRibbonControl)
Attribute OnChangeAlturaFilas.VB_ProcData.VB_Invoke_Func = " \n0"
    LogInfo MODULE_NAME, "[callback: OnChangeAlturaFilas]"
    On Error GoTo Finalizar
    Call AjustarAltoFilasSegunColor
Finalizar:
End Sub

Public Sub OnMakeEditableBook(control As IRibbonControl)
Attribute OnMakeEditableBook.VB_ProcData.VB_Invoke_Func = " \n0"
    Call modMACROWbkEditableCleaning.LimpiarLibroActual
End Sub

Public Sub OnFitForPrint(control As IRibbonControl)
Attribute OnFitForPrint.VB_ProcData.VB_Invoke_Func = " \n0"
    Call modMACROWbkEditableFormatting.AjustarSelWSheetsParaImpresionPDF
End Sub

Public Sub OnVBAExport(control As IRibbonControl)
Attribute OnVBAExport.VB_ProcData.VB_Invoke_Func = " \n0"
    Call modMACROImportExportMacros.ExportarComponentesVBA
End Sub

Public Sub OnVBAImport(control As IRibbonControl)
Attribute OnVBAImport.VB_ProcData.VB_Invoke_Func = " \n0"
    Call modMACROImportExportMacros.ImportarComponentesVBA
End Sub

Public Sub OnOpenLog(control As IRibbonControl)
Attribute OnOpenLog.VB_ProcData.VB_Invoke_Func = " \n0"
    Call mod_Logger.AbrirLog
End Sub

Public Sub OnVBABackup(control As IRibbonControl)
Attribute OnVBABackup.VB_ProcData.VB_Invoke_Func = " \n0"
    Call CrearBackupCodigoVBA
    LogInfo MODULE_NAME, "[callback: OnVBABackup] - Creada copia de seguridad del codigo en " & ThisWorkbook.Path & "\Backups"
    MsgBox "Creada copia de seguridad del codigo en " & _
            ThisWorkbook.Path & "\Backups", vbInformation, "Copia de seguridad"
End Sub

Public Sub OnProcMetadataSync(control As IRibbonControl)
Attribute OnProcMetadataSync.VB_ProcData.VB_Invoke_Func = " \n0"
    Call modMACROProceduresToWorksheet.WriteProcedimientosSheet_ConBackup
End Sub

Public Sub OnToggleXLAMVisibility(control As IRibbonControl)
Attribute OnToggleXLAMVisibility.VB_ProcData.VB_Invoke_Func = " \n0"
    LogInfo MODULE_NAME, "[callback: OnToggleXLAMVisibility]"
    ThisWorkbook.IsAddin = Not (ThisWorkbook.IsAddin)
End Sub

' ==========================================
' CALLBACKS DE APLICACION (disparan eventos via RibbonEvents)
' ==========================================
Public Sub OnGenerarGraficosDesdeCurvasRto(control As IRibbonControl)
Attribute OnGenerarGraficosDesdeCurvasRto.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Events.RibbonEvents.OnGenerarGraficosDesdeCurvasRto
End Sub

Public Sub OnInvertirEjes(control As IRibbonControl)
Attribute OnInvertirEjes.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Events.RibbonEvents.OnInvertirEjes
End Sub

Public Sub OnFormatearCGASING(control As IRibbonControl)
Attribute OnFormatearCGASING.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Events.RibbonEvents.OnFormatearCGASING
End Sub

Public Sub OnNuevaOportunidad(control As IRibbonControl)
Attribute OnNuevaOportunidad.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Events.RibbonEvents.OnNuevaOportunidad
End Sub

Public Sub OnReplaceWithNamesInValidations(control As IRibbonControl)
Attribute OnReplaceWithNamesInValidations.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Events.RibbonEvents.OnReplaceWithNamesInValidations
End Sub

'--------------------------------------------------------------
' CALLBACKS DE CONFIGURACION
'--------------------------------------------------------------

' Callback del boton de configuracion
Sub OnConfigurador(control As IRibbonControl)
Attribute OnConfigurador.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Events.RibbonEvents.OnMostrarConfigurador
End Sub

'--------------------------------------------------------------
' CALLBACKS DEL DROPDOWN DE OPORTUNIDADES
'--------------------------------------------------------------

'--------------------------------------------------------------
' @Description: Callback del boton de refresco de oportunidades.
' Refresca el listado de subcarpetas y actualiza el desplegable
' del Ribbon.
'--------------------------------------------------------------
Public Sub CallbackRefrescarOportunidades(control As IRibbonControl)
Attribute CallbackRefrescarOportunidades.VB_ProcData.VB_Invoke_Func = " \n0"
    LogInfo MODULE_NAME, "[callback: CallbackRefrescarOportunidades] - control de ribbon activado para actualizar la lista de oportunidades"
    App.OpportunitiesMgr.actualizarColeccionOportunidades
    App.RibbonUI.InvalidarRibbon
End Sub

'--------------------------------------------------------------
' @Description: Devuelve el numero de oportunidades disponibles.
'--------------------------------------------------------------
Sub GetOportunidadesCount(control As IRibbonControl, ByRef returnedVal)
Attribute GetOportunidadesCount.VB_ProcData.VB_Invoke_Func = " \n0"
    returnedVal = App.OpportunitiesMgr.numOpportunities
End Sub

'--------------------------------------------------------------
' @Description: Devuelve la etiqueta de cada oportunidad.
'--------------------------------------------------------------
Sub GetOportunidadesLabel(control As IRibbonControl, Index As Integer, ByRef label)
Attribute GetOportunidadesLabel.VB_ProcData.VB_Invoke_Func = " \n0"
    label = App.OpportunitiesMgr.OportunityLabel(Index)
End Sub

'--------------------------------------------------------------
' @Description: Gestiona el evento de seleccion de oportunidad.
'--------------------------------------------------------------
Sub OnOportunidadesSeleccionada(control As IRibbonControl, id As String, Index As Integer)
Attribute OnOportunidadesSeleccionada.VB_ProcData.VB_Invoke_Func = " \n0"
    LogInfo MODULE_NAME, "[callback: OnOportunidadesSeleccionada] - modificada la oportunidad seleccionada en el control de ribbon"
    App.OpportunitiesMgr.CurrOpportunity = Index
    ' invalidar, refrescar el UI
    App.RibbonUI.InvalidarControl "ddlOportunidades"
End Sub

' Indice del elemento seleccionado
Sub GetSelectedOportunidadIndex(control As IRibbonControl, ByRef Index)
Attribute GetSelectedOportunidadIndex.VB_ProcData.VB_Invoke_Func = " \n0"
    Index = App.OpportunitiesMgr.CurrOpportunity
End Sub

' ==========================================
' CALLBACKS DE SUPERTIPS DINAMICOS
' ==========================================
Sub GetSupertipRutaBaseOportunidades(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBaseOportunidades.VB_ProcData.VB_Invoke_Func = " \n0"
    returnedVal = GetSupertipRutaBase(App.Configuration.RutaOportunidades)
End Sub

Sub GetSupertipRutaBasePlantillas(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBasePlantillas.VB_ProcData.VB_Invoke_Func = " \n0"
    returnedVal = GetSupertipRutaBase(App.Configuration.RutaPlantillas)
End Sub

Sub GetSupertipRutaBaseOfergas(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBaseOfergas.VB_ProcData.VB_Invoke_Func = " \n0"
    returnedVal = GetSupertipRutaBase(App.Configuration.RutaOfergas)
End Sub

Sub GetSupertipRutaBaseGasVBNet(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBaseGasVBNet.VB_ProcData.VB_Invoke_Func = " \n0"
    returnedVal = GetSupertipRutaBase(App.Configuration.RutaGasVBNet)
End Sub

Sub GetSupertipRutaBaseCalcTmpl(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBaseCalcTmpl.VB_ProcData.VB_Invoke_Func = " \n0"
    returnedVal = GetSupertipRutaBase(App.Configuration.RutaExcelCalcTempl)
End Sub

' Para mostrar la ruta actual en el supertip (dinamico)
Function GetSupertipRutaBase(ByVal strSettingRuta As String)
Attribute GetSupertipRutaBase.VB_ProcData.VB_Invoke_Func = " \n0"
    If strSettingRuta = "" Then strSettingRuta = "No configurada"
    GetSupertipRutaBase = "Ruta actual: " & strSettingRuta & vbCrLf & "Haz clic para cambiar..."
End Function

' ==========================================
' CALLBACKS GetEnabled (habilitar/deshabilitar controles)
' ==========================================
' Habilita el boton de grafico si el fichero es valido y cumple condiciones internas
Public Sub GetGraficoEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetGraficoEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = App.Events.RibbonEvents.GetRibbonControlEnabled(control)
End Sub

' Habilita el boton de inversion de ejes si hay grafico valido en contexto
Public Sub GetInvertirEjesEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetInvertirEjesEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = App.Events.RibbonEvents.GetRibbonControlEnabled(control)
End Sub

' Habilita el boton de procesado C-GAS-ING si hoja valida en contexto
Public Sub GetCGASINGEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetCGASINGEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = App.Events.RibbonEvents.GetRibbonControlEnabled(control)
End Sub

' Habilita el boton de creacion de nuevas oportunidades
Public Sub GetNuevaOportunidadEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetNuevaOportunidadEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = App.Events.RibbonEvents.GetRibbonControlEnabled(control)
End Sub

' Habilita el boton de cumplimentacion de oferta FULL si hoja valida en contexto
Public Sub GetOfertaFullEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetOfertaFullEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = True
End Sub

Public Sub GetOpenLogEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetOpenLogEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = GetLogFilePath <> ""
End Sub

' Habilita el boton del menu contextual del Ribbon si el fichero tiene nombre valido
Public Sub GetMenuEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetMenuEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = EsFicheroOportunidad()
    enabled = True
End Sub

' ==========================================
' CALLBACKS GetLabel (cambia la etiqueta de controles)
' ==========================================
Public Sub GetLabelToggleXLAM(control As IRibbonControl, ByRef returnedVal)
Attribute GetLabelToggleXLAM.VB_ProcData.VB_Invoke_Func = " \n0"
    If ThisWorkbook.IsAddin Then
        returnedVal = "Mostrar XLAM"
    Else
        returnedVal = "Ocultar XLAM"
    End If
End Sub

Public Sub GetLabelGrpConfiguracion(control As IRibbonControl, ByRef returnedVal)
Attribute GetLabelGrpConfiguracion.VB_ProcData.VB_Invoke_Func = " \n0"
    returnedVal = App.RibbonState.RibbonStateDescription
End Sub

' ==========================================
' CALLBACKS getVisible
' ==========================================
'@Description: Callback getVisible de la pestana "ABC"
Public Sub GetTabABCVisible(control As IRibbonControl, ByRef Visible)
Attribute GetTabABCVisible.VB_ProcData.VB_Invoke_Func = " \n0"
    If Not App() Is Nothing Then
        Visible = App.RibbonState.IsRibbonTabVisible
    Else
        Visible = False
    End If
End Sub

Public Sub GetGrpDeveloperAdminVisible(control As IRibbonControl, ByRef Visible)
Attribute GetGrpDeveloperAdminVisible.VB_ProcData.VB_Invoke_Func = " \n0"
    Visible = False
    If Not App() Is Nothing Then
        Visible = App.RibbonState.IsAdminGroupVisible
    End If
End Sub

