VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'TODO Migrar de Registry a archivo JSON para configuraci�n (NO PRIORITARIO)
'TODO Implementar importar/exportar configuraci�n (NO PRIORITARIO)

'@Folder "2-Servicios.Configuracion"
'@IgnoreModule ProcedureNotUsed
Option Explicit

'=============================================================================
' IMPLEMENTACION IService
'=============================================================================
Implements IService

Private mIsInitialized As Boolean

Private mRutaOportunidades As String
Private mRutaPlantillas As String
Private mRutaOfergas As String
Private mRutaGasVBNet As String
Private mRutaExcelCalcTempl As String
Private mArrComprImgs As Variant
Private mArrComprDrawPIDs As Variant
Private mSAM As Integer

Public Property Get RutaOportunidades() As String
    RutaOportunidades = mRutaOportunidades
End Property

Public Property Let RutaOportunidades(newRuta As String)
    mRutaOportunidades = newRuta
    GuardarRuta CFG_RUTA_OPORTUNIDADES, newRuta
End Property

Public Property Get RutaPlantillas() As String
    RutaPlantillas = mRutaPlantillas
End Property

Public Property Let RutaPlantillas(newRuta As String)
    mRutaPlantillas = newRuta
    GuardarRuta CFG_RUTA_PLANTILLAS, newRuta
End Property

Public Property Get RutaOfergas() As String
    RutaOfergas = mRutaOfergas
End Property

Public Property Let RutaOfergas(newRuta As String)
    mRutaOfergas = newRuta
    GuardarRuta CFG_RUTA_OFERGAS, newRuta
End Property

Public Property Get RutaGasVBNet() As String
    RutaGasVBNet = mRutaGasVBNet
End Property

Public Property Let RutaGasVBNet(newRuta As String)
    mRutaGasVBNet = newRuta
    GuardarRuta CFG_RUTA_GAS_VBNET, newRuta
End Property

Public Property Get RutaExcelCalcTempl() As String
    RutaExcelCalcTempl = mRutaExcelCalcTempl
End Property

Public Property Let RutaExcelCalcTempl(newRuta As String)
    mRutaExcelCalcTempl = newRuta
    GuardarRuta CFG_RUTA_PLCALCS, newRuta
End Property

Public Property Get ListComprImgs() As Variant
    ListComprImgs = mArrComprImgs
End Property

Public Property Let ListComprImgs(arrRutas As Variant)
    mArrComprImgs = arrRutas
    GuardarListaCarpetas arrRutas, CFG_RUTA_COMPRIMGS
End Property

Public Property Get ListComprDrawPIDs() As Variant
    ListComprDrawPIDs = mArrComprDrawPIDs
End Property

Public Property Let ListComprDrawPIDs(arrRutas As Variant)
    mArrComprDrawPIDs = arrRutas
    GuardarListaCarpetas arrRutas, CFG_RUTA_COMPRDRAWPID
End Property

Public Property Get SAM() As Integer
    SAM = mSAM
End Property

Public Property Let SAM(newSAM As Integer)
    mSAM = newSAM
    GuardarSAM newSAM
End Property

'=============================================================================
' CICLO DE VIDA (modificado para IService)
'=============================================================================

Private Sub Class_Initialize()
    ' NO cargar aqui - esperar a IService_Initialize
    LogInfo "clsConfiguration", "[Class_Initialize] - instancia creada"
End Sub

Private Sub Class_Terminate()
    LogInfo "clsConfiguration", "[Class_Terminate]"
End Sub

'-----------------------------------------------------------------------------
' IService Implementation
'-----------------------------------------------------------------------------

Private Sub IService_Initialize(ByVal dependencies As Object)
    If mIsInitialized Then Exit Sub
    
    LogInfo "clsConfiguration", "[IService_Initialize] - Cargando configuracion..."
    
    ' Cargar configuracion del registro
    Call CargarConfiguracionDesdeRegistro
    
    mIsInitialized = True
    
    LogInfo "clsConfiguration", "[IService_Initialize] - Configuracion cargada"
End Sub

Private Sub IService_Dispose()
    LogInfo "clsConfiguration", "[IService_Dispose] - Liberando recursos"
    mIsInitialized = False
End Sub

Private Property Get IService_IsInitialized() As Boolean
    IService_IsInitialized = mIsInitialized
End Property

Private Property Get IService_ServiceName() As String
    IService_ServiceName = "Configuration"
End Property

'-----------------------------------------------------------------------------
' resto del  codigo
'-----------------------------------------------------------------------------

Private Sub CargarConfiguracionDesdeRegistro()
    On Error GoTo ErrHandler
    
    ' Cargar cada ruta desde el registro y mostrarla
    LogInfo "clsConfiguration", "[CargarConfiguracionDesdeRegistro] - inicializacion de todas las rutas desde registro o valores por defecto"
    mRutaOportunidades = ObtenerRuta(CFG_RUTA_OPORTUNIDADES, CFG_RUTA_OPORTUNIDADES_DEFAULT)
    mRutaPlantillas = ObtenerRuta(CFG_RUTA_PLANTILLAS, CFG_RUTA_PLANTILLAS_DEFAULT)
    mRutaOfergas = ObtenerRuta(CFG_RUTA_OFERGAS, CFG_RUTA_OFERGAS_DEFAULT)
    mRutaGasVBNet = ObtenerRuta(CFG_RUTA_GAS_VBNET, CFG_RUTA_GAS_VBNET_DEFAULT)
    mRutaExcelCalcTempl = ObtenerRuta(CFG_RUTA_PLCALCS, CFG_RUTA_PLCALCS_DEFAULT)
    mArrComprImgs = ObtenerListaCarpetas(CFG_RUTA_COMPRIMGS, CFG_RUTA_COMPRIMGS_DEFAULT)
    mArrComprDrawPIDs = ObtenerListaCarpetas(CFG_RUTA_COMPRDRAWPID, CFG_RUTA_COMPRDRAWPID_DEFAULT)
    mSAM = ObtenerSAM()

    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "clsConfiguration.CargarConfiguracionDesdeRegistro", _
              "Error en la carga de configuraci�n desde registro: " & Err.Description
End Sub

'@Description: Devuelve diccionario de carpetas a monitorear (cacheado)
'@Note: La cache se invalida automaticamente cuando cambian las rutas
'
'TODO: revisar el uso de esta funci�n: Buscar otra forma de que el folder watcher tenga conocimiento de qu� carpetas son supervisadas;
' posiblemente presentando de otra manera el formulario de configuraci�n de carpetas, y creando en �l opciones para su selecci�n...
'TODO: Considerar refactorizar a una estructura mas clara:
' - Crear clase clsFolderWatchConfig con propiedades: TipoCarpeta, Ruta, EsMonitoreable
' - Exponer como Collection tipada en lugar de Dictionary
' - Permitir habilitar/deshabilitar monitoreo por carpeta desde UI
Public Property Get oDicFoldersToWatch() As Object
    Set oDicFoldersToWatch = CreateObject("scripting.dictionary")
    oDicFoldersToWatch.Add CFG_RUTA_OPORTUNIDADES, RutaOportunidades
    oDicFoldersToWatch.Add CFG_RUTA_PLANTILLAS, RutaPlantillas
    oDicFoldersToWatch.Add CFG_RUTA_OFERGAS, RutaOfergas
    oDicFoldersToWatch.Add CFG_RUTA_GAS_VBNET, RutaGasVBNet
    oDicFoldersToWatch.Add CFG_RUTA_PLCALCS, RutaExcelCalcTempl
End Property

'-------------------------------------------
' PERSISTENCIA (Registro de Windows)
'-------------------------------------------
Private Sub GuardarSAM(SAM As Integer)
    CreateObject("WScript.Shell").RegWrite CFG_PATH_SAM, CStr(SAM), "REG_SZ"
    LogDebug "clsConfiguration", "[GuardarSAM] - S.A.M. guardado en registro"
End Sub

Private Function ObtenerSAM() As Integer
    Dim valorLeido As Variant

    On Error Resume Next
    valorLeido = CreateObject("WScript.Shell").RegRead(CFG_PATH_SAM)
    Dim errNum As Long: errNum = Err.Number
    On Error GoTo 0

    ' Verificar si hubo error o valor invalido
    If errNum <> 0 Or IsEmpty(valorLeido) Or Not IsNumeric(valorLeido) Then
        LogWarning "clsConfiguration", "[ObtenerSAM] - no existente en registro o invalido, establecido valor por defecto"
        ObtenerSAM = CFG_SAM
        GuardarSAM CFG_SAM
    ElseIf CInt(valorLeido) = 0 Then
        LogWarning "clsConfiguration", "[ObtenerSAM] Valor cero en registro, usando valor por defecto"
        ObtenerSAM = CFG_SAM
        GuardarSAM CFG_SAM
    Else
        ObtenerSAM = CInt(valorLeido)
        LogDebug "clsConfiguration", "[ObtenerSAM] Leido del registro: " & ObtenerSAM
    End If
End Function


' Guardar una ruta en el registro
Private Sub GuardarRuta(nombreConfig As Variant, ruta As String)
    'SaveSetting APP_NAME & "\" & CFG_SECTION_RUTAS, nombreConfig, "", ruta
    CreateObject("WScript.Shell").RegWrite CFG_PATH_SECTION_RUTAS & nombreConfig & "\", ruta, "REG_SZ"
    LogDebug "clsConfiguration", "[GuardarRuta] - ruta guardada en registro"
End Sub

' Recuperar una ruta del registro
Private Function ObtenerRuta(nombreConfig As Variant, strDefault As String) As String
    'ObtenerRuta = GetSetting(APP_NAME & "\" & CFG_SECTION_RUTAS, nombreConfig, "", "")
    On Error Resume Next
    ObtenerRuta = CreateObject("WScript.Shell").RegRead(CFG_PATH_SECTION_RUTAS & nombreConfig & "\")
    If ObtenerRuta = "" Or Not RutaExiste(ObtenerRuta) Then
        ObtenerRuta = strDefault
        GuardarRuta nombreConfig, ObtenerRuta
        LogWarning "clsConfiguration", "[ObtenerRuta] - ruta no existente en registro o invalida, establecido valor por defecto"
    Else
        LogDebug "clsConfiguration", "[ObtenerRuta] - ruta leida del registro"
    End If
    On Error GoTo 0
End Function

' Guarda la lista de Carpetas
Private Sub GuardarListaCarpetas(ByVal carpetas As Variant, nombreConfig As String)
    On Error GoTo ErrHandler

    Dim lista As String
    Dim key As Variant

    ' Validar array
    If Not IsArray(carpetas) Then
        LogWarning "clsConfiguration", "[GuardarListaCarpetas] - No es un array, borrando lista"
        BorrarListaCarpetas nombreConfig
        Exit Sub
    End If

    ' Verificar si el array esta vacio
    On Error Resume Next
    Dim upperBound As Long
    upperBound = UBound(carpetas)
    If Err.Number <> 0 Or upperBound < 0 Then
        On Error GoTo ErrHandler
        LogWarning "clsConfiguration", "[GuardarListaCarpetas] - Array vacio, borrando lista"
        BorrarListaCarpetas nombreConfig
        Exit Sub
    End If
    On Error GoTo ErrHandler

    ' Construir string concatenado
    For Each key In carpetas
        If Len(lista) > 0 Then lista = lista & "|"
        lista = lista & CStr(key)
    Next key

    ' Guardar en registro
    GuardarRuta nombreConfig, lista
    LogDebug "clsConfiguration", "[GuardarListaCarpetas] - Lista guardada: " & nombreConfig

    Exit Sub

ErrHandler:
    LogError "clsConfiguration", "[GuardarListaCarpetas] - Error guardando lista: " & nombreConfig, Err.Number, Err.Description
End Sub

' Obtiene la lista de Carpetas
Private Function ObtenerListaCarpetas(nombreConfig As String, strDefault As String) As Variant
    On Error Resume Next
    ObtenerListaCarpetas = Split(CreateObject("WScript.Shell").RegRead(CFG_PATH_SECTION_RUTAS & nombreConfig & "\"), "|")
    
    If Err.Number <> 0 Then
        LogWarning "clsConfiguration", "[ObtenerListaCarpetas] - No se encontr� lista guardada en registro."
        ObtenerListaCarpetas = Split(strDefault, "|")
        Call GuardarListaCarpetas(ObtenerListaCarpetas, nombreConfig)
    Else
        LogDebug "clsConfiguration", "[ObtenerListaCarpetas] - Lista de carpetas leida desde registro."
    End If
    On Error GoTo 0
End Function

' Borra la lista guardada
Private Sub BorrarListaCarpetas(nombreConfig As String)
    On Error Resume Next
    CreateObject("WScript.Shell").RegDelete CFG_PATH_SECTION_RUTAS & nombreConfig & "\"

    If Err.Number = 0 Then
        LogInfo "clsConfiguration", "[BorrarListaCarpetas] - Lista eliminada: " & nombreConfig
    End If

    On Error GoTo 0
End Sub

'=============================================================================
' FUNCIONES DE DIAGNOSTICO
'=============================================================================

'@Description: Devuelve un resumen de la configuracion actual (para diagnostico)
Public Function GetConfigSummary() As String
    Dim summary As String

    summary = "=== CONFIGURACION ACTUAL ===" & vbCrLf
    summary = summary & "Oportunidades: " & mRutaOportunidades & vbCrLf
    summary = summary & "Plantillas: " & mRutaPlantillas & vbCrLf
    summary = summary & "Ofergas: " & mRutaOfergas & vbCrLf
    summary = summary & "GasVBNet: " & mRutaGasVBNet & vbCrLf
    summary = summary & "ExcelCalcs: " & mRutaExcelCalcTempl & vbCrLf
    summary = summary & "SAM: " & mSAM & vbCrLf
    summary = summary & "Imgs Compresores: " & ArrayToString(mArrComprImgs) & vbCrLf
    summary = summary & "Planos PIDs: " & ArrayToString(mArrComprDrawPIDs) & vbCrLf
    summary = summary & "=============================="

    GetConfigSummary = summary
End Function

'@Description: Valida que todas las rutas configuradas existan
Public Function ValidateAllPaths() As Boolean
    Dim allValid As Boolean
    allValid = True

    If Not RutaExiste(mRutaOportunidades) Then
        LogWarning "clsConfiguration", "Ruta Oportunidades no existe: " & mRutaOportunidades
        allValid = False
    End If

    If Not RutaExiste(mRutaPlantillas) Then
        LogWarning "clsConfiguration", "Ruta Plantillas no existe: " & mRutaPlantillas
        allValid = False
    End If

    If Not RutaExiste(mRutaOfergas) Then
        LogWarning "clsConfiguration", "Ruta Ofergas no existe: " & mRutaOfergas
        allValid = False
    End If

    If Not RutaExiste(mRutaGasVBNet) Then
        LogWarning "clsConfiguration", "Ruta GasVBNet no existe: " & mRutaGasVBNet
        allValid = False
    End If

    ValidateAllPaths = allValid
End Function

'@Description: Convierte array a string para diagnostico
Private Function ArrayToString(arr As Variant) As String
    On Error Resume Next

    If Not IsArray(arr) Then
        ArrayToString = "(no es array)"
        Exit Function
    End If

    If UBound(arr) < 0 Then
        ArrayToString = "(vacio)"
        Exit Function
    End If

    ArrayToString = Join(arr, "; ")
End Function
