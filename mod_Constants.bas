Attribute VB_Name = "mod_Constants"
'@Folder("2-Infrastructure.Configuration")
Option Explicit

' ===============================================================================
' MODULO: mod_Constants (migrado desde main:mod_ConstantsGlobals.bas)
' ===============================================================================

' Constantes de compilación
#Const RubberduckTest = True
#Const DebugMode = True

' Constantes para organizar la configuración
Public Const APP_NAME As String = "ABC_ofertas maquina especial"
Public Const FOLDERWATCHERCOM_NAME As String = "FolderWatcherCOM.dll"

' Nombres de las configuraciones
Public Const CFG_BASEFOLDER As String = "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\"

Public Const CFG_SAM As Integer = 41
Public Const CFG_PATH_SAM As String = CFG_BASEFOLDER & APP_NAME & "\SAM"

' Configuracion de rutas de carpetas
Public Const CFG_SECTION_RUTAS As String = "Folders"
Public Const CFG_PATH_SECTION_RUTAS As String = CFG_BASEFOLDER & APP_NAME & "\" & CFG_SECTION_RUTAS & "\"

Public Const CFG_RUTA_OPORTUNIDADES As String = "BaseFolderOportunidades"
Public Const CFG_RUTA_OPORTUNIDADES_DEFAULT As String = "C:\abc compressors\INTRANET\OilGas\3_OFERTAS\OFERTAS\2025\41 - SER\"
Public Const CFG_RUTA_PLANTILLAS As String = "BaseFolderPlantillas"
Public Const CFG_RUTA_PLANTILLAS_DEFAULT As String = "C:\abc compressors\INTRANET\OilGas\3_OFERTAS\OFERTAS\2025\41 - SER\_plantilla\"
Public Const CFG_RUTA_OFERGAS As String = "BaseFolderOfergas"
Public Const CFG_RUTA_OFERGAS_DEFAULT As String = "C:\Program Files (x86)\Ofertas_Gas\Excel\"
Public Const CFG_RUTA_GAS_VBNET As String = "BaseFolderGasVBNet"
Public Const CFG_RUTA_GAS_VBNET_DEFAULT As String = "C:\Aire\"
Public Const CFG_RUTA_PLCALCS As String = "BaseFolderXLSCalculos"
Public Const CFG_RUTA_PLCALCS_DEFAULT As String = "C:\abc compressors\2-0-OPORTUNIDADES\_Mis notas\"
Public Const CFG_RUTA_PLCALCNACE As String = "C:\abc compressors\INTRANET\OficinaTecnica\Documentacion\Normas\NACE\Herramienta\Herramienta_para_seleccion_de_materiales_v1.4.xlsx"
Public Const CFG_RUTA_PLSELCILS As String = "C:\abc compressors\INTRANET\OilGas\3_OFERTAS\ADJUNTOS OFERTAS\Datos cilindros 2.xlsx"
Public Const CFG_RUTA_COMPRIMGS As String = "BaseFolderImgsCompresores"
Public Const CFG_RUTA_COMPRIMGS_DEFAULT As String = "C:\abc compressors\INTRANET\OilGas\1_COMUNICACION\0- MARKETING\2- FOTOS\Fotos y planos\FOTOS\"
Public Const CFG_RUTA_COMPRDRAWPID As String = "BaseFolderPlanosPIDs"
Public Const CFG_RUTA_COMPRDRAWPID_DEFAULT As String = "C:\abc compressors\INTRANET\OilGas\5_DOCUMENTACION TECNICA\ADJUNTOS OFERTAS\3-1-PLANOS\|C:\abc compressors\INTRANET\OilGas\1_COMUNICACION\0- MARKETING\2- FOTOS\Fotos y planos\PLANOS"

' Polling de cambios en carpetas
Public Const CFG_FW_HEARTBEAT As String = "Heartbeat"
Public Const POLLING_SECONDS As Integer = 8
Public Const INACTIVITY_MINUTES As Integer = 10
Public Const WARMUP_MAX_CHECKS As Long = 5

' Configuracion de registro de UDFs
Public Const CFG_RUTA_UDFS As String = CFG_BASEFOLDER & APP_NAME & "\UDFsRegistradas"
Public Const DEFAULT_CATEGORY As String = "Funciones Personalizadas"
Public Const DEFAULT_NOPARAMS As String = "(sin parámetros)"
Public Const DEFAULT_NORETURNS As String = "(ninguno)"
Public Const DEFAULT_NOARGS As String = "(sin argumentos)"

' Patrones para parsing de nombres de archivo
Public Const QUOTENR_PATTERN As String = "\d{9}(?:[\-_]\d+)?"
Public Const QUOTENR_REV_PATTERN As String = "(" & QUOTENR_PATTERN & ")(?:[ \-_]*rev\.?[ \-_]*(\d+)\b)?"
Public Const CUSTOMER_PATTERN As String = "(?:.(?! \- ))+."
Public Const PROJECT_OTHERS_PATTERN As String = "(?:.(?! \- ))+."
Public Const MODEL_PATTERN As String = "(\d)\s?T?\s*E\s?(H[AGPX])\s?\-\s?(\d)\s?\-\s?[LGT]{2,3}"
Public Const FULLMODEL_PATTERN As String = MODEL_PATTERN & "(?:\-\d\x\d+T?)+(?: (?:NACE|ATEX))*"
Public Const OPPORTUNITY_MODEL_PATTERN As String = "((?:(?:" & MODEL_PATTERN & ")[ ,y]*)+|X{3,})"
Public Const FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_PATTERN As String = "^(" & QUOTENR_PATTERN & ")\s*\-(?:[#-]-)?\s*(" & CUSTOMER_PATTERN & ")(?:\s*\-\s*(" & PROJECT_OTHERS_PATTERN & "))??"
Public Const FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_MODEL_PATTERN As String = FILEORFOLDERNAME_QUOTE_CUSTOMER_OTHER_PATTERN & "\s*\-\s*" & OPPORTUNITY_MODEL_PATTERN

' Enumeraciones
Public Enum ProcType
    Macro
    udf
    internalPrivate
    internalSubPublicWithParams
    eventHandler
End Enum

Public Enum ProcKind
    proc
    PropLet
    PropSet
    PropGet
    ProcSub
    ProcFunction
End Enum

Public Enum ProcContainerType
    StdModule = 1
    ClassModule = 2
    Form = 3
    ActiveXDesigner = 11
    Sheet = 100
End Enum

Public Enum eRibbonMode
    Ribbon_Undefined = 0
    Ribbon_Hidden = 1
    Ribbon_User = 2
    Ribbon_OpportunityOnly = 3
    Ribbon_Admin = 4
End Enum

Public Enum TipoArchivo
    UnDef = 0
    Unknown = 1
    oportunidad = 2
    CGASING_CurvasRendimiento = 3
    CGASING_Calcs = 4
    PlantillaBudget = 5
    PlantillaOferta = 6
End Enum

' Tipos personalizados
Public Type T_InfoArchivo
    EsValido As Boolean
    TipoDetectado As TipoArchivo
    Customer As String
    OpportunityNr As String
End Type

Public Type T_CodeBlock
    strCode As String
    procStartLine As Long
    procSignatureLine As Long
    procNumLines As Long
End Type

' Declaraciones API
#If Win64 Then
    Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Public Declare Function GetTickCount Lib "kernel32" () As Long
#End If

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
