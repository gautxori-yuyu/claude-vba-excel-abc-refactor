VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("1-Application")
Option Explicit

' ===============================================================================
' CLASE: ThisWorkbook - Entry Point de la Aplicación
' ===============================================================================
' PROPOSITO: Entry point con CERO lógica de negocio
' MIGRADO DESDE: main:ThisWorkbook.cls (172 lineas)
' PATRÓN: Singleton Factory
' ===============================================================================

' ==========================================
' SINGLETON GLOBAL
' ==========================================

Private mApp As clsApplication

'@Description: Retorna instancia singleton de la aplicación
Friend Function App() As clsApplication
    If mApp Is Nothing Then
        Set mApp = New clsApplication
    End If
    Set App = mApp
End Function

' ==========================================
' EVENTOS DE CICLO DE VIDA
' ==========================================

Private Sub Workbook_Open()
    ' 1. Limpiar logs antiguos
    RemoveOldLogs

    ' 2. Inicializar logger
    InitLogger LOG_DEBUG, True, ThisWorkbook.Path & "\ABC_Log_" & Format(Now, "yyyy-mm-dd_hh-nn") & ".txt"
    LogInfo "ThisWorkbook", "Iniciando aplicación..."

    ' 3. Inicializar aplicación
    App.Initialize

    LogInfo "ThisWorkbook", "Aplicación iniciada correctamente"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next

    LogInfo "ThisWorkbook", "Cerrando aplicación..."

    ' Terminar aplicación
    If Not mApp Is Nothing Then
        mApp.Terminate
        Set mApp = Nothing
    End If

    LogInfo "ThisWorkbook", "Aplicación cerrada"
    On Error GoTo 0
End Sub

Private Sub Workbook_AddinInstall()
    ' Evento cuando se marca en Opciones > Complementos
    On Error Resume Next
    On Error GoTo 0
End Sub

Private Sub Workbook_AddinUninstall()
    ' Evento cuando se desmarca en Opciones > Complementos
    On Error Resume Next
    On Error GoTo 0
End Sub

' ==========================================
' UTILIDADES PRIVADAS
' ==========================================

Private Sub RemoveOldLogs()
    ' Elimina logs con más de 3 días si hay más de 3 archivos
    On Error Resume Next
    
    Dim fso As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim logFiles As Collection
    Dim filePattern As String
    Dim folderPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = ThisWorkbook.Path & "\"
    filePattern = "ABC_Log_????-??-??_??-??.txt"
    
    If Not fso.FolderExists(folderPath) Then Exit Sub
    Set oFolder = fso.GetFolder(folderPath)
    
    ' Recopilar archivos de log
    Set logFiles = New Collection
    For Each oFile In oFolder.Files
        If oFile.Name Like filePattern Then
            logFiles.Add oFile
        End If
    Next
    
    ' Eliminar logs antiguos (más de 3 días) si hay más de 3 archivos
    For Each oFile In logFiles
        If logFiles.Count > 3 And DateDiff("d", oFile.DateLastModified, Now) > 3 Then
            oFile.Delete True
        End If
    Next
    
    On Error GoTo 0
End Sub
