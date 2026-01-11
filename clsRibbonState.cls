VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRibbonState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Description: Estado lógico del Ribbon (NO es un servicio - es estado puro)
'@Category: State
'@Folder "2.Infrastructure.State"
Option Explicit

'=============================================================================
' clsRibbonState - Estado del Ribbon (gestionado por ApplicationContext)
'=============================================================================
' Responsabilidades:
'   1. Mantener el modo actual del Ribbon (Admin, User, Hidden, etc.)
'   2. Notificar cambios de estado via eventos
'   3. Determinar visibilidad de tabs/grupos según el modo
'=============================================================================
' NOTA: Esta clase NO es un servicio. Es un objeto de estado puro
'       gestionado por clsApplicationContext.
'=============================================================================

Public Event StateChanged()

Private mModoRibbon As eRibbonMode
Private mVisible As Boolean

'=============================================================================
' PROPIEDADES - Modo del Ribbon
'=============================================================================

Public Property Get Modo() As eRibbonMode
    Modo = mModoRibbon
End Property

Public Property Let Modo(Value As eRibbonMode)
    If mModoRibbon = Ribbon_Undefined Then
        mModoRibbon = Value
    ElseIf mModoRibbon <> Value Then
        mModoRibbon = Value
        RaiseEvent StateChanged
    End If
End Property

Public Sub ToggleModo()
    Select Case mModoRibbon
    Case Ribbon_Admin
        mModoRibbon = Ribbon_Hidden
    Case Ribbon_Hidden
        mModoRibbon = Ribbon_User
    Case Ribbon_User
        mModoRibbon = Ribbon_OpportunityOnly
    Case Ribbon_OpportunityOnly
        mModoRibbon = Ribbon_Admin
    Case Else
        mModoRibbon = Ribbon_OpportunityOnly
    End Select
    RaiseEvent StateChanged
End Sub

Public Function RibbonStateDescription() As String
    Select Case mModoRibbon
    Case Ribbon_Admin
        RibbonStateDescription = "Admin"
    Case Ribbon_Hidden
        RibbonStateDescription = "(oculto)"
    Case Ribbon_OpportunityOnly
        RibbonStateDescription = "Fich. oportunidad"
    Case Ribbon_User
        RibbonStateDescription = "User"
    Case Else
        RibbonStateDescription = "Fich. oportunidad"
    End Select
End Function

'=============================================================================
' PROPIEDADES - Visibilidad
'=============================================================================

Public Property Get Visible() As Boolean
    Visible = mVisible
End Property

Public Property Let Visible(Value As Boolean)
    If mVisible <> Value Then
        mVisible = Value
        RaiseEvent StateChanged
    End If
End Property

'=============================================================================
' MÉTODOS - Determinación de visibilidad
'=============================================================================

Public Function IsRibbonTabVisible() As Boolean
    Select Case mModoRibbon
    Case Ribbon_Admin
        IsRibbonTabVisible = True
    Case Ribbon_User
        IsRibbonTabVisible = True
    Case Ribbon_OpportunityOnly
        IsRibbonTabVisible = EsFicheroOportunidad()
    Case Else
        IsRibbonTabVisible = False
    End Select
End Function

Public Function IsAdminGroupVisible() As Boolean
    IsAdminGroupVisible = (mModoRibbon = Ribbon_Admin)
End Function

'=============================================================================
' MÉTODOS - Reset
'=============================================================================

Public Sub Reset()
    mModoRibbon = Ribbon_Undefined
    mVisible = False
End Sub

