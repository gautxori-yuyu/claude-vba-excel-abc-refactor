VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConfiguracion
   Caption         =   "Configuración"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9180
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("6-Presentation.Forms")
Option Explicit

' ===============================================================================
' FORMULARIO: frmConfiguracion
' ===============================================================================
' PROPOSITO:
'   UI para configurar parámetros de la aplicación:
'     - Rutas de carpetas (oportunidades, plantillas, gas)
'     - Números SAM
'     - Carpetas de imágenes y planos
'
'   Persiste cambios en clsApplicationConfiguration (Registry).
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   USA:
'     - clsApplicationConfiguration (lee/escribe configuración)
'     - App() (acceso global a la aplicación)
'
'   LLAMADO POR:
'     - modRibbonCallbacks.OnConfigurador()
'     - clsApplication (si configuración no válida al inicio)
'
' -------------------------------------------------------------------------------
' CONTROLES PRINCIPALES:
' -------------------------------------------------------------------------------
'   - txtRutaOportunidades (TextBox) - Ruta carpeta oportunidades
'   - txtRutaPlantillas (TextBox) - Ruta carpeta plantillas
'   - txtRutaGasVBNet (TextBox) - Ruta carpeta gas
'   - cmdSeleccionar* (CommandButton) - Selectores de carpeta
'   - cmdAceptar (CommandButton) - Confirmar y guardar
'   - cmdCancelar (CommandButton) - Cancelar cambios
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' ===============================================================================
' Archivo origen: frmConfiguracion.frm (líneas 1-378)
'
' F-001: Inicialización del formulario
'   - Legacy: UserForm_Initialize() (línea ~50)
'   - Migra: Carga valores desde App.Configuration
'
' F-002: Selección de carpetas
'   - Legacy: cmdSeleccionar_Click() (línea ~100)
'   - Migra: Usar BrowseForFolder API
'
' F-003: Validación de rutas
'   - Legacy: ValidarRutas() (línea ~150)
'   - Migra: Validar que carpetas existan
'
' F-004: Guardar configuración
'   - Legacy: cmdAceptar_Click() (línea ~200)
'   - Migra: App.Configuration.SaveToRegistry()
'
' F-005: Cancelar cambios
'   - Legacy: cmdCancelar_Click() (línea ~250)
'   - Migra: Unload Me sin guardar
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 2: Diseñar controles en editor VBA
' TODO Sprint 2: Private Sub UserForm_Initialize()
' TODO Sprint 2: Private Sub cmdSeleccionar_Click()
' TODO Sprint 2: Private Sub cmdAceptar_Click()
' TODO Sprint 2: Private Sub cmdCancelar_Click()
' TODO Sprint 3: Private Function ValidarRutas() As Boolean
