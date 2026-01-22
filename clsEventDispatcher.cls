VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventDispatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("2-Infrastructure.Events")
Option Explicit

' ===============================================================================
' CLASE: clsEventDispatcher
' ===============================================================================
' PROPOSITO: Dispatcher generico de eventos (patron Observer manual)
' MIGRADO DESDE: main:clsEventDispatcher.cls (10 lineas)
' ===============================================================================
'
' RESPONSABILIDADES:
' - Gestionar suscripciones a eventos custom
' - Notificar a multiples suscriptores de un evento
' - Alternativa a WithEvents cuando se necesita N suscriptores
'
' RELACIONES:
' - PUEDE SER USADO POR: Cualquier servicio que necesite notificar multiples consumidores
' - EN REFACTOR: Uso minimo (preferimos PULL y callbacks directos)
'
' COMUNICACION:
' - PUSH: Servicios invocan NotifySubscribers(eventName, args)
' - PULL: Consumidores se suscriben via Subscribe(eventName, callback)
'
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' - clsEventDispatcher.cls (lineas 1-10):
'   * Sub Subscribe(eventName As String, handler As Object)
'   * Sub Unsubscribe(eventName As String, handler As Object)
'   * Sub NotifySubscribers(eventName As String, ParamArray args())
'   * Property Get SubscriberCount(eventName As String) As Long
'
' NOTA ARQUITECTONICA:
' En la rama main, esta clase tiene solo 10 lineas y es minimalista.
' En el refactor, su uso sera EXCEPCIONAL:
' - Preferimos PULL (property getters) para consultas de estado
' - Preferimos CALLBACKS directos (App.MetodoPublico) para acciones
' - Solo usamos eventos para:
'   1. Eventos COM (Excel.Application, Chart, FolderWatcher)
'   2. Eventos de dominio 1-a-N (OpportunityCreated → multiples suscriptores)
'
' TODO Sprint 3: Evaluar si es necesario migrar o eliminar
' Si se migra, mantener implementacion simple (Dictionary de Collections)
' ===============================================================================
