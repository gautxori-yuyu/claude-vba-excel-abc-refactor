VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsServiceContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("2-Infrastructure.DependencyInjection")
Option Explicit

' ===============================================================================
' CLASE: clsServiceContainer
' ===============================================================================
' PROPOSITO:
'   Contenedor de servicios (DI Container) que gestiona el ciclo de vida
'   de servicios singleton. Permite registrar servicios por nombre y
'   resolverlos bajo demanda con lazy loading opcional.
'
'   Implementa el patrón Service Locator simplificado para VBA.
'
' -------------------------------------------------------------------------------
' RELACIONES CON OTRAS CLASES:
' -------------------------------------------------------------------------------
'   ALMACENA:
'     - Dictionary<String, IService> (servicios registrados)
'
'   USA:
'     - IService (interfaz para todos los servicios)
'
'   USADO POR:
'     - clsApplication (registro y resolución de servicios)
'
' -------------------------------------------------------------------------------
' MECANISMOS DE COMUNICACION:
' -------------------------------------------------------------------------------
'   - PULL: Los clientes llaman GetService("nombre") para obtener servicios
'   - PUSH: clsApplication llama RegisterService() para registrar
'
' -------------------------------------------------------------------------------
' FUNCIONALIDADES MIGRADAS DESDE MAIN:
' -------------------------------------------------------------------------------
' NO EXISTE EN MAIN - NUEVA FUNCIONALIDAD
'
' Basado en propuesta PROP-002 de POOL_PROPUESTAS.md:
'   - Lazy loading de servicios
'   - Registro por nombre
'   - Resolución de dependencias simple
'
' Reemplaza patrón anterior donde clsAplicacion creaba TODO manualmente.
'
' ===============================================================================

' --- IMPLEMENTACION PENDIENTE ---
' TODO Sprint 1: Implementar Dictionary privado para servicios
' TODO Sprint 1: Implementar RegisterService(name, service)
' TODO Sprint 1: Implementar GetService(name) con lazy loading
' TODO Sprint 2: Implementar DisposeAll() para limpieza
