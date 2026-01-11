# CLAUDE.md - Rama Refactorizada (claude/refactor-main-code-YeuB5)

## Descripción del Proyecto

Add-in de Excel (.xlam) para gestión de oportunidades comerciales, con funcionalidades de:
- Generación y formateo de gráficos de sensibilidad
- Monitorización de carpetas del sistema de archivos
- Gestión de archivos Excel abiertos
- Interfaz Ribbon personalizada

## Estado Actual de la Rama

Esta rama contiene la **arquitectura refactorizada** del proyecto, implementando patrones de diseño profesionales adaptados a VBA.

### Arquitectura Implementada

```
clsAplicacion (Composition Root)
    │
    ├── clsApplicationContext (Estado compartido - NO es servicio)
    │       ├── RibbonState (clsRibbonState)
    │       ├── ChartState (clsChartState)
    │       └── ExecutionContext (clsExecutionContext)
    │
    ├── clsServiceManager (Contenedor DI - resolución por tipo)
    │       ├── .Configuration → clsConfiguration
    │       ├── .ExecutionContext → clsExecutionContext
    │       ├── .FileManager → clsFileManager
    │       ├── .OpportunitiesMgr → clsOpportunitiesMgr
    │       ├── .ChartEventsManager → clsChartEventsManager
    │       ├── .FSMonitoringCoord → clsFSMonitoringCoord
    │       └── .RibbonUI → clsRibbonUI
    │
    └── clsEventCoordinator (Mediator - eventos centralizados)
            └── WithEvents para todos los servicios
```

### Patrones de Diseño Aplicados

| Patrón | Implementación | Clase(s) |
|--------|----------------|----------|
| **Dependency Injection** | Contenedor con resolución tipada | `clsServiceManager` |
| **Mediator** | Coordinación centralizada de eventos | `clsEventCoordinator` |
| **Composition Root** | Punto único de creación | `clsAplicacion` |
| **Observer** | `WithEvents` para comunicación desacoplada | Múltiples clases |
| **Singleton** | Servicios únicos gestionados por DI | Via `RegisterSingleton` |
| **Facade** | Acceso simplificado a servicios | `clsAplicacion` properties |

### Decisiones Arquitectónicas

1. **clsApplicationContext NO es un servicio**
   - Es un agregador de estado puro
   - Se crea ANTES que los servicios (orden de dependencias)
   - Los servicios acceden a él via `ServiceManager.AppContext`

2. **clsRibbonState NO es un servicio**
   - Es estado puro, gestionado por ApplicationContext
   - Emite eventos `StateChanged` que maneja el EventCoordinator

3. **Resolución de servicios sin strings**
   - `TypeName()` como clave interna del diccionario
   - Propiedades tipadas: `mServiceManager.RibbonUI` en lugar de `GetService("IRibbonUI")`
   - Método `GetService()` marcado como DEPRECATED

4. **Eventos centralizados (Mediator)**
   - TODOS los eventos fluyen a través de `clsEventCoordinator`
   - Facilita trazabilidad, testing y mantenimiento
   - Reduce acoplamiento entre servicios

### Interfaz IService

Todos los servicios implementan `IService`:

```vba
' IService.cls
Public Sub Initialize(ByVal dependencies As Object)
Public Sub Dispose()
Public Property Get IsInitialized() As Boolean
Public Property Get ServiceName() As String  ' Para logging/debug
```

### Flujo de Inicialización

1. `ThisWorkbook_Open` → `App.Initialize`
2. Crear objetos de estado (RibbonState, ChartState)
3. Crear ApplicationContext y componerlo
4. Crear ServiceManager e inyectar AppContext
5. Registrar servicios (lazy initialization)
6. Crear EventCoordinator y suscribir a eventos
7. Ribbon callback `RibbonOnLoad` → `RibbonUI.Init`

### Ficheros Principales

| Fichero | Responsabilidad |
|---------|-----------------|
| `clsAplicacion.cls` | Composition Root, orquestación |
| `clsServiceManager.cls` | Contenedor DI tipado |
| `clsEventCoordinator.cls` | Mediator de eventos |
| `clsApplicationContext.cls` | Agregador de estado |
| `clsRibbonUI.cls` | Gestión puntero IRibbonUI |
| `clsRibbonState.cls` | Estado lógico del Ribbon |
| `clsRibbonEvents.cls` | Callbacks XML del Ribbon |
| `clsExecutionContext.cls` | Eventos de Application |
| `clsFileManager.cls` | Tracking de archivos Excel |
| `clsOpportunitiesMgr.cls` | Gestión de oportunidades |
| `clsChartEventsManager.cls` | Eventos de gráficos |
| `clsFSMonitoringCoord.cls` | Monitorización filesystem |

---

## Estándares de Desarrollo VBA

### Codificación de Ficheros

**CRÍTICO**: Todos los ficheros `.bas`, `.cls`, `.frm` deben guardarse en codificación **ANSI (Windows-1252)** para compatibilidad con el editor VBA de Office 365.

### Convenciones de Nombres

Formato: `[ámbito][tipo]NombreDescriptivo`

- **Ámbito**: `l` (local), `m` (módulo), `g` (global)
- **Tipos**: `str`, `lng`, `dbl`, `bln`, `obj`, `col`, `dic`, `rng`, `ws`, `wb`, `arr`

**IMPORTANTE**: No usar guiones bajos en nombres de procedimientos (reservado para eventos).

### Gestión de Errores

```vba
Public Sub EjemploProcedimiento()
    On Error GoTo ErrorHandler

    ' Código principal

CleanExit:
    ' Limpieza de recursos
    Exit Sub

ErrorHandler:
    LogError MODULE_NAME, "[EjemploProcedimiento]", Err.Number, Err.Description
    Resume CleanExit
End Sub
```

### Documentación de Procedimientos

```vba
'@Description: Breve descripción de lo que hace
'@Scope: Contexto de uso
'@ArgumentDescriptions: param1: descripción | param2: descripción
'@Returns: Tipo | Descripción
'@Category: Categoría funcional
'@Example: Ejemplo de uso representativo
```

### Optimización

- Usar arrays en memoria para procesar rangos grandes
- Evitar `Select`/`Activate` - usar referencias directas
- Usar `Value2` en lugar de `Value` para datos numéricos
- Liberar objetos en orden inverso a su creación
- Preservar/restaurar estado de Application (ScreenUpdating, etc.)

---

## Tareas Pendientes

- [ ] Migrar lógica de negocio de gráficos a servicio dedicado
- [ ] Implementar IFormatter para formateo de CGASING
- [ ] Completar handlers de eventos de FSMonitoringCoord
- [ ] Tests unitarios con Rubberduck
- [ ] Documentación de cabecera en todos los módulos

---

## Referencias

- `vba-development/SKILL.md` - Especificaciones generales VBA
- `vba-development/references/patrones-diseno.md` - Patrones de diseño
- `vba-development/references/optimizacion.md` - Técnicas de optimización

