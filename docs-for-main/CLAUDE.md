# CLAUDE.md - Rama Principal (main)

## Descripción del Proyecto

Add-in de Excel (.xlam) para gestión de oportunidades comerciales y herramientas de productividad, con funcionalidades de:
- Generación y formateo de gráficos de sensibilidad
- Monitorización de carpetas del sistema de archivos
- Gestión de archivos Excel abiertos
- Interfaz Ribbon personalizada
- UDFs (User Defined Functions) para cálculos especializados
- Integración con CoolProp para propiedades termodinámicas
- Gestión de ofertas y presupuestos

## Estado Actual de la Rama

Esta es la **rama principal** del proyecto, conteniendo el código base original. Existe una rama de refactorización (`claude/refactor-main-code-YeuB5`) que implementa una arquitectura mejorada con patrones de diseño profesionales.

### Estructura de Ficheros

```
Proyecto/
├── Módulos estándar (.bas)
│   ├── mod_ConstantsGlobals.bas   ' Constantes y tipos globales
│   ├── mod_Logger.bas             ' Sistema de logging
│   ├── modCALLBACKSRibbon.bas     ' Callbacks del Ribbon XML
│   ├── modAPP*.bas                ' Módulos de aplicación
│   ├── modMACRO*.bas              ' Macros ejecutables
│   ├── modUTILS*.bas              ' Utilidades
│   └── UDFs_*.bas                 ' User Defined Functions
│
├── Clases de dominio (.cls)
│   ├── clsAplicacion.cls          ' Coordinador principal
│   ├── clsExecutionContext.cls    ' Contexto de ejecución (eventos Application)
│   ├── clsFileManager.cls         ' Gestión de archivos Excel
│   ├── clsOpportunitiesMgr.cls    ' Gestión de oportunidades
│   ├── clsChartEventsManager.cls  ' Eventos de gráficos
│   ├── clsFSMonitoringCoord.cls   ' Monitorización filesystem
│   ├── clsConfiguration.cls       ' Configuración desde registro
│   ├── clsRibbonState.cls         ' Estado del Ribbon
│   ├── clsRibbonEvents.cls        ' Eventos del Ribbon
│   └── cls*.cls                   ' Otras clases de dominio
│
└── Interfaces (.cls)
    └── IFile.cls                  ' Interfaz para archivos
```

### Componentes Principales

| Componente | Fichero(s) | Responsabilidad |
|------------|-----------|-----------------|
| **Coordinador** | `clsAplicacion.cls` | Orquestación principal, facade de servicios |
| **Contexto** | `clsExecutionContext.cls` | Eventos de Application, acceso seguro a ActiveWorkbook/Sheet/Chart |
| **Archivos** | `clsFileManager.cls`, `clsExcelFile.cls` | Tracking de workbooks abiertos |
| **Oportunidades** | `clsOpportunitiesMgr.cls`, `clsOpportunity.cls` | Gestión de carpetas de oportunidades |
| **Gráficos** | `clsChartEventsManager.cls`, `clsChartEvents.cls` | Eventos de activación/desactivación de gráficos |
| **Filesystem** | `clsFSMonitoringCoord.cls`, `clsFSWatcher.cls` | Monitorización de cambios en carpetas |
| **Ribbon** | `clsRibbonState.cls`, `clsRibbonEvents.cls` | Estado y callbacks del Ribbon |
| **UDFs** | `UDFs_*.bas` | Funciones de hoja de cálculo |

### Funcionalidades UDF

| Módulo | Funcionalidades |
|--------|-----------------|
| `UDFs_CGASING.bas` | Funciones para cálculos CGASING |
| `UDFs_COOLPROP.bas` | Propiedades termodinámicas via CoolProp |
| `UDFs_Units.bas` | Conversión de unidades |
| `UDFs_FileSystem.bas` | Operaciones de sistema de archivos |
| `UDFs_UtilsExcel.bas` | Utilidades de Excel |
| `UDFs_UtilsExcelChart.bas` | Utilidades de gráficos |
| `UDFs_Backups.bas` | Gestión de backups |

---

## Objetivo de Refactorización

La rama `claude/refactor-main-code-YeuB5` implementa las siguientes mejoras arquitectónicas:

1. **Inyección de Dependencias (DI)**
   - `clsServiceManager` como contenedor de servicios
   - Resolución por tipo (sin strings hardcodeados)
   - Lazy initialization

2. **Patrón Mediator**
   - `clsEventCoordinator` centraliza todos los eventos
   - Reduce acoplamiento entre componentes

3. **Separación Estado/Servicios**
   - `clsApplicationContext` agrega estado (NO es un servicio)
   - Servicios implementan `IService` con ciclo de vida

4. **Interfaz IService**
   - `Initialize()`, `Dispose()`, `IsInitialized`, `ServiceName`
   - Gestión uniforme del ciclo de vida

---

## Estándares de Desarrollo VBA

### Codificación de Ficheros

**CRÍTICO**: Todos los ficheros `.bas`, `.cls`, `.frm` deben guardarse en codificación **ANSI (Windows-1252)** para compatibilidad con el editor VBA de Office 365.

### Convenciones de Nombres

Formato: `[ámbito][tipo]NombreDescriptivo`

- **Ámbito**: `l` (local), `m` (módulo), `g` (global)
- **Tipos comunes**: `str`, `lng`, `dbl`, `bln`, `obj`, `col`, `dic`, `rng`, `ws`, `wb`, `arr`

**IMPORTANTE**: No usar guiones bajos en nombres de procedimientos (reservado para eventos VBA).

### Organización de Módulos

- `mod_*` - Módulos de infraestructura (constantes, logger)
- `modAPP*` - Módulos de aplicación
- `modMACRO*` - Macros ejecutables desde Ribbon/menús
- `modUTILS*` - Utilidades reutilizables
- `UDFs_*` - User Defined Functions
- `cls*` - Clases de dominio

### Gestión de Errores

```vba
Public Sub EjemploProcedimiento()
    On Error GoTo ErrorHandler

    ' Código principal

CleanExit:
    ' Limpieza de recursos
    Exit Sub

ErrorHandler:
    LogError "NombreModulo", "[Procedimiento]", Err.Number, Err.Description
    Resume CleanExit
End Sub
```

### Documentación

```vba
'@Description: Breve descripción
'@Scope: Contexto de uso
'@ArgumentDescriptions: param1: desc | param2: desc
'@Returns: Tipo | Descripción
'@Example: Ejemplo de uso
```

### Optimización

- Arrays en memoria para rangos grandes
- Evitar `Select`/`Activate`
- Usar `Value2` para datos numéricos
- Preservar/restaurar estado de Application
- Liberar objetos en orden inverso

---

## Flujo de Trabajo Git

### Estructura de Ramas

- `main` - Código base del programador
- `claude/*` - Ramas de trabajo de asistentes IA

### Comparación de Ramas

Para revisar cambios propuestos en refactorización:

```bash
git diff main..claude/refactor-main-code-YeuB5
```

---

## Referencias

Ver ficheros de especificaciones VBA en rama refactorizada:
- `vba-development/SKILL.md`
- `vba-development/references/patrones-diseno.md`
- `vba-development/references/optimizacion.md`
