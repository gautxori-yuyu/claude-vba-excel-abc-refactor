---
name: vba-development
description: Desarrollo profesional de código VBA para Microsoft Office (Excel, Word, Access, PowerPoint), SolidWorks y otras aplicaciones COM. Usar cuando se solicite crear, revisar, refactorizar o documentar código VBA. Aplica estándares de calidad en arquitectura orientada a objetos, patrones de diseño, gestión de errores, optimización de rendimiento, documentación en castellano y codificación ANSI compatible con el editor VBA de Office 365.
---

# Desarrollo VBA Profesional

Skill para la creación y revisión de código VBA de alta calidad, aplicable tanto en Claude AI como en Claude Code.

## Principios fundamentales

Todo código VBA debe cumplir con criterios de: mantenibilidad, flexibilidad, testeabilidad, reusabilidad, interoperabilidad, fiabilidad, eficiencia, integridad y usabilidad.

## Codificación de ficheros

**CRÍTICO**: Todos los ficheros `.bas`, `.cls`, `.frm` deben guardarse en codificación **ANSI (Windows-1252)** para compatibilidad con el editor VBA de Office 365. Esta codificación soporta perfectamente tildes y caracteres del castellano.

## Contextos de ejecución

El código puede ejecutarse desde múltiples contextos:
- **Ribbon/Formularios**: Macros invocadas desde la interfaz de usuario
- **UDFs**: Funciones de hoja de cálculo (User Defined Functions)
- **Eventos**: Respuesta a acciones del usuario o la aplicación
- **Macros**: Ejecución directa desde el editor o atajos

Diseñar procedimientos flexibles que funcionen en varios contextos cuando sea posible.

## Arquitectura y organización

### Estructura de módulos

```
Proyecto VBA/
├── Módulos estándar (.bas)
│   ├── modConstantes     ' Constantes globales y declaraciones Win32 API
│   ├── modUtilidades     ' Funciones auxiliares de uso general
│   ├── modMain           ' Punto de entrada y orquestación
│   └── modTestXxx        ' Módulos de pruebas unitarias (sin guiones bajos)
├── Módulos de clase (.cls)
│   ├── cls*              ' Clases de dominio
│   └── cls*Factory       ' Factorías para instanciación
└── Formularios (.frm)
    └── frm*              ' UserForms con lógica separada
```

### Patrones de diseño recomendados

**Estructurales:**
- **Repository**: Acceso a datos encapsulado
- **Facade**: Interfaz simplificada para subsistemas complejos

**Creacionales:**
- **Factory**: Instanciación de clases (VBA carece de constructores con parámetros)
- **Singleton**: Gestores globales (Logger, ConfigManager)

**Comportamiento:**
- **Observer/Events**: Uso de `WithEvents` para comunicación desacoplada
- **Strategy**: Algoritmos intercambiables mediante interfaces

### Separación en UserForms

Aplicar patrón MVP adaptado:
- **Vista (UserForm)**: Solo presentación y captura de eventos
- **Presentador (Clase)**: Lógica de presentación y validación
- **Modelo (Clase)**: Datos y reglas de negocio

Inicialización mediante método público `Init()` en lugar de `UserForm_Initialize` para permitir inyección de dependencias.

## Declaraciones y tipos

### Option Explicit obligatorio

Todos los módulos deben incluir `Option Explicit` en la primera línea.

### Convención de nombres (notación húngara adaptada)

Formato: `[ámbito][tipo]NombreDescriptivo`

**Ámbito:**
- `l` - Local (procedimiento)
- `m` - Módulo/clase (Private)
- `g` - Global (Public)

**Tipos comunes:**
- `str` - String
- `lng` - Long
- `dbl` - Double
- `bln` - Boolean
- `obj` - Object
- `col` - Collection
- `dic` - Dictionary
- `rng` - Range
- `ws` - Worksheet
- `wb` - Workbook
- `arr` - Array
- `var` - Variant (evitar si es posible especificar tipo)

**Calificadores recomendados:** Cont, Min, Max, Primer, Ult, Sig, Act, Prev, Tmp

**Orden de especificidad ascendente:** `FilasUsadasUlt` mejor que `UltFilasUsadas`

### Late Binding vs Early Binding

**Late Binding** (preferido para objetos externos):
```vba
Dim objFSO As Object ' Scripting.FileSystemObject
Dim objRegEx As Object ' VBScript.RegExp
Dim objXML As Object ' MSXML2.DOMDocument60
```

**Early Binding** (para objetos de la aplicación host):
```vba
Dim wsHoja As Worksheet
Dim rngDatos As Range
```

### Estructuras de datos

Seleccionar según necesidad:
- **Array**: Acceso por índice, tamaño fijo o conocido, máximo rendimiento
- **Collection**: Acceso por clave o índice, tamaño dinámico
- **Scripting.Dictionary**: Lookups O(1), verificación de existencia, claves únicas

Preferir `Enum` sobre strings para valores discretos conocidos.

## Gestión de errores

### Estructura estándar

```vba
Public Sub EjemploProcedimiento()
    On Error GoTo ErrorHandler
    
    ' Código principal
    
CleanExit:
    ' Limpieza de recursos
    Exit Sub
    
ErrorHandler:
    ' Gestión del error
    Err.Raise Err.Number, "NombreModulo.EjemploProcedimiento", Err.Description
End Sub
```

### Propagación de errores (gestión apilada)

Usar `Err.Raise` para propagar errores a procedimientos de nivel superior, similar al manejo de excepciones en Java:

```vba
' Procedimiento de bajo nivel
Private Sub ProcesoInterno()
    On Error GoTo ErrorHandler
    ' ...
ErrorHandler:
    Err.Raise Err.Number, "Modulo.ProcesoInterno", Err.Description
End Sub

' Procedimiento de alto nivel
Public Sub ProcesoMain()
    On Error GoTo ErrorHandler
    Call ProcesoInterno
    Exit Sub
ErrorHandler:
    MsgBox "Error en " & Err.Source & ": " & Err.Description
End Sub
```

### Aserciones y depuración

Usar compilación condicional con argumento `Debugging=1`:

```vba
#If Debugging Then
    Debug.Assert lngContador > 0
    Debug.Print "Valor actual: " & lngContador
#End If
```

### Errores COM frecuentes

Gestionar específicamente:
- **429**: No se puede crear el objeto ActiveX
- **462**: El servidor remoto no existe o no está disponible
- **-2147221005**: Cadena de formato no válida

## Optimización de rendimiento

### Principio de preservación del estado

Salvo que el objetivo explícito de un programa sea modificar el estado de la aplicación, un fichero u otro contexto, todo debe volver al mismo estado en que se encontraba antes de iniciar la ejecución.

### Gestión del entorno de ejecución

```vba
Private Type TEstadoAplicacion
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    Calculation As XlCalculation
    StatusBar As Variant
End Type

Private mEstadoOriginal As TEstadoAplicacion

Public Sub GuardarYOptimizarEntorno()
    With Application
        mEstadoOriginal.ScreenUpdating = .ScreenUpdating
        mEstadoOriginal.EnableEvents = .EnableEvents
        mEstadoOriginal.Calculation = .Calculation
        mEstadoOriginal.StatusBar = .StatusBar
        
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .StatusBar = "Procesando..."
    End With
End Sub

Public Sub RestaurarEntorno()
    With Application
        .ScreenUpdating = mEstadoOriginal.ScreenUpdating
        .EnableEvents = mEstadoOriginal.EnableEvents
        .Calculation = mEstadoOriginal.Calculation
        .StatusBar = mEstadoOriginal.StatusBar
    End With
End Sub
```

### Prácticas de rendimiento

- **Arrays en memoria**: Leer rangos a arrays, procesar, escribir de vuelta
- **Evitar Select/Activate**: Usar referencias directas (`ws.Range("A1")` no `ws.Select: Range("A1").Select`)
- **Value2 vs Value**: Usar `Value2` para datos numéricos (evita conversión de fechas/moneda)
- **Minimizar accesos al objeto Range**: Cada acceso cruza la frontera COM
- **Usar métodos nativos**: `Find`, `Match`, `WorksheetFunction` en lugar de bucles

### Gestión de memoria

- Establecer objetos a `Nothing` al terminar de usarlos
- Liberar objetos COM en **orden inverso** a su creación
- Evitar referencias circulares entre clases

```vba
CleanExit:
    Set objHijo = Nothing
    Set objPadre = Nothing  ' Liberar después del hijo
    Exit Sub
```

## Estilo de código

### Indentación

Usar indentación consistente de 4 espacios por nivel.

### Longitud de línea

Máximo 100 caracteres. Usar continuación de línea con ` _`:

```vba
strSQL = "SELECT Campo1, Campo2, Campo3 " & _
         "FROM Tabla " & _
         "WHERE Condicion = True"
```

### Evitar

- Sentencia `GoTo` excepto para gestión de errores
- Anidaciones excesivas (máximo 3 niveles): extraer a subrutinas o usar guardas con `Exit Sub/Function`
- Variables globales cuando se puede pasar por parámetro
- `On Error Resume Next` como gestión general (solo para comprobaciones puntuales)

### Referencias cualificadas

Usar referencias totalmente cualificadas para evitar ambigüedades:

```vba
' Correcto
ThisWorkbook.Worksheets("Datos").Range("A1").Value

' Evitar
Range("A1").Value  ' ¿De qué hoja? ¿De qué libro?
```

### Uso correcto de With

```vba
' Correcto - evita requalificación
With wsHoja.Range("A1:Z100")
    .Font.Bold = True
    .Interior.Color = vbYellow
End With

' Incorrecto - requalifica en cada línea
wsHoja.Range("A1:Z100").Font.Bold = True
wsHoja.Range("A1:Z100").Interior.Color = vbYellow
```

## Documentación

### Cabecera de módulo

```vba
'==============================================================
' Módulo: modNombreModulo
'--------------------------------------------------------------
' Descripción: Breve descripción del propósito del módulo
' Autor: [Nombre]
' Versión: 1.0
' Última modificación: [Fecha]
' Requiere: [Referencias COM necesarias]
'==============================================================
Option Explicit
```

### Cabecera de sección

```vba
'--------------------------------------------------------------
' Funciones de validación de datos
'--------------------------------------------------------------
```

### Documentación de procedimientos

```vba
'@Description: Calcula el importe total de una oferta aplicando descuentos.
'@Scope: Libro activo, hoja "Ofertas"
'@ArgumentDescriptions: lngIdOferta: Identificador único de la oferta | dblDescuento: Porcentaje de descuento (0-100)
'@Returns: Double | Importe total con descuento aplicado, -1 si error
'@Category: Cálculos
'@Example: dblTotal = CalcularImporteOferta(1001, 15) ' Aplica 15% descuento
Public Function CalcularImporteOferta(ByVal lngIdOferta As Long, _
                                       Optional ByVal dblDescuento As Double = 0) As Double
```

Para funciones públicas complejas, incluir siempre `@Example` con un caso de uso representativo.

### Comentarios en código

- Comentar el **por qué**, no el **qué** (el código ya dice qué hace)
- Documentar algoritmos complejos por fases/etapas
- Todo en castellano correcto, con tildes

## Interoperabilidad

### Win32 API

Las declaraciones Win32 API se incluyen en `modConstantes` junto con las demás constantes globales, dado que son invariables:

```vba
'--------------------------------------------------------------
' Declaraciones Win32 API
'--------------------------------------------------------------
#If VBA7 Then
    Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Public Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
        (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
#Else
    Public Declare Function GetTickCount Lib "kernel32" () As Long
    Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
        (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
#End If
```

### Componentes COM externos

Para funcionalidad que requiera componentes COM en VB.NET:
- Registrar en HKCU (HKEY_CURRENT_USER) del registro de Windows, o mediante Activation Context
- **NO usar** `regasm /codebase` (requiere permisos de administrador y afecta a todos los usuarios)
- Usar Late Binding desde VBA para evitar dependencias de versión
- Documentar requisitos de instalación en cabecera del módulo

### Seguridad con ADO/DAO

Usar consultas parametrizadas siempre que el objeto lo soporte, para prevenir inyección SQL:

```vba
' Correcto - parametrizado con ADODB.Command
With cmd
    .CommandText = "SELECT * FROM Clientes WHERE CodCliente = ?"
    .Parameters.Append .CreateParameter(, adInteger, adParamInput, , lngCodCliente)
End With

' Correcto - parametrizado con DAO.QueryDef
Set qdf = db.CreateQueryDef("", "SELECT * FROM Clientes WHERE CodCliente = [pCodCliente]")
qdf.Parameters("pCodCliente") = lngCodCliente

' Incorrecto - vulnerable a inyección
strSQL = "SELECT * FROM Clientes WHERE CodCliente = " & lngCodCliente
```

## Testing

### Estructura de módulos de test

Crear módulos `modTestNombreModulo` para cada módulo a probar (sin guiones bajos en el nombre).

### Convención de nombres

**IMPORTANTE**: No usar guiones bajos en nombres de procedimientos. El carácter `_` está reservado por VBA para identificar manejadores de eventos (ej: `CommandButton1_Click`). Usar PascalCase:

```vba
Public Sub TestCalcularImporteSinDescuentoDevuelveImporteBase()
Public Sub TestCalcularImporteDescuento50DevuelveMitad()
Public Sub TestCalcularImporteIdInvalidoDevuelveMenosUno()
```

### Gestión con Rubberduck

Los tests se gestionan mediante Rubberduck para:
- Ejecución automatizada
- Reporte de resultados
- Integración con el IDE de VBA

## Control de versiones

### Codificación de ficheros exportados

Los ficheros exportados para Git deben mantener codificación **ANSI (Windows-1252)** para preservar tildes y caracteres del castellano. Verificar que el cliente Git no altere la codificación.

### Exportación para Git

Exportar módulos como ficheros independientes:
- `.bas` para módulos estándar
- `.cls` para módulos de clase
- `.frm` + `.frx` para formularios

### Estructura de repositorio

**Estructura plana (por defecto)**: Una única carpeta raíz que contiene todos los ficheros fuente. Facilita la comparación carpeta contra carpeta con herramientas externas (Beyond Compare).

```
proyecto-vba/          ' Rama main - código del programador
├── modConstantes.bas
├── modUtilidades.bas
├── modMain.bas
├── clsOferta.cls
├── frmPrincipal.frm
├── frmPrincipal.frx
└── README.md
```

**Flujo de trabajo con ramas**: La rama `main` contiene el código del programador. Las ramas adicionales (típicamente una por asistente IA) contienen las modificaciones propuestas, permitiendo comparar directamente con la rama principal.

**Estructura ramificada (proyectos complejos)**: Solo para proyectos de mayor envergadura que requieran organización por carpetas.

## Lista de verificación para revisión de código

Al revisar código VBA, verificar:

- [ ] Codificación ANSI (Windows-1252) del fichero
- [ ] `Option Explicit` presente
- [ ] Convención de nombres respetada (sin guiones bajos en procedimientos)
- [ ] Gestión de errores implementada con preservación/restauración de estado
- [ ] Objetos liberados correctamente (orden inverso)
- [ ] Sin Select/Activate innecesarios
- [ ] Documentación de módulos con Versión, Última modificación, Requiere
- [ ] Documentación de procedimientos con @Example en funciones públicas complejas
- [ ] Comentarios en castellano correcto, con tildes
- [ ] Indentación consistente (4 espacios)
- [ ] Sin anidaciones excesivas (máximo 3 niveles)
- [ ] Referencias cualificadas donde corresponda
- [ ] Consultas SQL parametrizadas donde el objeto lo soporte
- [ ] Declaraciones Win32 API en modConstantes
- [ ] Componentes COM registrados en HKCU o con Activation Context
