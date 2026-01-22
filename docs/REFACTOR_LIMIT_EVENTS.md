# REFACTORIZACION: LIMITAR USO DE EVENTOS

> **Objetivo:** Reducir complejidad eliminando eventos innecesarios y reemplazándolos con patrones más explícitos

---

## SITUACION ACTUAL

### Estadísticas de Eventos

| Componente | Eventos Emitidos | Eventos Escuchados | WithEvents |
|------------|------------------|-------------------|------------|
| clsAplicacion | 0 | 26 | 6 clases |
| clsExecutionContext | 6 | 6 (de Application) | 1 (Application) |
| clsFSMonitoringCoord | 12 | 7 (de clsFSWatcher) | 1 |
| clsRibbonEvents | 6 | 0 | 0 |
| clsOpportunitiesMgr | 2 | 0 | 0 |
| clsChartEventsManager | 2 | N (de charts) | N |
| clsRibbonState | 1 | 0 | 0 |

**Total: 30+ eventos en el sistema**

### Problemas Identificados

#### 1. Eventos que Solo Logean (ELIMINAR)

```vba
' En clsAplicacion - ESTOS SOLO LOGEAN, NO HACEN NADA MAS:

Private Sub mFSMonitoringCoord_TemplateCreated(...)
    LogDebug MODULE_NAME, "[mFSMonitoringCoord_TemplateCreated] " & fileName
    ' TODO: Implementar logica... <- NUNCA SE IMPLEMENTO
End Sub

Private Sub mFSMonitoringCoord_TemplateChanged(...)
    LogDebug MODULE_NAME, "[mFSMonitoringCoord_TemplateChanged] " & fileName
    ' TODO: Notificar cambio... <- NUNCA SE IMPLEMENTO
End Sub

Private Sub mFSMonitoringCoord_GasFileCreated(...)
    LogDebug MODULE_NAME, "[mFSMonitoringCoord_GasFileCreated] " & fileName
    ' TODO: Implementar logica... <- NUNCA SE IMPLEMENTO
End Sub

Private Sub mFSMonitoringCoord_GasFileChanged(...)
    LogDebug MODULE_NAME, "[mFSMonitoringCoord_GasFileChanged] " & fileName
    ' TODO: Actualizar referencias... <- NUNCA SE IMPLEMENTO
End Sub

Private Sub mFSMonitoringCoord_MonitoringReconnected(...)
    LogInfo MODULE_NAME, "[...] " & folder & " tras " & attempts & " intentos"
    ' SOLO LOG
End Sub
```

**Solución:** Mover el log al componente que emite el evento. Eliminar estos eventos.

---

#### 2. Cascada de Eventos Innecesaria

```vba
' Excel.Application dispara evento
m_xlApp_SheetActivated(sh)
    ' clsExecutionContext RE-EMITE
    RaiseEvent SheetActivated(sh)
        ' clsAplicacion ESCUCHA
        ctx_SheetActivated(sh)
            ' clsAplicacion DELEGA
            mChartManager.WatchSheet sh
            evRibbon.InvalidarControl "btnGenerarGraficos"
```

**Problema:**
- 3 niveles de indirección para algo simple
- clsExecutionContext es solo un "relay" de eventos
- clsAplicacion hace el trabajo real

**Solución:**
- clsAplicacion puede suscribirse DIRECTAMENTE a Application vía WithEvents (si es necesario)
- O mejor: usar patrón **Command** para acciones y **Query** para estado

---

#### 3. Eventos para Consultas de Estado

```vba
' clsChartEventsManager
Private Sub mChartManager_ChartActivated(cht As Chart)
    m_bChartActive = True  ' <- SOLO CAMBIA UN FLAG
    evRibbon.InvalidarControl "btnInvertirSeries"
End Sub

Private Sub mChartManager_ChartDeactivated(cht As Chart)
    m_bChartActive = False  ' <- SOLO CAMBIA UN FLAG
    evRibbon.InvalidarControl "btnInvertirSeries"
End Sub
```

**Problema:**
- Evento solo para actualizar un flag booleano
- Se puede consultar directamente cuando se necesite

**Solución:**
```vba
' En lugar de evento + flag, usar Property Get:
Public Property Get HasActiveChart() As Boolean
    HasActiveChart = Not (ctx.Chart Is Nothing)
End Property

' Usar cuando se necesite:
enabled = App.HasActiveChart And EsValidoInvertirEjes()
```

---

#### 4. Eventos de Ribbon que Delegan a Procedimientos

```vba
' clsRibbonEvents
Public Event GenerarGraficosDesdeCurvasRto()
Public Event InvertirEjes()
Public Event FormatearCGASING()

' clsAplicacion escucha y delega
Private Sub evRibbon_GenerarGraficosDesdeCurvasRto()
    Call EjecutarGraficoEnLibroActivo  ' <- Solo delega a procedimiento
End Sub

Private Sub evRibbon_InvertirEjes()
    Call InvertirEjesDelGraficoActivo  ' <- Solo delega a procedimiento
End Sub
```

**Problema:**
- Evento solo para llamar a un procedimiento
- Callback del Ribbon → Evento → Manejador → Procedimiento (4 niveles)

**Solución - Patrón Command:**
```vba
' Interfaz ICommand (simulada con clase base)
' clsCommand.cls
Public Sub Execute()
    ' Implementar en subclases
End Sub

' clsCmdGenerarGraficos.cls
Public Sub Execute()
    Call EjecutarGraficoEnLibroActivo
    App.ChartManager.RefreshCurrentSheet
End Sub

' Callback del Ribbon llama directamente
Public Sub OnGenerarGraficosDesdeCurvasRto(control As IRibbonControl)
    Dim cmd As New clsCmdGenerarGraficos
    cmd.Execute
End Sub
```

---

## ESTRATEGIA DE REFACTORIZACION

### Fase 1: Eliminar Eventos Innecesarios (Quick Win)

**Eventos a eliminar de clsFSMonitoringCoord:**
- TemplateCreated → mover log a clsFSMonitoringCoord
- TemplateChanged → mover log a clsFSMonitoringCoord
- GasFileCreated → mover log a clsFSMonitoringCoord
- GasFileChanged → mover log a clsFSMonitoringCoord
- MonitoringReconnected → mover log a clsFSMonitoringCoord

**Impacto:** -5 eventos, -5 manejadores en clsAplicacion

**Tiempo estimado:** 1 hora

---

### Fase 2: Reemplazar Eventos de Estado por Queries

**Eventos a eliminar:**
- ChartActivated/ChartDeactivated → usar Property Get HasActiveChart

**Cambio:**
```vba
' ANTES (eventos)
Private m_bChartActive As Boolean
Private Sub mChartManager_ChartActivated(cht As Chart)
    m_bChartActive = True
End Sub

' DESPUES (query)
Public Property Get HasActiveChart() As Boolean
    HasActiveChart = Not (ctx.Chart Is Nothing)
End Property
```

**Eventos de ribbon state:**
- StateChanged → consultar mRibbonState.Modo cuando se necesite

**Impacto:** -3 eventos, código más simple

**Tiempo estimado:** 2 horas

---

### Fase 3: Patrón Command para Acciones de Ribbon

**Estructura:**
```
Commands/
  ├─ ICommand (interface simulada con clase base)
  ├─ clsCmdGenerarGraficos
  ├─ clsCmdInvertirEjes
  ├─ clsCmdFormatearCGASING
  ├─ clsCmdConfigurador
  ├─ clsCmdNuevaOportunidad
  └─ clsCmdReplaceNamesValidations
```

**Callbacks simplificados:**
```vba
' modCALLBACKSRibbon.bas
Public Sub OnGenerarGraficosDesdeCurvasRto(control As IRibbonControl)
    ExecuteCommand New clsCmdGenerarGraficos
End Sub

Private Sub ExecuteCommand(cmd As Object)
    On Error GoTo ErrHandler
    cmd.Execute
    Exit Sub
ErrHandler:
    LogError "Ribbon", "ExecuteCommand", Err.Number, Err.Description
End Sub
```

**Ventajas:**
- Elimina 6 eventos de clsRibbonEvents
- Código más testable (cada comando es independiente)
- Facilita agregar funcionalidad (nuevo comando = nueva clase)

**Impacto:** -6 eventos, +6 clases Command (pero más simples)

**Tiempo estimado:** 4 horas

---

### Fase 4: Simplificar ExecutionContext

**Opción A: Eliminar Re-emisión de Eventos**
```vba
' ANTES: clsExecutionContext re-emite todo
Private Sub m_xlApp_SheetActivated(sh As Object)
    RaiseEvent SheetActivated(sh)  ' <- Re-emisión
End Sub

' clsAplicacion escucha
Private WithEvents ctx As clsExecutionContext
Private Sub ctx_SheetActivated(sh As Object)
    mChartManager.WatchSheet sh
End Sub

' DESPUES: clsAplicacion suscribe directamente (si realmente lo necesita)
Private WithEvents m_xlApp As Application
Private Sub m_xlApp_SheetActivated(sh As Object)
    mChartManager.WatchSheet sh
End Sub
```

**Opción B: Mantener ExecutionContext pero simplificar**
- Solo re-emitir eventos que REALMENTE necesitan múltiples suscriptores
- Eliminar eventos que solo tienen 1 suscriptor

**Análisis de suscriptores:**
| Evento | Suscriptores | Acción |
|--------|-------------|--------|
| WorkbookActivated | 1 (clsAplicacion) | Eliminar evento, suscribir directo |
| SheetActivated | 1 (clsAplicacion) | Eliminar evento, suscribir directo |
| SheetDeactivated | 1 (clsAplicacion) | Eliminar evento, suscribir directo |
| WorkbookBeforeClose | 1 (clsFileManager) | ??? Evaluar |

**Impacto:** -3 a -6 eventos, simplifica arquitectura

**Tiempo estimado:** 3 horas

---

### Fase 5: Dependency Injection vs WithEvents

**Problema actual:**
```vba
' clsFileManager
Private WithEvents ctx As clsExecutionContext

Public Sub Initialize(ByVal execCtx As clsExecutionContext)
    Set ctx = execCtx  ' <- Dependencia implícita por eventos
End Sub
```

**Solución A: Inyectar Application directamente**
```vba
' clsFileManager
Private WithEvents m_xlApp As Application

Public Sub Initialize(ByVal xlApp As Application)
    Set m_xlApp = xlApp
End Sub
```

**Solución B: Eliminar WithEvents, usar callbacks explícitos**
```vba
' clsFileManager
Public Sub OnWorkbookClosed(wb As Workbook)
    UntrackFile wb.Name
End Sub

' clsAplicacion (suscriptor directo de Application)
Private Sub m_xlApp_WorkbookBeforeClose(wb As Workbook, Cancel As Boolean)
    mFileMgr.OnWorkbookClosed wb
End Sub
```

**Ventaja:** Dependencias explícitas, más fácil de testear

**Impacto:** Cambia arquitectura, pero simplifica

**Tiempo estimado:** 4 horas

---

## RESUMEN DE IMPACTO

| Fase | Eventos Eliminados | Tiempo | Prioridad |
|------|-------------------|--------|-----------|
| 1. Eliminar eventos innecesarios | -5 | 1h | Alta |
| 2. Queries en lugar de eventos estado | -3 | 2h | Alta |
| 3. Patrón Command para Ribbon | -6 | 4h | Media |
| 4. Simplificar ExecutionContext | -3 a -6 | 3h | Media |
| 5. Dependency Injection | -2 a -4 | 4h | Baja |
| **TOTAL** | **-19 a -24 eventos** | **14h** | |

**Reducción:** De 30+ eventos a ~6-11 eventos (reducción del 60-80%)

---

## ORDEN DE EJECUCION RECOMENDADO

### Sprint 1 (Quick Wins - 3 horas)
1. Fase 1: Eliminar eventos innecesarios
2. Fase 2: Queries en lugar de eventos

### Sprint 2 (Arquitectura - 7 horas)
3. Fase 3: Patrón Command para Ribbon
4. Fase 4: Simplificar ExecutionContext

### Sprint 3 (Opcional - 4 horas)
5. Fase 5: Dependency Injection

---

## CRITERIOS DE EXITO

- [ ] Reducción de al menos 15 eventos del sistema
- [ ] clsAplicacion tiene máximo 10 manejadores (vs 26 actuales)
- [ ] Código más testable (comandos y queries separados)
- [ ] Funcionalidad existente se mantiene 100%
- [ ] No se introducen nuevos bugs

---

## PROXIMOS PASOS

¿Qué fase quieres que implemente primero?

**Recomendación:** Empezar con Fase 1 + Fase 2 (Quick Wins) para ver resultados inmediatos.
