# Proposal: Aplicación de Gestión de Cuentas Personales

## Intent

Sil y Gera llevan el control de sus cuentas familiares en un Excel manual (`Cuentas 2026.xlsx`). No hay validaciones, no hay vista rápida de resúmenes, y marcar items como Pagado/Pendiente requiere abrir el archivo y editar celdas. Queremos una app desktop que facilite la carga visual de datos, muestre resúmenes mensuales, y persista TODO en el mismo Excel para mantener compatibilidad con Google Drive sync.

## Scope

### In Scope
- UI desktop (Flet) para ver y editar gastos e ingresos por mes
- CRUD de items de gasto (tarjetas, fijos, préstamos, personales)
- CRUD de ingresos (sueldos, otros conceptos)
- Vista de resumen mensual (totales, saldo)
- Marcar items como Pagado/Pendiente
- Lectura/escritura directa al Excel manteniendo formato exacto
- Soporte para items que aparecen/desaparecen entre meses

### Out of Scope
- Reportes gráficos o charts
- Multi-usuario o permisos
- Sincronización cloud propia (usa Google Drive Desktop existente)
- Backup automático
- Exportación a otros formatos

## Capabilities

### New Capabilities
- `expense-management`: CRUD de gastos por mes con soporte de categorías (tarjetas, fijos, préstamos, personales) y cálculo automático de "A pagar"
- `income-management`: CRUD de ingresos por mes (sueldos, conceptos varios) asociados a persona
- `monthly-summary`: Vista de total de gastos, total de ingresos, y saldo del mes
- `status-tracking`: Cambiar estado de items entre Pagado y Pendiente
- `excel-persistence`: Parser/escritor openpyxl que lee y escribe el Excel manteniendo layout, merged cells y fórmulas

### Modified Capabilities
- None (no existing specs)

## Approach

Arquitectura por capas. En memoria: modelo Python con clases `Month`, `ExpenseItem`, `IncomeItem` con reglas de negocio (ej: cálculo de "A pagar"). Persistencia: openpyxl lee todo el Excel al iniciar y escribe cambios de vuelta manteniendo el formato exacto. UI: Flet (componentes: `DataTable`, `Dropdown`, botones de acción) con dos paneles (gastos + ingresos) y selector de mes vía tabs.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/` | New | Código fuente de la app (modelos, persistencia, UI) |
| `Cuentas 2026.xlsx` | Modified | La app escribe sobre este archivo |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Excel layout change rompe parser | Medium | Tests de integración contra copia del Excel |
| Google Drive sync conflict al escribir | Low | Atomic writes: escribir a temp file y renombrar |
| Character encoding (ñ, tildes) | High | Normalizar strings con `unicodedata` al leer/escribir |
| Items dinámicos rompen modelo | Low | Modelo con `Dict[Mes, MontoItem]`, no lista fija |

## Rollback Plan

Cada escritura al Excel se hace sobre una copia de seguridad automática (`Cuentas 2026.backup.xlsx`). Si algo sale mal, el usuario renombra el backup y recupera el estado anterior. El parser valida la estructura esperada antes de escribir.

## Dependencies

- Python 3.14.5, openpyxl 3.1.5, Flet (latest compatible)
- `Cuentas 2026.xlsx` en disco sincronizado con Google Drive Desktop

## Success Criteria

- [ ] App levanta y muestra el resumen del mes activo en < 3 segundos
- [ ] Se puede cambiar estado de un item de "Pendiente" a "Pagado" y persiste al recargar
- [ ] Se puede editar un monto y el cambio sobrevive a cerrar/reabrir la app
- [ ] El Excel resultante se puede abrir manualmente con la misma estructura que antes
- [ ] Tests de integración cubren lectura de los 6 meses existentes sin errores
