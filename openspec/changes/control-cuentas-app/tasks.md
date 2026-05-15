# Tasks: Control de Cuentas Personales

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~1,445 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | 6 PRs |
| Delivery strategy | auto-chain |
| Chain strategy | pending |

Decision needed before apply: Yes
Chained PRs recommended: Yes
Chain strategy: pending
400-line budget risk: High

### Suggested Work Units

| Unit | Goal | Likely PR | Base |
|------|------|-----------|------|
| 1 | Foundation (models + utils + tests) | PR 1 | main |
| 2 | Excel Reader + tests | PR 2 | main |
| 3 | Excel Writer + Backup + tests | PR 3 | main |
| 4 | Services + tests | PR 4 | main |
| 5 | UI components | PR 5 | main |
| 6 | App wiring + integration | PR 6 | main |

## Batch 1: Modelos & Utiles

- [ ] 1.1 Crear `src/utils/encoding.py` — NFC normalization helper (~15 loc)
- [ ] 1.2 Crear `src/utils/validation.py` — Monto, Mínimo validators (~25 loc)
- [ ] 1.3 Crear `src/models/expense.py` — ExpenseItem, Category, PaymentStatus, MonthlyExpense, calcular_a_pagar (~80 loc)
- [ ] 1.4 Crear `src/models/income.py` — IncomeItem dataclass (~25 loc)
- [ ] 1.5 Crear `src/models/month.py` — MonthSheet, MonthSummary (~30 loc)
- [ ] 1.6 Crear `src/models/workbook.py` — Workbook container (~25 loc)
- [ ] 1.7 Crear `src/tests/conftest.py` — fixture con copia del Excel real (~40 loc)
- [ ] 1.8 Crear `src/tests/test_models.py` — test calcular_a_pagar, PaymentStatus.from_excel (~50 loc)

## Batch 2: Persistencia

- [ ] 2.1 Crear `src/persistence/excel_reader.py` — openpyxl → Workbook, parsear 6 hojas con secciones (~200 loc)
- [ ] 2.2 Crear `src/persistence/excel_writer.py` — diff-based write, preservar fórmulas, columnas F/J (~130 loc)
- [ ] 2.3 Crear `src/persistence/backup.py` — BackupManager con timestamp (~30 loc)
- [ ] 2.4 Crear `src/tests/test_excel_reader.py` — test lectura 6 hojas, encoding, edge cases (~60 loc)
- [ ] 2.5 Crear `src/tests/test_excel_writer.py` — test write + read-back + idempotencia (~70 loc)

## Batch 3: Servicios

- [ ] 3.1 Crear `src/services/expense_service.py` — CRUD + calcular_a_pagar automático (~60 loc)
- [ ] 3.2 Crear `src/services/income_service.py` — CRUD ingresos (~40 loc)
- [ ] 3.3 Crear `src/services/month_service.py` — navegación, resúmenes, crear mes (~50 loc)
- [ ] 3.4 Crear `src/tests/test_services.py` — test CRUD + resumen contra fixtures (~60 loc)

## Batch 4: UI

- [ ] 4.1 Crear `src/ui/theme.py` — colores, estilos consistentes (~30 loc)
- [ ] 4.2 Crear `src/ui/controls/month_tabs.py` — TabStrip + botón "+" nuevo mes (~70 loc)
- [ ] 4.3 Crear `src/ui/controls/summary_card.py` — TOTAL CUENTAS, Ingresos, SALDO (~40 loc)
- [ ] 4.4 Crear `src/ui/controls/edit_dialog.py` — modal crear/editar con validación (~80 loc)
- [ ] 4.5 Crear `src/ui/controls/expense_table.py` — tabla gastos con toggles Pagado/Pendiente, edit-inline (~120 loc)
- [ ] 4.6 Crear `src/ui/controls/income_panel.py` — lista ingresos editables (~50 loc)

## Batch 5: Integración

- [ ] 5.1 Crear `src/ui/app.py` — layout raíz Flet, montar controles, conectar servicios (~50 loc)
- [ ] 5.2 Crear `src/main.py` — entry point flet.app(target=main) (~30 loc)
