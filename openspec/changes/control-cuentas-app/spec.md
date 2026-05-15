# SDD Spec: Control de Cuentas Personales

## Purpose

App desktop (Flet + openpyxl) para gestión de cuentas familiares. Persistencia directa a `Cuentas 2026.xlsx` sincronizado via Google Drive Desktop. 3 personas (Sil, Gera, Juli), 6 meses (Diciembre-Mayo, expandible), 17 items de gasto en 4 categorías.

---

## 1. Functional Requirements

| ID | Title | Description | Acceptance Criteria |
|----|-------|-------------|-------------------|
| RF-001 | Navegación Mensual | Selector de meses (tabs/dropdown). Muestra gastos + ingresos del mes activo. | Mes seleccionado mustra sus datos. Agregar mes nuevo crea hoja con mismo layout. |
| RF-002 | Expense CRUD | Crear/leer/actualizar/eliminar gastos por mes, agrupados por categoría (TARJETAS, Fijas). | Editar monto persiste al recargar. Nuevo item aparece en categoría correcta. Eliminar solo afecta el mes actual. Items no existentes en otros meses se preservan. |
| RF-003 | Cálculo A Pagar | A_pagar = Monto − Mínimo (si Mínimo > 0); sino A_pagar = Monto. | Monto=100000, Mínimo=25000 → A_pagar=75000. Monto=50000, sin mínimo → A_pagar=50000. |
| RF-004 | Income CRUD | Crear/leer/actualizar/eliminar ingresos por mes (sueldos + extras), asociados a persona. | Editar sueldo persiste al recargar. Nuevo ingreso extra aparece en panel derecho. |
| RF-005 | Status Toggle | Toggle Pagado ↔ Pendiente por item de gasto. | Pendiente→Pagado persiste. Pagado→Pendiente revierte. Cambio visible sin recargar. |
| RF-006 | Resumen Mensual | Muestra TOTAL CUENTAS, total ingresos, SALDO TOTAL. Se actualiza al editar items. | Totales coinciden con Excel al cargar. Al modificar un monto, totales se recalculan. |
| RF-007 | Persistencia Excel | Parser openpyxl: lee todas las hojas al iniciar, escribe cambios preservando layout, merged cells, fórmulas, columnas A-J. | Lee 6 hojas sin error. Escritura preserva estructura y fórmulas. Excel post-escritura abre sin advertencias. |
| RF-008 | Backup Pre-escritura | Crea `Cuentas 2026.backup.{ts}.xlsx` antes de cada escritura. | Backup creado antes de modificar original. Backup funcional (abre en Excel). |

---

## 2. Non-Functional Requirements

| ID | Title | Description |
|----|-------|-------------|
| NFR-001 | Startup | App mustra datos en < 3 segundos desde el lanzamiento. |
| NFR-002 | Compatibilidad | Excel resultante MUST abrirse sin warnings en Excel (no corrupto, fórmulas intactas). |
| NFR-003 | Encoding | Strings con ñ/tildes MUST normalizarse via `unicodedata.normalize('NFC', ...)` al leer/escribir. |
| NFR-004 | Atomic Writes | Escritura MUST usar archivo temporal + rename para evitar corrupción por Google Drive sync. |
| NFR-005 | Validación | Inputs numéricos (Monto, Mínimo) MUST validarse antes de escribir al Excel (tipo, rango, no negativos). |

---

## 3. Scenarios

### E1: Cargar gasto nuevo
- **GIVEN** usuario viendo mes "Enero"
- **WHEN** agrega "Supermercado" con Monto=50000, Pendiente
- **THEN** item aparece en lista de gastos, Excel se actualiza, TOTAL CUENTAS se recalcula

### E2: Marcar como Pagado
- **GIVEN** item con Estado "Pendiente"
- **WHEN** usuario hace clic en "Pagado"
- **THEN** estado cambia a "Pagado" en UI y persiste al cerrar/reabrir app

### E3: Ver resumen del mes
- **GIVEN** usuario selecciona mes con datos
- **WHEN** el mes carga
- **THEN** mustra total_gastos, total_ingresos, saldo; coinciden con Excel

### E4: Agregar mes nuevo
- **GIVEN** existen meses Diciembre–Mayo
- **WHEN** usuario agrega "Junio"
- **THEN** se crea hoja "Junio" con mismo layout y aparece en selector

### E5: Encoding corrupto
- **GIVEN** Excel tiene "M?nimo" en vez de "Mínimo"
- **WHEN** parser lee el archivo
- **THEN** normaliza via unicodedata, UI mustra "Mínimo"

### E6: Modificación manual del Excel
- **GIVEN** Excel fue editado manualmente fuera de la app
- **WHEN** app inicia
- **THEN** lee estado actual y refleja todos los cambios manuales

---

## 4. Edge Cases

| Case | Handling |
|------|----------|
| Item existe en un mes pero no en otro | Modelo `Dict[Mes, MontoItem]` — no lista fija por item |
| Mínimo ausente en algunos items de tarjeta | Regla: si Mínimo > 0 aplica descuento, sino A_pagar = Monto |
| Columna F (pago extra Tuya Gera, 1 celda) | Preservar al leer/escribir, MUST no modificar |
| Columna J (data ambigua, 16 celdas) | Leer y preservar sin modificar |
| Nombre de hoja con espacio trailing ("Marzo ") | Normalizar con `.strip()` al leer |
| Google Drive sync durante escritura | Atomic write (temp + rename) minimiza ventana de conflicto |
| Mes sin items (recién creado) | Mustrar layout vacío con secciones preparadas, sin errores |

---

## 5. Updated Risks

| Risk | Likelihood | Mitigation |
|------|-----------|------------|
| Excel layout change rompe parser | Medium | Tests de integración contra copia del Excel |
| Character encoding (ñ, tildes) | High | Normalizar con unicodedata al leer/escribir |
| Google Drive sync conflict | Low | Atomic writes: temp file + rename |
| Items dinámicos rompen modelo | Low | Dict-based model, no lista fija |
| Usuario abre Excel mientras app escribe | Low | Backup pre-escritura permite recovery manual |
