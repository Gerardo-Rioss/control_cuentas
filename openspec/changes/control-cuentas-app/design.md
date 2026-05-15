# Design: Control de Cuentas Personales — App Flet + openpyxl

## Technical Approach

Arquitectura limpia por capas con carga total del Excel en memoria al startup y escritura
síncrona por cada mutación. La UI (Flet) es el driver: cada acción del usuario → valida →
actualiza el modelo en memoria → persiste al Excel vía tempfile + rename con backup
pre-escritura. El modelo usa dicts por item_id para soportar items dinámicos entre meses.
openpyxl opera en modo no-destructive: preserva merged cells, fórmulas (como strings),
formatos, y columnas no administradas (F, J).

## Architecture Decisions

| Decisión | Opciones | Tradeoff | Resolución |
|----------|----------|----------|------------|
| Carga total vs lazy-load | (a) Todo en RAM al startup, (b) Load por mes bajo demanda | (a) startup ~1.5s, pero operaciones rápidas sin I/O; (b) más rápido al inicio pero cada cambio de mes requiere I/O | **(a) Carga total** — Excel es pequeño (~50KB), 6 hojas, cabe en RAM sin problema |
| Flet DataTable vs Column+Row | (a) DataTable nativo, (b) Custom con Column/Row/TextField | DataTable da ordenamiento y selección gratis pero es rígido con edit-inline; (b) es más flexible pero más código | **(b) Custom con Column/Row** — necesitamos edit-in-place y toggles, DataTable no lo soporta bien |
| Excel formulas vs valores calculados | (a) Preservar fórmulas como strings, (b) Evaluar en Python y escribir solo valores | (a) mantiene compatibilidad total con Excel, (b) la app no dependería de Excel para cálculos | **(a) Preservar fórmulas** — RF-007 exige que el Excel post-escritura funcione standalone |
| Dict vs Lista plana para items | (a) Dict[item_id, MonthlyExpense], (b) Lista plana | Dict permite lookup O(1) y maneja items que aparecen/desaparecen naturalmente | **(a) Dict** — exploración confirmó items dinámicos entre meses |
| Atomic writes con copia vs rename | (a) shutil.copy2 + os.remove + os.rename, (b) tempfile.NamedTemporaryFile + replace | (b) es más seguro porque el archivo temporal está en el mismo filesystem, rename es atómico en NTFS | **(b) tempfile + replace** — NFR-004, minimiza ventana de conflicto con Google Drive sync |

## Data Flow

```
┌─────────────────────────────────────────────────────────────────────┐
│                        APP STARTUP                                  │
│                                                                     │
│  main.py ──→ ExcelReader.read_all() ──→ Workbook (models) ──→ UI    │
│                   │                           │                     │
│              openpyxl.load_workbook()    MonthSheet[]               │
│              (data_only=False)           ExpenseItem[]              │
│                                          IncomeItem[]               │
└─────────────────────────────────────────────────────────────────────┘
                               │
                          User edits UI
                               │
                               ▼
┌─────────────────────────────────────────────────────────────────────┐
│                        EDIT FLOW                                    │
│                                                                     │
│  UI callback → Service.validate() → update model → call write()     │
│                                                       │             │
│                                            ┌──────────┘             │
│                                            ▼                        │
│  ExcelWriter.write_month(month_name):                              │
│    1. BackupManager.create() → "Cuentas 2026.backup.{ts}.xlsx"     │
│    2. tempfile.NamedTemporaryFile(delete=False)                     │
│    3. Escribir celdas modificadas en la hoja específica            │
│    4. wb.save(temp_path) → os.replace(temp, original)              │
│    5. UI refresh                                                    │
└─────────────────────────────────────────────────────────────────────┘
```

## File Structure

```
C:\Users\grios\Desktop\Proyecto_Prueba\control_cuentas\
├── src/
│   ├── main.py                    # Entry point: flet.app(target=main)
│   ├── models/
│   │   ├── __init__.py
│   │   ├── expense.py             # ExpenseItem, Category, PaymentStatus
│   │   ├── income.py              # IncomeItem
│   │   ├── month.py               # MonthSheet, MonthSummary
│   │   └── workbook.py            # Workbook (colección de MonthSheet)
│   ├── services/
│   │   ├── __init__.py
│   │   ├── expense_service.py     # CRUD + cálculo A_pagar
│   │   ├── income_service.py      # CRUD
│   │   └── month_service.py       # Navegación, resúmenes
│   ├── persistence/
│   │   ├── __init__.py
│   │   ├── excel_reader.py        # openpyxl → Workbook
│   │   ├── excel_writer.py        # MonthSheet → openpyxl
│   │   └── backup.py              # BackupManager
│   ├── ui/
│   │   ├── __init__.py
│   │   ├── app.py                 # page.on_route_change + layout raíz
│   │   ├── controls/
│   │   │   ├── month_tabs.py      # TabStrip + botón "+"
│   │   │   ├── expense_table.py   # Tabla de gastos con toggles
│   │   │   ├── income_panel.py    # Panel de ingresos
│   │   │   ├── summary_card.py    # Totales + saldo
│   │   │   └── edit_dialog.py     # Diálogo modal crear/editar
│   │   └── theme.py              # Colores, estilos consistentes
│   └── utils/
│       ├── __init__.py
│       ├── encoding.py            # unicodedata.normalize('NFC', ...)
│       └── validation.py          # validar_monto(), validar_estado()
├── tests/
│   ├── conftest.py                # Fixture: copia del Excel real
│   ├── test_excel_reader.py       # Lectura de 6 hojas
│   ├── test_excel_writer.py       # Escritura + verificación idempotente
│   ├── test_models.py             # Cálculo A_pagar, validación
│   └── test_services.py           # CRUD + resumen
├── Cuentas 2026.xlsx
└── requirements.txt
```

## Data Model

```python
# ─── Enums ───────────────────────────────────────────────

class PaymentStatus(str, Enum):
    PAGADO = "Pagado"
    PENDIENTE = "Pendiente"

    @classmethod
    def from_excel(cls, value: Any) -> Optional["PaymentStatus"]:
        if not value: return None
        v = str(value).strip().lower()
        if v in ("pagado", "1"): return cls.PAGADO
        if v in ("pendiente", "0"): return cls.PENDIENTE
        return None

class Category(str, Enum):
    TARJETA_CREDITO = "TARJETAS DE CREDITO"
    FIJO = "Fijas"

class Person(str, Enum):
    SIL = "sil"
    GERA = "gera"
    JULI = "juli"

# ─── Expense ─────────────────────────────────────────────

@dataclass
class ExpenseItem:
    """Un ítem de gasto individual (ej: 'Tarjeta Naranja')."""
    item_id: str                    # normalized slug: "tarjeta-naranja"
    nombre: str                     # display name from Col A
    categoria: Category
    persona: Optional[Person] = None

@dataclass
class MonthlyExpense:
    """Valores de un ExpenseItem para un mes específico."""
    item: ExpenseItem
    monto: Optional[float]          # Col B
    minimo: Optional[float]         # Col C (solo tarjetas)
    a_pagar: Optional[float]        # Col D (None = calcular)
    estado: Optional[PaymentStatus] # Col E
    pago_extra: Optional[float]     # Col F (preservar)
    excel_row: int                  # Fila en la hoja

    def calcular_a_pagar(self) -> Optional[float]:
        if self.a_pagar is not None:
            return self.a_pagar
        if self.monto is None:
            return None
        if self.minimo and self.minimo > 0:
            return self.monto - self.minimo
        return self.monto

# ─── Income ──────────────────────────────────────────────

@dataclass
class IncomeItem:
    """Item de ingreso (sueldo, TUYA, cuentas a recuperar)."""
    concepto: str                   # Col G
    monto: Optional[float]          # Col H
    persona: Optional[Person]
    col_j: Optional[float]          # Col J (preservar)
    excel_row: int

# ─── Month ───────────────────────────────────────────────

@dataclass
class MonthSummary:
    total_gastos: float             # Row 25, E
    total_ingresos: float           # Suma de ingresos
    saldo: float                    # Row 12, H (ingresos - gastos)

@dataclass
class MonthSheet:
    nombre: str                     # "Enero", "Febrero"...
    expenses: Dict[str, MonthlyExpense]  # key=item_id
    incomes: List[IncomeItem]
    summary: MonthSummary
    # Preservación de datos no administrados
    raw_col_f: Dict[int, Any]       # row → valor F
    raw_col_j: Dict[int, Any]       # row → valor J
    merged_cells: List[str]         # lista de rangos combinados

# ─── Workbook ────────────────────────────────────────────

@dataclass
class Workbook:
    months: Dict[str, MonthSheet]   # key = nombre normalizado
    filepath: str
    all_items: Dict[str, ExpenseItem]  # catálogo maestro de items
    # Referencia al workbook openpyxl abierto (data_only=False)
    _wb: Any  # openpyxl.Workbook (no serializable)

    @property
    def month_names(self) -> List[str]:
        return list(self.months.keys())
```

## Excel Parser (excel_reader.py)

### Algoritmo de Lectura

```
read_all(filepath) → Workbook:
  1. wb = load_workbook(filepath, data_only=False)
  2. Para cada sheet_name en wb.sheetnames:
     a. nombre = sheet_name.strip()  // normalizar espacios trailing
     b. ws = wb[nombre]
     c. merged = list(ws.merged_cells.ranges)
     d. Parsear secciones del panel izquierdo (A-E):
        - Row 2: header "TARJETAS DE CREDITO"
        - Rows 4-10: items de tarjetas (si A no es None)
           • Col B → monto, Col C → mínimo, Col D → a_pagar,
             Col E → estado, Col F → pago_extra
        - Row 13: item suelto "Lala Garello" (si A no es None)
        - Row 16: header "Fijas" (merged A16:E16)
        - Rows 17-24: items fijos/préstamos/personales
        - Row 25: "TOTAL DE CUENTAS", Col E → fórmula/valor
     e. Parsear panel derecho (G-J):
        - Row 2: "SUELDO SIL"
        - Rows 3-10: items de ingreso de Sil
        - Row 11: "SUELDO GERA", H → valor
        - Row 12: "SALDO TOTAL", H → fórmula/valor
        - Rows 13-16: items de ingreso extra
        - Col J: preservar raw value por fila
     f. Construir MonthSheet
  3. Construir catálogo all_items desde todos los meses
  4. Devolver Workbook
```

**Identificación de items**: La fila donde aparece un item define su categoría.
- Rows 4-10 → Category.TARJETA_CREDITO
- Si el nombre contiene "Préstamo" o "Prestamo" → asignar persona según contexto
- Si el nombre contiene "Personal" → asignar persona según el nombre
- Rows 17-24 → Category.FIJO (y subcategorizar por nombre)

**Manejo de dinámicos**: Items nuevos (ej: "Prestamo Sil NBCH" en Mayo) se agregan
automáticamente al catálogo. Items que desaparecen (ej: "Personal Gera" en Mayo) se
marcan con monto=None en lugar de borrarse.

### Algoritmo de Escritura

```
write_month(workbook, month_name) → None:
  1. BackupManager.create(filepath)  // backup pre-escritura
  2. wb = load_workbook(filepath, data_only=False)  // recargar fresco
  3. ws = wb[month_name]
  4. Para cada expense en workbook.months[month_name].expenses.values():
     ws[f'B{expense.excel_row}'] = expense.monto
     ws[f'C{expense.excel_row}'] = expense.minimo
     ws[f'D{expense.excel_row}'] = expense.a_pagar  (None = preservar fórmula)
     ws[f'E{expense.excel_row}'] = expense.estado.value si existe
  5. Para cada income en workbook.months[month_name].incomes:
     ws[f'H{income.excel_row}'] = income.monto
  6. temp = tempfile.NamedTemporaryFile(
       dir=os.path.dirname(filepath), suffix='.xlsx', delete=False)
  7. wb.save(temp.name)
  8. os.replace(temp.name, filepath)  // atómico en NTFS
```

**Regla crítica**: Solo se escriben las celdas que el usuario modificó. Si una celda
tiene fórmula (ej: `D5 = SUM(B5-C5)`), se preserva la fórmula a menos que el usuario
edite explícitamente el valor. Esto se logra cargando con `data_only=False`.

**Nuevos items**: Se escriben en la primera fila vacía dentro de su sección, y se
actualiza la fórmula de TOTAL DE CUENTAS para incluir la nueva fila.

## UI Components (Flet)

```
┌────────────────────────────────────────────────────────────────┐
│  Page                                                          │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │  AppBar: "Control de Cuentas 2026"                      │  │
│  ├──────────────────────────────────────────────────────────┤  │
│  │  MonthTabs: [Dic] [Ene] [Feb] [Mar] [Abr] [Mayo] [+]   │  │
│  ├────────────────────────┬─────────────────────────────────┤  │
│  │  ExpenseTable          │  IncomePanel                    │  │
│  │                        │                                 │  │
│  │  ┌── TARJETAS ──────┐  │  ┌── SUELDO SIL ────────────┐  │  │
│  │  │ Item   Mnt Min AP │  │  │ Sueldo Neto:  740000    │  │  │
│  │  │ ○ Njn  1M  -  1M  │  │  │ Total:        740000    │  │  │
│  │  │ ● Visa 47k -  47k │  │  ├── SUELDO GERA ─────────┤  │  │
│  │  │ ...               │  │  │ Gera:        1548000    │  │  │
│  │  └───────────────────┘  │  ├── EXTRAS ──────────────┤  │  │
│  │  ┌── FIJAS ──────────┐  │  │ TUYA:         385000    │  │  │
│  │  │ ○ Proc 316k - 316k│  │  │ Cuentas rec:  0        │  │  │
│  │  │ ...               │  │  └─────────────────────────┘  │  │
│  │  └───────────────────┘  │                                 │  │
│  │  TOTAL: $4,200,000 [+]  │                                 │  │
│  ├────────────────────────┴─────────────────────────────────┤  │
│  │  SummaryCard: Gastos $4.2M | Ingresos $2.3M | Saldo -$1.9M│ │
│  └──────────────────────────────────────────────────────────┘  │
└────────────────────────────────────────────────────────────────┘
```

### Componente: MonthTabs
- `Tab` por cada mes en `Workbook.month_names`
- Botón "+" al final que dispara `on_add_month` → crea hoja nueva copiando layout
- Al cambiar tab: `page.update()` con los datos del mes seleccionado
- Props: `months: List[str]`, `active: str`, `on_change: Callable[[str], None]`

### Componente: ExpenseTable
- Dos secciones: TARJETAS DE CREDITO y Fijas
- Cada fila es un `Row` con:
  - `IconButton` (check/circle) para toggle Pagado/Pendiente
  - `Text` (nombre del item)
  - `TextField` (monto, editable inline)
  - `Text` (mínimo, si aplica)
  - `Text` (a pagar, calculado automáticamente)
- Color de fondo alterna según estado (verde claro = Pagado, blanco = Pendiente)
- Botón "+" por sección para agregar nuevo item

### Componente: IncomePanel
- Columnas G-H del Excel renderizadas como lista vertical
- `TextField` editable para montos (H)
- Secciones: SUELDO SIL, SUELDO GERA, ingresos extra (TUYA, Cuentas a recuperar, etc.)

### Componente: SummaryCard
- Tres métricas en horizontal:
  - TOTAL CUENTAS (rojo si > ingresos)
  - Total Ingresos (verde)
  - SALDO TOTAL (rojo si negativo, verde si positivo)
- Se actualiza en tiempo real al modificar cualquier monto

### Componente: EditDialog
- Modal overlay para crear/editar items
- Campos: Nombre (solo crear), Monto, Mínimo (solo tarjetas), Estado (dropdown)
- Validación: monto ≥ 0, tipo numérico

## Error Handling

| Caso | Dónde | Manejo |
|------|-------|--------|
| Excel no encontrado | Startup | Mostrar error con ruta completa, ofrecer selector de archivo |
| Encoding corrupto (M?nimo) | Reader | `unicodedata.normalize('NFC')` en todos los strings leídos |
| Celda con fórmula no evaluable | Reader | Preservar string de fórmula, mostrar "—" en UI |
| Error de escritura (permiso) | Writer | Capturar PermissionError, mostrar snackbar "Archivo en uso. Cerralo e intentá de nuevo." |
| Backup falla | Writer | Loggear warning, continuar sin backup (no bloquear escritura) |
| Item duplicado al crear | Service | Validar por nombre normalizado antes de insertar |
| Monto inválido | Validation | Mostrar error en línea (borde rojo en TextField), no persiste |app. |
| Google Drive sync parcial | Writer | Atomic write minimiza ventana; backup permite recovery manual |

## Risks

| Risk | Likelihood | Mitigation |
|------|-----------|------------|
| Excel layout change rompe parser | Medium | Tests de integración contra copia del Excel + validación de estructura al startup |
| Character encoding (ñ, tildes) | High | `unicodedata.normalize('NFC', ...)` en reader y writer |
| Google Drive sync conflict | Low | Atomic write con tempfile + os.replace en mismo filesystem |
| Items dinámicos rompen modelo | Low | Dict-based model + registro de items por presencia real |
| Usuario abre Excel mientras app escribe | Low | Backup pre-escritura permite recovery manual |
| Fórmulas que referencian filas fuera de rango | Medium | Writer ajusta fórmulas de totales al agregar/quitar filas |
| Flet version compatibility | Low | Pin version exacta en requirements.txt |

## Testing Strategy

| Layer | What | Approach |
|-------|------|----------|
| Unit | ExpenseItem.calcular_a_pagar() | 3 casos: con mínimo >0, sin mínimo, monto None |
| Unit | PaymentStatus.from_excel() | strings, ints, None, bordercases |
| Unit | encoding.normalize() | strings con ñ/tildes vs corruptos |
| Integration | ExcelReader.read_all() | Contra copia real del Excel, verificar 6 hojas |
| Integration | ExcelWriter.write_month() | Escribir, leer de vuelta, comparar valores vs original |
| Integration | BackupManager | Verificar archivo creado, contenido válido |
| E2E | App startup carga datos | Flet test runner o manual |

## Migration / Rollout

No migration required. El Excel es el source of truth — la app se sienta encima y
empieza a funcionar inmediatamente. La primera ejecución es idempotente: lee el
Excel, construye el modelo, muestra la UI.

## Open Questions

- [ ] **Columna J**: Confimar semántica con el usuario. Por ahora se preserva sin modificar.
- [ ] **Nuevos items en sección TARJETAS vs Fijas**: El parser usa posición de fila para determinar categoría. Si el layout cambia (ej: items entremezclados), el parser falla. Validar si esta regla es aceptable.
- [ ] **Fórmulas en col D**: Algunas son `=B6-C6`, otras `=SUM(B5-C5)`, otras `=SUM(B8,-C8)`. El writer debe preservar la fórmula exacta, no el patrón.
