# Exploration: Aplicación de Gestión de Cuentas Personales

## Current State

No existe código aún. El proyecto consiste en una planilla Excel `Cuentas 2026.xlsx` que funciona como sistema manual de control de cuentas familiares. Hay 6 hojas mensuales (Diciembre a Mayo) con una estructura idéntica que combina gastos (tarjetas de crédito + gastos fijos) e ingresos (sueldos) en un mismo layout.

---

## Data Structure — Análisis Exhaustivo del Excel

### General

| Propiedad | Valor |
|-----------|-------|
| Archivo | `Cuentas 2026.xlsx` |
| Hojas | `Diciembre`, `Enero`, `Febrero`, `Marzo ` (con espacio), `Abril`, `Mayo` |
| Rango con datos | Filas 2-25, Columnas A-J (de 15 disponibles A-O) |
| Columnas vacías | I (col 9), K-O (cols 11-15) |
| Celdas combinadas | `A2:E2`, `G2:H2`, `A16:E16`, `G7:H7`, `K24:M24` |

### Layout por Fila

```
Row 1:  (vacía)
Row 2:  [A] TARJETAS DE CREDITO  |  [G] SUELDO SIL        ← Section headers
Row 3:  [A] Item  [B] Monto  [C] Mínimo  [D] A pagar  [E] Estado  |  [G] Guarderia  ← Column headers
Rows 4-10: Items de tarjetas + primeros items de ingreso
Row 11: [G] SUELDO GERA
Row 12: [G] SALDO TOTAL
Rows 13-15: Items de ingreso extra
Row 16: [A] Fijas  (merged A16:E16)                         ← Section header
Rows 17-24: Items de gastos fijos (préstamos, personales)
Row 25: [A] TOTAL DE CUENTAS  [E] <suma total>
```

### Mapeo de Columnas

| Col | Nombre | Tipo | Descripción | Aparece en |
|-----|--------|------|-------------|------------|
| A | Item | `str` | Nombre del item (gasto o sección) | 115 celdas |
| B | Monto | `float` (ocasional `str`) | Monto total del gasto del mes | 77 celdas |
| C | Mínimo | `float` | Pago mínimo (solo tarjetas de crédito) | 18 celdas |
| D | A pagar | `int` | Monto a pagar efectivamente (Monto - Mínimo o Monto completo) | 30 celdas |
| E | Estado | `str` o `int` | Estado de pago o totales de resumen | 89 celdas |
| F | — | `float` | Aparece SOLO en Marzo Row 8 (F=200000.0, pago extra de Tuya Gera) | 1 celda |
| G | Concepto | `str` | Etiquetas de ingresos/resumen (columna derecha) | 76 celdas |
| H | Monto ingreso | `float`/`int` | Valores asociados a col G (sueldos, saldos) | 40 celdas |
| J | Monto extra | `float`/`int` | Valores adicionales, posible data acumulada | 16 celdas |

### Estados Posibles (Col E)

| Estado | Significado |
|--------|-------------|
| `Pagado` | El gasto fue pagado |
| `Pendiente` | El gasto está pendiente de pago |

### Items Únicos (normalizados)

**Tarjetas de Crédito (sección TARJETAS DE CREDITO, rows 4-10):**
1. Tarjeta Naranja
2. Visa NBCH Sil
3. VISA NBCH Gera
4. Tuya Sil
5. Tuya Gera
6. Ingles chicos
7. VISA ICBC

**Items sueltos en sección de tarjetas (rows 13):**
8. Lala Garello

**Fijas / Gastos Fijos (sección Fijas, rows 17-24):**
9. Procrear
10. San Roque
11. Prestamo Sil ANSES
12. Mercado Prestamo
13. Prestamo NBCH
14. Prestamos Naranja
15. Personal Gera
16. Personal Juli
17. Prestamo Sil NBCH (aparece solo en Mayo)

**Item de resumen (row 25):**
18. TOTAL DE CUENTAS

### Secciones del Panel Derecho (Columnas G-J)

| Concepto (G) | Significado |
|--------------|-------------|
| SUELDO SIL | Sección: ingresos de Silvia |
| Guarderia | Sub-item del sueldo |
| Sueldo Neto | Neto salary |
| Total | Total salary |
| Tikes Tuya | Income from Tikes/Tuya |
| Gera | Income asignado a Gera |
| Sil | Income asignado a Sil |
| SUELDO GERA | Sección: sueldo de Gerardo |
| SALDO TOTAL | Cálculo de balance del mes |
| TUYA | Income/ahorro Tuya |
| Cuentas a recuperar | Deudas a cobrar |
| Pedro Bebidas | Income from Pedro |
| Neli | Income from Neli |
| Gestion Alquiler | Income from property management |

---

## Domain Model

### Conceptos del Dominio

```
Persona
  ├── id: str (sil | gera | juli)
  ├── nombre: str

Mes
  ├── id: int (1-12)
  ├── nombre: str (Enero, Febrero...)
  ├── anio: int

ItemGasto
  ├── id: str
  ├── nombre: str
  ├── categoria: CategoriaGasto (TARJETA_CREDITO | FIJO | PRESTAMO | PERSONAL)
  ├── persona_asociada: Persona (opcional — quién es responsable del gasto)
  ├── montos_por_mes: Dict[Mes, MontoItem]

MontosItem
  ├── monto_total: float
  ├── minimo: float (opcional, solo tarjetas)
  ├── a_pagar: float (opcional, calculado o manual)
  ├── pago_extra: float (opcional, col F)
  ├── estado: EstadoPago

Ingreso
  ├── id: str
  ├── concepto: str
  ├── persona: Persona
  ├── monto_por_mes: Dict[Mes, float]

EstadoPago: Pagado | Pendiente

CategoriaGasto: TARJETA_CREDITO | FIJO | PRESTAMO | PERSONAL | OTRO

ResumenMensual
  ├── mes: Mes
  ├── total_gastos: float (TOTAL DE CUENTAS)
  ├── total_ingresos: float
  ├── saldo: float (SALDO TOTAL = ingresos - gastos)
  ├── cuentas_a_recuperar: float
```

### Reglas de Negocio Descubiertas

1. **Cálculo de "A pagar"**: Para tarjetas de crédito: `A_pagar = Monto - Mínimo` (cuando existe mínimo). Si no hay mínimo, `A_pagar = Monto`.
2. **TOTAL DE CUENTAS**: Suma de todos los montos (col B) de items con estado, excluyendo los totales mismos. Es el total de egresos del mes.
3. **SALDO TOTAL**: Diferencia entre ingresos totales y gastos totales. Generalmente negativo (los gastos superan los ingresos).
4. **Items dinámicos**: Algunos items aparecen/desaparecen entre meses (ej: Prestamo Sil NBCH solo en Mayo, Personal Gera aparece hasta Abril y desaparece en Mayo).
5. **Estados por mes**: Un mismo item puede tener estado "Pagado" un mes y "Pendiente" al siguiente.
6. **Personas**: Sil (Silvia) y Gera (Gerardo) son las dos personas del hogar. Juli aparece como "Personal Juli" — posiblemente un préstamo a un tercero.
7. **Columna J**: Datos adicionales que parecen ser montos acumulados/proyectados. Su lógica exacta no está clara del Excel solo.

---

## Architecture Notes

### Estrategia de Modelado

Propongo un modelo **híbrido** donde:
1. **Capa de datos en memoria**: Modelo de datos Python con clases `Month`, `ExpenseItem`, `IncomeItem`, `Person`
2. **Capa de persistencia**: openpyxl como backend de lectura/escritura sincronizado con el Excel
3. **Capa de UI**: Flet (Python) como interfaz desktop/web

### Excel como "Base de Datos"

El Excel no es ideal como DB, pero es el source of truth actual. Estrategia:
- **READ**: Cargar todo el Excel en memoria al iniciar la app (todas las hojas)
- **WRITE**: Escribir cambios de vuelta al Excel en el mismo formato (para mantener compatibilidad con Google Drive sync)
- **Cache**: Mantener un modelo en memoria para operaciones rápidas en la UI
- **Sync**: Cada cambio en la UI escribe al Excel inmediatamente

### Operaciones CRUD que la UI necesita

| Operación | Descripción |
|-----------|-------------|
| `list_items(mes)` | Listar todos los items de gasto de un mes |
| `get_item(mes, item_id)` | Obtener detalle de un item |
| `update_estado(mes, item_id, estado)` | Marcar como Pagado/Pendiente |
| `update_monto(mes, item_id, monto)` | Actualizar monto de un item |
| `add_item(mes, item)` | Agregar un nuevo gasto al mes |
| `remove_item(mes, item_id)` | Eliminar un gasto |
| `list_ingresos(mes)` | Listar ingresos del mes |
| `update_ingreso(mes, concepto, monto)` | Actualizar un ingreso |
| `get_resumen(mes)` | Obtener totales y saldo del mes |
| `compare_months(meses)` | Comparar gastos entre meses |

### Migración desde Planilla Actual

1. **Fase 0 (exploración)**: Ya estamos acá — entender la estructura
2. **Fase 1 (lectura)**: Parser en Python que lea todas las hojas y construya el modelo en memoria
3. **Fase 2 (UI básica)**: Flet app que muestre la data ya parseada, permitiendo ver y cambiar estados
4. **Fase 3 (escritura)**: Escribir cambios al Excel manteniendo el formato exacto
5. **Fase 4 (features)**: Agregar features que el Excel no tiene (comparativas, gráficos, alertas)

### Riesgos

1. **Excel como DB es frágil**: Si el usuario mueve columnas o cambia el layout, el parser se rompe
2. **Encoding**: Los caracteres especiales (ñ, tildes) aparecen corruptos en algunos strings del Excel (ej: "Mínimo" se lee como "M?nimo")
3. **Concurrencia**: Si Google Drive Desktop sincroniza mientras la app escribe, hay riesgo de corrupción
4. **Columna J ambigua**: No sabemos con certeza qué representa — puede ser data acumulada, proyección, o notas
5. **Items dinámicos**: El set de items cambia entre meses, hay que soportar que no todos los items existan en todos los meses
6. **Layout rígido**: El Excel usa celdas combinadas, lo que hace el parser más complejo

---

## Ready for Proposal

Sí — el dominio está claro, la estructura de datos está completamente analizada, y hay suficiente información para diseñar la arquitectura y proponer el cambio.

---

## Affected Areas

No hay código existente. El "sistema actual" es el archivo Excel:
- `Cuentas 2026.xlsx` — fuente de datos única
