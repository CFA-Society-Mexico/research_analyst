# Skill `research-note` y comando `/present-coverage` — plan de implementación (PR 2)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Dar dueña a la etapa 7 del flujo: una skill que enseña a armar el deck de
una cobertura sobre `tools/deck_builder.py`, el comando que la orquesta con dos
gates del analista, y los contratos del repo al día. Cierra el issue 54.

**Architecture:** Solo Markdown y contratos; el tool ya está en `main` (PR 56). La
skill vive en ~110 líneas y manda el material largo (API real del tool, regla de
bloqueo, checks P1 a P10, verificación) a `references/deck-format.md`. El comando
es orquestación pura con el formato de `commands/update-industry.md`.

**Tech Stack:** Markdown, `unittest` de stdlib (`tests/test_repo_contracts.py`),
`tools/deck_builder.py` sin cambios.

Spec: `docs/superpowers/specs/2026-11-26-research-note-deck-design.md`, secciones
D1, D4 y la tabla "Contratos que cambian" de D5.

---

## Estado de partida (verificado el 2026-11-26)

- PR 56 fusionado (`d94c570`). Rama `feat/research-note-skill` sale de
  `origin/main`, en el worktree `.claude/worktrees/research-note-skill`.
- Línea base (antes del rebase): `python -m unittest discover -s tests` -> 430 pruebas, OK (4
  saltadas); `python tools/deck_builder.py demo ./scratch` -> VERDE.
- **`main` ya está en v0.8.0** (`3007b01`, check C11). El spec y el encargo dicen
  "bump a v0.8.0": ese número ya se usó. Este PR sube a **v0.9.0** (`minor`:
  capacidad nueva, CONTRIBUTING 7).
- El PR 58 (`feat/contract-check`, abierto) toca `AGENTS.md`, `CONTRIBUTING.md`,
  `commands/update-quarter.md` y `tests/test_repo_contracts.py`. El que se fusione
  segundo rebasa; los cambios no se pisan por contenido, solo por cercanía.
- El issue 54 no tiene comentarios: la enmienda de 8 a 9 skills no está registrada.

## Decisiones que el spec no cubre

| # | Tema | Decisión | Porqué |
|---|---|---|---|
| 1 | Versión | v0.9.0 | v0.8.0 ya está en `main` |
| 2 | `build_deck.py` y "nada se borra" | Nombre fijo para el vigente. Antes de reescribirlo para un deck con `as_of` nuevo, se guarda copia `build_deck_<as_of anterior>.py` | El script es lo que reproduce un deck ya entregado; sobrescribirlo rompe R5. Mismo trato que el respaldo del modelo en `/update-quarter` |
| 3 | Regenerar con el mismo `as_of` | Permitido hasta el gate 2. Después del gate 2, un deck nuevo lleva `as_of` nuevo | El HTML fechado es el entregable aprobado |
| 4 | Recalcular el modelo en disco | Lo hace el preflight del comando, con el sí del analista, vía `python tools/recalc_validate.py <modelo> --in-place` | `research-note` lee el modelo y nunca lo escribe. R1 garantiza que el recálculo no cambia fórmulas |
| 5 | Familia de checks P | Documentada en `references/deck-format.md` (no en `integrity-checks.md`, que es del modelo), con fila en CONTRIBUTING 4.2 y una prueba de contrato "P documentados = P emitidos" | Misma garantía que ya tienen F y S contra la deriva |
| 6 | Gates del analista en el dogfood | Se corre con un agente sin contexto que solo recibe el comando; los gates quedan registrados como simulados. El analista revisa las capturas al final | Quien escribió los dogfoods del PR 1 conocía la regla de bloqueo; hace falta medir la fricción con alguien que no |

## Mapa de archivos

| Archivo | Acción | Responsabilidad |
|---|---|---|
| `skills/research-note/SKILL.md` | Crear | Cuándo corre, inputs, artefactos, procedimiento de 4 pasos, reglas duras, entrevista de cierre |
| `skills/research-note/references/deck-format.md` | Crear | Guion, registro de hechos, API del tool, regla de bloqueo y `lit()`, ciclo de medición, checks P1 a P10, verificación del gate 2, límites |
| `commands/present-coverage.md` | Crear | Preflight, pipeline de 4 pasos, 2 gates, post-run, reporte final |
| `tests/test_repo_contracts.py` | Modificar | Prueba nueva: checks P documentados = emitidos |
| `AGENTS.md` | Modificar | 2 filas de enrutamiento; 1 viñeta en restricciones de entorno |
| `README.md` | Modificar | 6 comandos, 9 skills, fila de la skill, árbol, extras |
| `CONTRIBUTING.md` | Modificar | 4.6 (9 skills y la razón), 4.2 (familia P), sección 2 (tools) |
| `docs/architecture.md` | Modificar | Etapa 7 en el diagrama, fila de dueña, fila de decisión, estado |
| `templates/coverage-tree.md` | Modificar | Archivos de `research/notes/`; slots `deck_*` |
| `commands/update-quarter.md` | Modificar | Post-run: `audit` del deck vigente |
| `docs/superpowers/specs/2026-11-26-research-note-deck-design.md` | Modificar | v0.9.0 en D5; estado del PR 2 |
| `.claude-plugin/plugin.json`, `.codex-plugin/plugin.json`, `README.md` | Modificar | v0.9.0, commit propio |

`tools/deck_builder.py` no se toca. Si el dogfood destapa un defecto real: prueba
que falle primero, arreglo después, commit `fix:` aparte y nota en el PR.

---

### Task 1: Prueba de contrato de los checks P (falla primero)

**Files:**
- Modify: `tests/test_repo_contracts.py`

- [ ] **Step 1: Escribir la prueba**

Constantes junto a `CHECKER`:

```python
DECK_BUILDER = ROOT / "tools" / "deck_builder.py"
DECK_DOC = SKILLS_DIR / "research-note" / "references" / "deck-format.md"
_P_ROW = re.compile(r"^\| (P\d+) \|", re.MULTILINE)
```

Método nuevo en `IntegrityCheckTests`:

```python
    def test_documented_p_checks_are_exactly_the_emitted_ones(self) -> None:
        documented = _P_ROW.findall(DECK_DOC.read_text(encoding="utf-8"))
        self.assertEqual(sorted(set(documented)), sorted(documented),
                         "los IDs de check P no se repiten")
        emitted = set(re.findall(r'Finding\("(P\d+) ',
                                 DECK_BUILDER.read_text(encoding="utf-8")))
        self.assertEqual(set(documented), emitted,
                         "checks P documentados vs emitidos por tools/deck_builder.py")
```

(`"P5-P8 medicion de espacio"`, el hallazgo de "sin navegador", no coincide con el
patrón por el guion: es un estado, no un check.)

- [ ] **Step 2: Verla fallar**

Run: `python -m unittest tests.test_repo_contracts -v`
Expected: ERROR en la prueba nueva, `FileNotFoundError` de `deck-format.md`.

No se hace commit todavía: entra con la Task 2, que la pone en verde.

### Task 2: Skill `research-note`

**Files:**
- Create: `skills/research-note/SKILL.md`
- Create: `skills/research-note/references/deck-format.md`
- Modify: `AGENTS.md`, `README.md`, `CONTRIBUTING.md` (lo mínimo para que
  `test_repo_contracts` siga verde con la carpeta nueva: ruta, fila y conteo)

- [ ] **Step 1: `references/deck-format.md`**

Secciones, en este orden, con este contenido (todo verificado contra
`tools/deck_builder.py`, no contra el spec):

1. **Guion.** Orden fijo (qué es, qué la mueve, qué tan rentable es, cuánto vale),
   actos adaptados a la emisora. Formato del outline: una fila por lámina con
   número, acto, tipo (`cover` / `section` / `slide` / `raw` / `legal`), hallazgo
   (la frase del título, sin cifra tecleada: la cifra se nombra por su origen,
   p. ej. "crecimiento de Ingresos netos 2027"), visual, series y celdas del
   modelo (`hoja` + etiqueta), hechos externos (clave). Anatomía de `slide`:
   título = hallazgo, entradilla, visual, Nota de 2 o 3 frases, cierre, fuente.
   Un concepto por lámina; cada término técnico se explica la primera vez.
2. **Registro de hechos.** Esquema exacto de `<ticker>-facts_<fecha>.json`
   (`as_of`, `facts[]` con `key`, `value`, `unit`, `tag`, `doc`, `locator`,
   `quote`). Reglas: `key` solo letras, dígitos y `_`; `tag` `observado` o
   `guidance`; `doc` relativo a la cobertura y existente; `quote` de al menos 2
   caracteres, **copiada del texto visible** (en un filing iXBRL, de lo que se lee
   en pantalla, no del HTML crudo); fuente no textual (PDF) queda "verificación
   manual" (P4). Sin fuente: el dato sale del guion o lo aporta el analista.
3. **API del tool.** Esqueleto completo de `build_deck.py` (carga por
   `sys.path.insert(0, "<plugin>/tools")`, `ROOT`, `AS_OF` fijo, `Deck(root,
   model, ticker, as_of, facts, design)`). Lectura: `series(hoja, etiqueta,
   section=)`, `cell(hoja, etiqueta, column=, section=)`, `range(hoja, "C6:G10")`,
   `fact(clave)`; etiquetas duplicadas se desambiguan con `section=` (prefijo de
   la banda, sin distinguir mayúsculas). Índice por periodo (`rev[2027]`,
   `rev["3Q2026"]`). Derivados: aritmética entre `Num`, `growth()`, `over()`,
   `cagr(a, b)` (solo años), `apply(fn, etiqueta)`. Formato admitido (signo,
   miles, decimales, tipo; sin relleno). Láminas: `cover(title, subtitle,
   figures)`, `section(title, question)`, `slide(title, lede, visual, note,
   source, takeaway)`, `raw(content, source, takeaway)`, `legal(text)`.
   Gráficas con sus firmas reales: `bars`, `lines`, `football`, `waterfall`,
   `heatmap`, `table`. `build()` y la ruta de salida. Prohibidos en contenido:
   `<style>`, `<script>`, `<iframe>`, `<object>`, `<embed>`.
4. **Regla de bloqueo y `lit()`.** Tabla pasa / no pasa (la de D2). Recetas de
   prosa: un número que viene del modelo o de un hecho va siempre por f-string
   de un `Num`; un año del eje se escribe solo; un conteo editorial ("cuatro
   pasos") va en letra o con `lit("4", why=...)`; nombres propios con dígitos
   separados ("IFRS 16", "S&P 500", "top 3") van con `lit()` o se reescriben.
   `lit()` no es salida para una cifra del negocio: cada uso sale listado en el
   gate 2 (P3). Constantes dentro de un derivado (`* 100`, `/ 1e6`) también
   quedan en el reporte.
5. **Construir, medir, ajustar.** `measure` no ajusta nada: reporta. Lo normal son
   2 o 3 rondas. P6 (vacía): subir `h=` del visual; P5 (encimada): bajarlo o
   acortar la Nota. El SVG escala al ancho (`1280 - 2 * deck_margin` entre `w`):
   con `w=1100` y margen 56, 10 px de `h` son ~10.6 px en pantalla.
6. **Checks P1 a P10.** Tabla con ID, enunciado, cómo se detecta, comando
   (`audit` o `measure`) y si **bloquea** o **avisa**. P1, P2, P5, P6, P7, P8
   bloquean; P3, P4, P9, P10 avisan. Códigos de salida 0 / 1 / 2 / 3; sin
   navegador `[pendiente]`, nunca verde. Los IDs no se reciclan.
7. **Verificación del gate 2.** Comandos exactos:
   `python tools/deck_builder.py audit <deck> <modelo> --facts <hechos> --design brand/DESIGN.md --report research/notes/<ticker>-deck-verify_<fecha>.md`
   y `python tools/deck_builder.py measure <deck> --shots research/notes/shots_<fecha>`.
   Evidencia por frase: la tabla "Afirmaciones" del reporte se llena fila por
   fila (celda del modelo, o documento con localizador); frase sin evidencia se
   elimina del script y se regenera. Con subagentes disponibles, lo hace uno sin
   el contexto de redacción. **Revisión de capturas una por una**, con la lista
   de lo que la medición no ve (etiquetas pegadas aunque no encimadas, gráfica
   ilegible por escala, tabla densa, serie que no se distingue, lámina que no
   dice su hallazgo). Pasada de texto con los errores que sí aparecieron: una
   diferencia entre dos tasas se dice en puntos porcentuales
   (`(a - b) * 100` con formato `.1f`, nunca `:.1%`); sin contrastes "no es X,
   sino Y"; sin frases de efecto ni metadiscurso; "récord", "primero", "único"
   y fechas de eventos exigen evidencia aunque no lleven número.
8. **Límites conocidos.** Los del spec que afectan a quien escribe: el bloqueo
   solo ve números; el sentido alrededor de una cifra no se revisa; atributos y
   `<ol>` no se escanean; graficar hechos exige armar la `Series` a mano.

- [ ] **Step 2: `SKILL.md`** (~110 líneas, estructura de `skills/model-standards/SKILL.md`)

Frontmatter: `name: research-note`; `description` con disparadores en lenguaje
de analista ("arma el deck", "presenta la cobertura", "prepara la presentación
de la tesis", "actualiza el deck", "¿de dónde salió esta cifra del deck?",
"verifica el deck").

Secciones: apertura (dueña única de `research/notes/` y de la etapa 7; dos
formatos, hoy el deck, la nota .docx pendiente en el issue 55; "en un deck un
número no es texto") · **Cuándo corre** (`/present-coverage` pasos 1 a 4;
`/update-quarter` post-run, solo `audit`; standalone: auditar o regenerar) ·
**Inputs** (tabla: modelo recalculado, `driver-map.md`, industry report vigente,
`journal/*`, `issuer-profile.yaml`, `brand/DESIGN.md`; todos de solo lectura) ·
**Artefactos** (tabla de D1 más la copia fechada del script, decisión 2) ·
**Procedimiento** (1 guion, 2 hechos, 3 script y build, 4 verificación; cada
paso remite a su sección de la referencia) · **Reglas duras** (sin rating ni
recomendación salvo dictado del analista; aviso legal genérico y marca solo con
`brand/DESIGN.md` del analista; ninguna cifra tecleada; `[pendiente]` nunca es
verde; decisión 3) · **Entrevista de cierre** (gate 1: ¿el guion cuenta la
tesis o recita el modelo?; gate 2: ¿qué frase no firmaría el analista?) ·
**Referencias**.

- [ ] **Step 3: Rutas mínimas en los contratos**

`AGENTS.md`, fila nueva al final de la tabla de enrutamiento:

```markdown
| Guion del deck, hechos citados, escribir o verificar el deck de una cobertura | `skills/research-note/SKILL.md` |
```

`README.md`: encabezado `## Las 9 skills — y cuándo usarlas sueltas` y fila:

```markdown
| `research-note` | El deck de la cobertura en `research/notes/`: guion, hechos citados, script y verificación | "arma el deck", "¿de dónde salió esta cifra del deck?" |
```

`CONTRIBUTING.md` 4.6, primer párrafo:

```markdown
Alto costo, gate alto. El repo tiene 9 skills por diseño (el tope original era 6;
cada enmienda está documentada en `docs/architecture.md`, con su porqué). Abre
issue ANTES de
```

- [ ] **Step 4: Suite en verde**

Run: `python -m unittest discover -s tests`
Expected: 431 pruebas, OK.

- [ ] **Step 5: Commit**

```bash
git add skills/research-note tests/test_repo_contracts.py AGENTS.md README.md CONTRIBUTING.md
git commit -m "feat: skill research-note, formato deck y checks P documentados"
```

### Task 3: Comando `/present-coverage`

**Files:**
- Create: `commands/present-coverage.md`
- Modify: `AGENTS.md`, `README.md`

- [ ] **Step 1: `commands/present-coverage.md`** (formato de `commands/update-industry.md`)

Uso: `/present-coverage <TICKER> [contexto]`.

Preflight (cualquier fallo DETIENE):

1. Árbol de cobertura y modelo vigente en `model/`.
2. `python tools/model_check.py <modelo> --driver-map driver-map.md --recalc`
   sin `[x]` ni checks vivos `[pendiente]`.
3. El modelo EN DISCO trae valores calculados (el recálculo del punto 2 corre
   sobre una copia). Si no: con el sí del analista,
   `python tools/recalc_validate.py <modelo> --in-place`.
4. Industry report vigente. Vencido (D10): `[aviso]`, no bloquea.
5. Navegador para `measure`: si falta, `[aviso]`; el deck saldrá `[pendiente]`.

La valuación no es preflight.

Pipeline: la tabla de D4 (pasos 1 a 4 con `research-note`, gate 1 tras el guion,
gate 2 tras la verificación), más una fila previa al paso 3 con la copia fechada
de `build_deck.py` si ya existe. Gate 2 exige revisar las capturas una por una.

Post-run: los seis artefactos fechados en `research/notes/`; dos entradas en
`journal/thesis-journal.md`; `audit` y `measure` en verde o el deck se reporta
como no entregable; `[pendiente]` de `measure` se reporta como tal.

Reporte final: láminas, cifras auditadas, usos de `lit()`, hechos con
verificación manual, frases eliminadas por falta de evidencia, estado de P1 a P10.

- [ ] **Step 2: Rutas**

`AGENTS.md`, fila tras `/update-macro`:

```markdown
| Presentar la cobertura: deck de la tesis desde el modelo ("arma el deck", "presenta la cobertura") | `commands/present-coverage.md` |
```

`README.md`: "Seis comandos orquestan todo."; línea nueva en el bloque de
comandos (`/present-coverage AAPL                # armar el deck de la tesis`);
párrafo tras el de `/update-industry` que explica guion, gate, cifras atadas al
modelo y verificación.

- [ ] **Step 3: Suite en verde y commit**

Run: `python -m unittest discover -s tests`
Expected: 431 pruebas, OK.

```bash
git add commands/present-coverage.md AGENTS.md README.md
git commit -m "feat: comando /present-coverage con dos gates del analista"
```

### Task 4: Contratos

**Files:**
- Modify: `docs/architecture.md`, `templates/coverage-tree.md`,
  `commands/update-quarter.md`, `CONTRIBUTING.md`, `AGENTS.md`, `README.md`

- [ ] **Step 1: `docs/architecture.md`**

Nodo del diagrama: `S6 --> S7[7 · Nota<br/>research-note · deck y nota]`.

Fila en la tabla de dueñas:

```markdown
| `research/notes/*`: guion, hechos citados, `build_deck.py`, deck HTML, verificación y capturas | research-note |
```

Fila en decisiones:

```markdown
| research-note como novena skill, una sola para deck y nota | La etapa 7 no tenía dueña y `research/notes/` nadie lo escribía. Deck y nota comparten insumos, registro de hechos y verificación de cifras: dos skills duplicarían el gate y partirían la propiedad de la carpeta. Tool determinista más script libre por cobertura, igual que `tools/xlsx_builder.py` con `build_model.py`: el tool fija lo que no depende de la emisora |
```

Estado: `- v0.9.0: 9 skills y 6 comandos; la etapa 7 tiene dueña (formato deck; la nota .docx sigue pendiente).`

- [ ] **Step 2: `templates/coverage-tree.md`**

Bloque `notes/` del árbol, en lugar de la línea `<ticker>-note_YYYY-MM-DD.md`:

```
    │   ├── notes/                # etapa 7 (research-note); todo fechado
    │   │   ├── <ticker>-deck-outline_YYYY-MM-DD.md  # guion aprobado (gate 1)
    │   │   ├── <ticker>-facts_YYYY-MM-DD.json       # hechos externos citados
    │   │   ├── build_deck.py                        # script vigente del deck
    │   │   ├── build_deck_YYYY-MM-DD.py             # copia del script anterior
    │   │   ├── <ticker>-deck_YYYY-MM-DD.html        # el deck
    │   │   ├── <ticker>-deck-verify_YYYY-MM-DD.md   # verificación (gate 2)
    │   │   └── shots_YYYY-MM-DD/                    # una captura por lámina
```

Sección `brand/DESIGN.md`: subsección "Slots del deck (opcionales)" con el
bloque `deck_*` de D3, la nota de que `load_brand` del xlsx los ignora y
`load_design` de `tools/deck_builder.py` ignora los `brand_*`, y que un slot
`deck_*` desconocido o mal formado es error.

- [ ] **Step 3: `commands/update-quarter.md`**, viñeta nueva en Post-run:

```markdown
- Si hay deck vigente en `research/notes/`: `python tools/deck_builder.py audit
  <deck> <modelo nuevo> --facts <hechos>`. P1 en `[x]` = deck desactualizado: se
  reporta y se recomienda `/present-coverage` (el guion anterior sirve de base).
  Este comando no regenera el deck.
```

- [ ] **Step 4: `CONTRIBUTING.md`**

Fila en la tabla de familias de 4.2:

```markdown
| **P** (presentación) | `tools/deck_builder.py` (`audit`, `measure`); la lista vive en `skills/research-note/references/deck-format.md` | `demo` en verde; una prueba que falle por cada check |
```

Sección 2, línea de tools: `lo determinista: descarga de datos, construcción/audit del xlsx y del deck.`

- [ ] **Step 5: `AGENTS.md` y `README.md`**

`AGENTS.md`, viñeta en Restricciones de entorno: el deck solo vía
`tools/deck_builder.py`; ninguna cifra tecleada; `audit` y `measure`; sin
navegador `[pendiente]`, nunca verde.

`README.md`: comentario de `notes/` en el árbol ("el deck de la cobertura: guion,
hechos citados y verificación"), `brand/` ("para el Excel y el deck") y la viñeta
de `brand/DESIGN.md` en Extras.

- [ ] **Step 6: Suite, grep de referencias colgadas y commit**

Run: `python -m unittest discover -s tests`
Expected: 431 pruebas, OK.

Run: `grep -rn "8 skills\|Cinco comandos\|note_YYYY" --include=*.md . | grep -v docs/superpowers`
Expected: solo la línea histórica de v0.6.0 en `docs/architecture.md`.

```bash
git add docs/architecture.md templates/coverage-tree.md commands/update-quarter.md CONTRIBUTING.md AGENTS.md README.md
git commit -m "docs: contratos de la etapa 7, research/notes y audit del deck"
```

### Task 5: Dogfood de punta a punta

Fuera del repo: una copia de una cobertura real en `scratch/` (ignorado por git).
Nada de esa copia entra al commit.

- [ ] **Step 1:** Copiar la cobertura de TSM a `scratch/dogfood-pr2/TSM/`, sin
  `research/notes/` (trae el deck hecho a mano y marca institucional). TSM
  ejercita lo que más falló: modelo en disco sin valores calculados, etiquetas
  duplicadas y citas contra un 20-F en iXBRL. Si TSM no recalcula en esta
  máquina, se usa AAPL y se anota.
- [ ] **Step 2:** Un agente sin contexto recibe una sola instrucción: seguir
  `commands/present-coverage.md` para esa copia, registrar los gates como
  simulados y anotar cada punto donde la skill no le dijo qué hacer, le dijo
  algo que el tool contradice, o la regla de bloqueo lo detuvo.
- [ ] **Step 3:** Revisar su salida sin fiarse del reporte: correr `audit` y
  `measure` de nuevo, abrir las capturas una por una, leer el reporte de
  verificación.
- [ ] **Step 4:** Cada fricción se vuelve un cambio en la skill, la referencia o
  el comando (commit `fix:` o `docs:`). Un defecto del tool: prueba que falle,
  arreglo, commit `fix:` aparte.
- [ ] **Step 5:** Anotar para el PR: láminas, cifras auditadas, usos de `lit()`,
  rondas de medición, bloqueos de la regla sobre prosa legítima, lo que salió mal.

### Task 6: Spec, bump y PR

- [ ] **Step 1:** Spec: en D5 `v0.8.0` pasa a `v0.9.0`; línea de estado del PR 2
  con el resultado del dogfood. Commit con este plan:
  `docs: plan del PR 2 y estado del spec de research-note`.
- [ ] **Step 2:** Bump en commit propio: `"version": "0.9.0"` en los dos
  manifests y `v0.9.0` en `README.md`.
  `chore: bump a v0.9.0 - skill research-note y /present-coverage`
- [ ] **Step 3:** Verificación completa (CONTRIBUTING 6): sintaxis de tools,
  `python tools/xlsx_builder.py demo ./scratch/demo.xlsx`,
  `python tools/deck_builder.py demo ./scratch`,
  `python -m unittest discover -s tests`.
- [ ] **Step 4:** Rebase sobre `origin/main` si avanzó; push; PR contra `main`
  con las cuatro preguntas de CONTRIBUTING 7, el resultado del dogfood con lo
  que salió mal, y `Closes #54`. Borrador si la enmienda de 9 skills no tiene
  visto bueno.
- [ ] **Step 5:** Solo con el sí del usuario: comentario en los issues 54 y 55
  con la decisión de una sola skill.

## Desviaciones al ejecutar (2026-11-26)

Las tareas 1 a 4 se ejecutaron como están escritas (`2b393f2`, `c0bc0fd`,
`4405014`). El dogfood (tarea 5) cambió tres cosas que este plan no preveía.

| Commit | Qué cambió | Origen |
|---|---|---|
| `cb99672` | `tools/deck_builder.py`: el número de acto usa la fuente de cuerpo. Con la de encabezado por defecto (Georgia, numerales de estilo antiguo) "01" se leía "O1". Prueba primero: `test_act_number_uses_the_body_font` | Revisión de capturas, con la medición en verde |
| `5adc398` | Skill, referencia y comando: volcado de etiquetas del modelo; comprobación de valores en disco (`model_check` sin `--recalc`); P6 en tabla, sensibilidad y football; P8 por nombres de serie encimados; componer visuales; `raw` sin título; etiquetas del registro de hechos y guidance sin documento; reporte que no conserva la evidencia ni incluye la medición; P9 y P10 `[n/a]` sin `brand/DESIGN.md`; lista ampliada de afirmaciones sin número; cambios posteriores al gate 1 | Bitácora de 23 fricciones del agente |
| `5adc398` | `templates/debate-protocol.md`: dos filas con el objetivo de los gates de `research-note`. Estaba en "Fuera de este plan"; el agente las buscó ahí | Fricción F09 |

Antes de abrir el PR la rama se rebasó sobre `origin/main` (`121b641`: checks K,
macro-view por país y `xlsx_builder` partido en módulos). Tres conflictos, todos
por cercanía y resueltos conservando ambos lados. Los conteos de pruebas de las
tareas (431) son anteriores al rebase; con `main` al día la suite son 513.

Resultado del dogfood, sobre una copia de TSMC sin su deck anterior: 12 láminas,
104 cifras auditadas, 24 hechos con cita cotejada (22 del 20-F, 2 del industry
report), cero `lit()`, el script construyó a la primera, tres rondas de
medición, `audit` y `measure` en verde (repetidos por quien revisó). Gates
simulados, sin analista.

Lo que salió mal:

- **El preflight detuvo el comando.** `model_check --recalc` sobre el modelo
  real de TSMC da 7 `[x]` (S4, S5, S6, S10, C9, C11, F19): el modelo se
  construyó antes de varios de esos checks. Se siguió como desviación
  declarada. El comando hace lo que debe; el modelo de esa cobertura necesita
  `/model-check`.
- **La medición empujó a rellenar.** Tres de seis láminas necesitaron contenido
  que el guion no tenía (8 hechos más para una tabla, un puente nuevo) para
  pasar P6. La skill ahora pide planear la densidad en el guion y declarar el
  cambio en el gate 2.
- **La evidencia por frase falla si la hace quien redactó.** La pasada propia
  cambió 3 frases; un revisor sin contexto encontró 13 imprecisas, entre ellas
  un rótulo que renombraba una fila del modelo y un "cliente mayor" que no es
  el mismo cada año. Tres estaban en celdas de tabla, que `claims` no lista.
- **Quedó en el deck una cifra de un tercero que no cuadra con su propia
  fuente** (participación de mercado), con reserva escrita. La skill ahora dice
  que se resuelve o se retira.
- P9 y P10 no se evaluaron (sin `brand/DESIGN.md`). `raw`, `lit()` y el logo no
  se ejercitaron.

Defectos visibles con la medición en verde que NO se corrigieron (tool, para
una iteración posterior): columna de texto de las tablas alineada a la
derecha; eje sin unidad; puente chato a `h` bajo; Nota a pocos px del cierre;
aviso legal en una lámina casi vacía.

## Fuera de este plan

- Nota .docx (issue 55), escenarios, consenso, exportar a PPTX.
- Mejoras del tool que el dogfood pidió: subcomando para listar etiquetas del
  modelo, validar el registro de hechos sin construir, `measure --report`,
  conservar la evidencia entre corridas de `audit --report`, incluir celdas de
  tabla y rótulos en `claims`, `football` que reparta sus filas, `audit
  --design` sobre la paleta neutra.
- Regenerar las 9 láminas de diagramas a medida del deck de TSMC (criterio 1 del
  spec, dogfood fuera del repo).
