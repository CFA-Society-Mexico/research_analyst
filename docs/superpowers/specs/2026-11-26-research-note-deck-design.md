# Diseño — skill `research-note`, formato 1: deck de cobertura por código

Fecha: 2026-11-26 · Aprobado por Alan (brainstorming P1 = una skill con dos
formatos · P2 = tool base + script libre por cobertura · P3 = cifra sin origen
bloquea; secciones 1 a 5 aprobadas una por una).
Origen: issues #54 (deck) y #55 (nota .docx). Este spec cubre el deck y el
esqueleto de la skill. La nota .docx tiene spec propio, después.

Actualizado el mismo día tras implementar el tool (PR 1), una revisión
independiente y el dogfood con dos coberturas: D2 y D3 describen el tool como
quedó. D1 (skill), D4 (comando) y los contratos de D5 siguen siendo diseño:
van en el PR 2.

## El problema

La etapa 7 del flujo ("Nota: tesis desde journal + disclosures",
`docs/architecture.md`) no tiene skill dueña. `research/notes/` existe en el
árbol y nadie lo escribe. Al cerrar `/init-coverage` el plugin deja modelo,
industry report y journal, y ninguna pieza para presentar la tesis.

En la cobertura de TSMC se construyó un deck de 31 láminas a mano. Revisión de
esa implementación (`research/notes/build_deck.py`, 1,087 líneas, fuera del
repo):

- Las cifras de valuación están tecleadas en el script: WACC, extremos del
  football field, matriz de sensibilidad, precio, acciones, EV. El #54 pide que
  ninguna cifra se teclee; la referencia no lo cumple.
- La emisora está cableada: año de corte, diagramas de wafer y de cadena de
  valor, textos de cada lámina en un diccionario fijo.
- Usa el sistema de diseño de una institución. No puede entrar a un repo
  público.
- El extractor que produjo `series.json` no se conservó, y la medición de
  espacio se hizo a mano en la sesión: no hay código que la repita.
- La verificación final atrapó seis afirmaciones falsas. Todas salieron de
  narrativa escrita sin atadura a una fuente.

Lo rescatable: las primitivas SVG, la anatomía de lámina y el CSS.

## Decisiones de brainstorming

| # | Decisión | Elección |
|---|---|---|
| P1 | Cuántas skills | **Una**, `research-note`, dueña de la etapa 7 y de todo `research/notes/`. Dos formatos: deck (este spec) y nota .docx (spec posterior). Los issues ya decían que comparten insumos y verificación |
| P2 | Qué entrega el plugin | **Tool determinista + script libre por cobertura.** Mismo patrón que `tools/xlsx_builder.py` + `model/build_model.py`. El tool fija lo que no depende de la emisora; guion, narrativa, número de láminas y diagramas a medida son de cada cobertura |
| P3 | Cifra sin origen | **Bloquea.** Un número en texto visible que no venga del modelo, de un hecho citado o de un derivado de ambos impide generar el deck |

Descartadas:

- Dos skills (como proponían los issues): 10 skills contra un repo de 8 por
  diseño, y la verificación de cifras duplicada.
- Generador declarativo (guion en YAML, sin script): las láminas a medida no
  caben sin válvula de escape y el esquema se diseñaría con una sola emisora
  de muestra.
- Solo instrucciones, y que el modelo escriba todo el script cada vez: es lo
  que produjo la referencia de TSMC.
- Verificar comparando los números del HTML contra "cualquier celda del
  modelo": un 38% redondeado coincide con decenas de celdas y los derivados
  (crecimientos, CAGR) no existen en el modelo.
- `series.json` intermedio: una copia más que se desincroniza del xlsx.

## D1 — Dueña, artefactos y nombres

| Pieza | Diseño |
|---|---|
| Skill | `skills/research-note/SKILL.md` + `references/deck-format.md`. Novena skill |
| Tool | `tools/deck_builder.py` (biblioteca + CLI) |
| Comando | `commands/present-coverage.md`. Sexto comando |
| Lee, nunca escribe | modelo recalculado, `driver-map.md`, industry report vigente, `journal/*`, `issuer-profile.yaml`, `brand/DESIGN.md` |

Artefactos en `research/notes/` (todos fechados; vigente = fecha más reciente;
nada se borra):

| Archivo | Qué es |
|---|---|
| `<ticker>-deck-outline_YYYY-MM-DD.md` | Guion aprobado en el gate 1 |
| `<ticker>-facts_YYYY-MM-DD.json` | Hechos externos citados. Lo compartirá la nota .docx |
| `build_deck.py` | Script de la cobertura, escrito sobre el tool |
| `<ticker>-deck_YYYY-MM-DD.html` | El deck |
| `<ticker>-deck-verify_YYYY-MM-DD.md` | Reporte de verificación del gate 2 |
| `shots_YYYY-MM-DD/` | Capturas por lámina |

Los nombres van en inglés (AGENTS.md), no en español como en el #54. El
`<ticker>-note_YYYY-MM-DD.md` que hoy declara `templates/coverage-tree.md` se
reemplaza por estos archivos y, en el spec de la nota, por
`<ticker>-note_<tipo>_YYYY-MM-DD.docx`.

Comando propio y no un paso de `/init-coverage`: el deck es opcional y se
regenera cada trimestre.

## D2 — Contrato de `tools/deck_builder.py`

Dependencias: stdlib + `openpyxl`. Sin matplotlib, sin Node, sin Playwright.

### Lectura del modelo

```python
deck = Deck(root=ROOT, ticker="TSM", as_of="2026-11-26",
            model="model/tsm-model_2026-11-25_v3.xlsx",
            facts="research/notes/tsm-facts_2026-11-26.json",
            design="brand/DESIGN.md")     # rutas relativas a la cobertura (root)
rev  = deck.series("Annual", "Ingresos netos")       # fila por etiqueta
wacc = deck.cell("DCF", "WACC")                      # celda por etiqueta
grid = deck.range("Sensitivity", "C6:G10")           # rango explícito
```

- Lee el xlsx con valores calculados. Si una fórmula pedida no tiene valor en
  caché, el tool no corre y manda a `python tools/recalc_validate.py`.
- `series(hoja, etiqueta, section=None)`: fila cuya columna B coincide exacto
  con la etiqueta. Cero coincidencias o más de una: error. Con duplicados se
  desambigua con `section`, que se compara por prefijo y sin distinguir
  mayúsculas contra la banda de sección (`section="Assumptions"` encuentra
  "Assumptions (única zona de input del libro)"). La búsqueda de banda es
  propia del tool: no depende de un helper privado de `model_check`.
- `Series.apply(fn, etiqueta)` deriva una serie punto a punto (índice base
  100, cambio de escala); `growth()`, `over(otra)` y `cagr(a, b)` cubren los
  casos comunes.
- El eje de periodos y la marca observado/estimado salen del encabezado del
  modelo, con las reglas de `_period_axis` de `tools/model_check.py`. El script
  indexa por periodo: `rev[2027]`, `rev["3Q2026"]`. Ningún año de corte se
  escribe a mano.
- `cell(hoja, etiqueta, column=None)`: la única celda numérica de esa fila. Si
  hay varias, error que pide `column`.
- `range(hoja, "C6:G10")`: para cuadrículas (sensibilidad). El origen queda
  registrado igual.
- `as_of` lo fija el script, como `BUILD_DATE` en `build_model.py`. El tool no
  lee el reloj (CONTRIBUTING §4.3).

### Número con procedencia

`series`, `cell`, `range` y `fact` devuelven objetos `Num`: valor más origen
(hoja y celda, o clave de hecho). La aritmética entre `Num`, y entre `Num` y
constantes, devuelve otro `Num` que conserva la expresión completa:

```python
g27 = rev[2027] / rev[2026] - 1
deck.slide(title=f"2027: vender {g27:+.0%} más y generar menos caja", ...)
```

Al formatearse, un `Num` emite un marcador interno que el render convierte en
`<span data-src data-fmt>` (HTML) o `<tspan data-src data-fmt>` (SVG).
`data-src` guarda la expresión (`Annual!H14/Annual!G14-1`); `data-fmt`, el
formato.

Esto es código determinista calculando, no el modelo de lenguaje: no roza la
regla central.

El formato admite solo signo, separador de miles, decimales y tipo (`+,.1f`,
`.0%`). Relleno, ancho y alineación se rechazan: un relleno como `9>6`
inyecta dígitos que no vienen del modelo. Un resultado que no es un número
real finito (división entre cero, raíz de un negativo) es error.

`build()` aplica a su propia salida la misma comprobación que `audit`: cada
cifra debe tener al menos una referencia (celda u hecho) y su texto debe ser
el de su origen. Un origen inventado o un texto retocado después de formatear
no llegan al archivo.

### Hechos externos

```json
{"as_of": "2026-11-26",
 "facts": [{"key": "foundry_share", "value": 72.5, "unit": "%",
            "tag": "observado",
            "doc": "research/industry/industry-report_2026-11-20.md",
            "locator": "sec. 2, tabla 1", "quote": "72.5%"}]}
```

- `key`: letras, dígitos y `_`. Sin `.` ni `-`, que en una expresión de origen
  serían ambiguos con una resta.
- `tag`: `observado` o `guidance`.
- `doc` debe existir dentro de la cobertura. Si no existe, o si la ruta sale
  de la cobertura (`..`, ruta absoluta): error. Lo mismo aplica al logo.
- Fuente de texto (`.md`, `.txt`, `.htm`, `.html`, `.csv`, `.json`): `quote`
  debe aparecer en el **texto visible** del documento, con espacios y saltos
  de línea normalizados. En HTML se quitan las etiquetas antes de cotejar: un
  filing en iXBRL parte cada cifra con etiquetas y una cita correcta no
  aparece en el HTML crudo. Si no aparece: error.
- Otros formatos (PDF): el hecho se acepta y sale en el reporte como
  "verificación manual".

### Regla de bloqueo

Tras renderizar, el tool recorre el texto visible del HTML y de todo SVG
(incluidos los que el script escriba a mano):

| Caso | Resultado |
|---|---|
| Número que no proviene de un `Num`, con o sin sufijo de unidad o símbolo de moneda (`38%`, `11x`, `180pbs`, `US$5`, `USD5`, `2,150`) | **Error: no se genera el deck** |
| El mismo número rodeado de símbolos (`~38%`, `>38%`, `n=30`, `38%*`), en un rango (`10%-12%`), con dígitos no ASCII o caracteres invisibles | **Error** |
| Año del eje del modelo escrito solo (`2027`, `(2027)`) | Pasa. Con símbolo (`$2027`) no |
| Dígitos pegados a letras que no son unidad ni moneda (`N2`, `A14`, `3nm`, `IFRS16`, `1Q2026E`) y dígitos-guion-letras (`20-F`, `10-K`, `COVID-19`) | Pasa |
| `lit("4", why="pasos del método")` | Pasa y queda listado en el reporte |
| Lo que genera el tool (`class="auto"`: marcas de eje, número de lámina, fecha de portada) | Pasa |
| La línea de fuente (`class="source"`) | No se escanea: una fuente cita años y páginas |

### CLI

| Comando | Qué hace |
|---|---|
| `deck_builder.py demo <carpeta>` | Deck neutro desde un libro sintético; debe quedar en verde |
| `deck_builder.py audit <deck.html> <modelo> [--facts f.json] [--design brand/DESIGN.md] [--root carpeta] [--report r.md]` | Relee el modelo, re-evalúa cada `data-src` con su `data-fmt` y compara contra el texto mostrado. Repite el escaneo de bloqueo. Detecta deck viejo contra modelo nuevo. Con `--design` agrega los avisos de contraste |
| `deck_builder.py measure <deck.html> [--shots carpeta]` | Medición de espacio y capturas (D3) |
| `deck_builder.py claims <deck.html> <salida.md>` | Tabla de todas las frases visibles, por lámina. Va a archivo: las frases llevan acentos y la consola es ASCII-only |

Los checks se identifican `P1` a `P10`: P1 cifras iguales a su origen, P2 sin
números sueltos, P3 literales declarados (aviso), P4 hechos con verificación
manual (aviso), P5 sin contenido encimado, P6 sin lámina vacía, P7 sin
desborde horizontal, P8 etiquetas dentro de su gráfica, P9 contraste del
texto (aviso), P10 series distinguibles con daltonismo (aviso).

Con `--report`, `audit` escribe las partes deterministas del reporte de
verificación: cifras con su origen, expresiones derivadas, usos de `lit()`,
hechos con verificación manual y la tabla de frases con la columna de
evidencia vacía.

Códigos de salida, iguales a `model_check.py`: 0 en verde, 1 con `[x]`, 2 uso
incorrecto, 3 con `[pendiente]`. Consola ASCII-only.

## D3 — Lámina, gráficas y medición de espacio

### Lámina

Canvas de 1280×720. HTML autocontenido: CSS y logo embebidos; una lámina por
página al imprimir a PDF.

| Tipo | Regla |
|---|---|
| `deck.cover`, `deck.section` | Plantillas fijas. Exentas de medición |
| `deck.slide` | Título (el hallazgo, con su número), entradilla, visual, Nota (2 o 3 frases bajo el visual), cierre, fuente. Sin Nota o sin fuente: error |
| `deck.raw` | HTML/SVG propio de la emisora. Sin anatomía obligatoria; pasa por el bloqueo de cifras y por la medición |

El contenido de una lámina no admite `<style>`, `<script>`, `<iframe>`,
`<object>` ni `<embed>` propios: con ellos se puede pintar un número que el
escaneo no ve. Sí admite atributos `style=`.

### Gráficas

Primitivas SVG sin librerías: barras, líneas, crecimiento anual, football
field, cascada, mapa de sensibilidad, tabla. Reciben `Num`.

- El corte entre observado y proyectado sale del eje del modelo y se rotula.
- El script elige qué puntos etiquetar.
- Un eje por gráfica.
- Son opcionales: el script puede no usarlas y dibujar lo suyo.

### Tokens de diseño

Mismo `brand/DESIGN.md` y mismo formato (`slot: valor`) que ya parsea
`load_brand` en `tools/xlsx_builder.py`. Slots nuevos, todos opcionales:

```
deck_background: #FFFFFF
deck_text:       #1A1A1A
deck_muted:      #5B6670      # texto secundario, ejes
deck_negative:   #B03A2E      # barras negativas
deck_series_1:   #1F4E79      # hasta deck_series_6
deck_font_heading: Georgia
deck_font_body:    Arial
deck_margin:     56           # px
deck_gap_max:    90           # px, umbral de lámina vacía
deck_logo:       brand/logo.svg
```

Sin archivo o sin slot: estilo neutro por defecto. El repo no trae ninguna
marca. Al cargar, el tool calcula el contraste de cada color de serie y de
texto contra el fondo: contraste bajo en una serie obliga etiquetas de valor;
la separación para daltonismo sale como `[aviso]` y no bloquea.

### Medición de espacio

El HTML lleva un script de medición embebido. `measure` abre el deck con el
navegador ya instalado (Edge, Chrome o Chromium; detección con el patrón de
`find_engine` en `tools/recalc_validate.py`) en modo `--headless --dump-dom`,
lee el resultado y, con `--shots`, guarda una captura por lámina.

| Check | Criterio |
|---|---|
| Encimado | Distancia entre fin del contenido y línea de cierre > 0 |
| Cortado | Ninguna lámina, tampoco portada, acto o aviso legal, con contenido que su marco corta |
| Vacío | Esa distancia < `deck_gap_max` |
| Desborde horizontal | Ningún bloque de texto más ancho que su caja |
| Texto fuera de gráfica | Ninguna etiqueta SVG fuera del área de su gráfica |
| Etiquetas encimadas | Ningún par de etiquetas de una misma gráfica con sus cajas encimadas |

Las gráficas ayudan a pasar esos checks por construcción: menos marcas de eje
cuando la gráfica es baja, etiquetas de los extremos de una línea hacia
adentro, la segunda serie etiquetada por debajo de su línea y el rótulo
"Proyectado" sobre el área de la gráfica.

- Sin navegador: `[pendiente]` y salida 3. Nunca verde.
- Sin auto-ajuste. `measure` reporta por lámina la distancia y la altura
  disponible para el visual; el script se corrige y se vuelve a correr.

Verificado el 2026-11-26 en Windows con Edge: `--dump-dom` ejecuta el script
embebido y devuelve la medición (distancia y etiqueta SVG fuera de área
detectadas); `--screenshot` con `?shot=N` genera una captura de 1280×720 por
lámina, sin compensar nada. La CI corre la suite en Linux y Windows con Python
3.10 y 3.13; las pruebas con navegador real se activan con
`RA_BROWSER_TESTS=1` y solo se han corrido en Windows con Edge. macOS: sin
verificar.

### Resultado del dogfood (2026-11-26)

Dos coberturas reales, fuera del repo, sobre copias de trabajo:

| | TSMC | Apple |
|---|---|---|
| Modelo | Anual, FY2016–FY2035, TWD | Trimestral + anual, FY2016–FY2031, USD |
| Láminas | 22 (de las 31 del deck hecho a mano) | 15 (guion nuevo, sin gate de analista) |
| Cifras auditadas contra el modelo | 220 | 174 |
| Cifras tecleadas / `lit()` | 0 / 0 | 0 / 0 |
| Hechos externos con cita cotejada | 18 | 9 |

Lo que el tool obligó a corregir del deck de TSMC hecho a mano: un redondeo
tecleado (1.86% donde el modelo da 1.85%), un rango de valuación que no está
en el modelo y dos datos de operación sin cita en el 20-F.

Lo que el dogfood obligó a corregir del tool: el cotejo de citas contra HTML
crudo, la búsqueda de sección por título exacto, la falta de `Series.apply`,
y varias colisiones de etiquetas en las gráficas que la medición original no
veía. La lección que queda para la skill: **la medición no sustituye mirar
las capturas**; el gate 2 las exige.

## D4 — `/present-coverage <TICKER>`

### Preflight (cualquier fallo detiene)

1. Existen el árbol de cobertura y un modelo vigente en `model/`.
2. `python tools/model_check.py <modelo> --recalc` termina sin `[x]` y sin
   checks vivos `[pendiente]`.
3. Existe industry report vigente. Vencido (D10): `[aviso]`, no bloquea.
4. Navegador para `measure`: si falta, `[aviso]`; el deck saldrá `[pendiente]`.

La valuación no es preflight. Un guion sin acto de valuación es válido; si lo
tiene y la celda no existe, el tool falla al leerla.

### Pipeline

| Paso | Skill | Qué hace |
|---|---|---|
| 1 | `research-note` | Guion: lista de láminas, cada una con su hallazgo, tipo de visual y las series y hechos que necesita. Actos adaptados a la emisora. Escribe el outline |
| Gate 1 | — | El analista aprueba el guion. Debate según `templates/debate-protocol.md`; entrada en thesis-journal |
| 2 | `research-note` | Registra cada dato externo del guion en el archivo de hechos. Sin fuente: sale del guion o el analista la aporta |
| 3 | `research-note` | Escribe `build_deck.py` y lo corre. El bloqueo de cifras aplica aquí |
| 4 | `research-note` | Verificación: `audit --report`, `measure --shots`, evidencia por frase y pasada de texto |
| Gate 2 | — | El analista revisa capturas y reporte y aprueba la entrega. Entrada en thesis-journal |

### Verificación del paso 4

| Parte | Quién | Qué |
|---|---|---|
| Cifras | Código | Cada número contra su origen |
| Espacio | Código | Los cuatro checks de D3 |
| Excepciones | Código | Usos de `lit()`, expresiones derivadas, hechos con verificación manual |
| Afirmaciones | Código lista, modelo evalúa | Cada frase del deck recibe evidencia (celda o documento con localizador) o se elimina |
| Texto | Modelo | Sin frases de efecto, sin contrastes "no es X, es Y", sin metadiscurso |

Si el runtime permite subagentes, la parte de afirmaciones la hace uno sin el
contexto de redacción. Es opcional: no todos los runtimes lo tienen.

### Reglas fijas de la skill

- Sin rating ni recomendación, salvo que el analista los dicte.
- Atribución y aviso legal genéricos. Marca institucional solo si el analista
  la pide y aporta su `brand/DESIGN.md`.
- Audiencia con conocimiento promedio de finanzas: cada término técnico se
  explica la primera vez que aparece.
- Trimestre nuevo: `/update-quarter` no regenera el deck. Su post-run corre
  `audit` contra el modelo nuevo; si no coincide, reporta el deck como
  desactualizado y recomienda `/present-coverage`. El outline anterior sirve
  de base.

## D5 — Pruebas, contratos y aceptación

### Pruebas (`tests/test_deck_builder.py`, unittest de stdlib, sin red)

Con un libro sintético de valores. Ningún dato de emisora real en el repo.

| Qué | Caso que debe fallar |
|---|---|
| Lectura | Etiqueta inexistente; etiqueta duplicada sin `section`; fórmula sin valor en caché |
| Procedencia | Un derivado que pierda su expresión tras la aritmética |
| Bloqueo | Número suelto en texto y en SVG propio. Pasan años del eje, `N2`, `20-F` y `lit()` |
| Hechos | Documento inexistente; `quote` ausente en fuente de texto |
| `audit` | Número alterado en el HTML; modelo cambiado después de generar |
| `measure` | Lámina encimada, vacía, con desborde, con etiqueta fuera de área. Sin navegador: `[pendiente]` |
| Determinismo | Misma entrada, mismo HTML byte a byte |

Las pruebas de `measure` con navegador real usan `skipUnless`; la lógica se
prueba con `find_engine` simulado, como en `tests/test_recalc_validate.py`.
`tests/test_cli_ascii.py` ya escanea todo `tools/*.py`.

### Contratos que cambian

| Archivo | Cambio |
|---|---|
| `AGENTS.md`, `README.md` | Rutas nuevas; 9 skills, 6 comandos (lo exige `tests/test_repo_contracts.py`) |
| `CONTRIBUTING.md` §4.6 | "8 skills por diseño" pasa a 9 |
| `docs/architecture.md` | Dueña de la etapa 7; fila en la tabla de dueñas; fila en decisiones (novena skill y por qué una sola para dos formatos) |
| `templates/coverage-tree.md` | Archivos de `research/notes/`; slots `deck_*` en la sección de `brand/DESIGN.md` |
| `commands/update-quarter.md` | Post-run: `audit` del deck vigente |
| `.claude-plugin/plugin.json`, `.codex-plugin/plugin.json` | v0.8.0, en commit propio |

Coberturas existentes: nada se rompe. `research/notes/` ya existe y hoy está
vacío o con archivos que ningún tool lee.

### Criterios de aceptación

1. El deck de TSMC se regenera desde el tool con cero cifras tecleadas y
   `audit` y `measure` en `[ok]`.
2. Se construye el deck de AAPL sin modificar el tool. Si hay que modificarlo,
   el cambio entra antes de fusionar.
3. La suite completa pasa en CI (Linux y Windows).

Los criterios 1 y 2 son dogfood fuera del repo.

Estado al 2026-11-26:

- Criterio 1: cumplido en parte. 22 de las 31 láminas, con `audit` y
  `measure` en `[ok]`. Faltan las de diagramas a medida.
- Criterio 2: el deck de Apple se construyó, auditó y midió sin cambiar la API
  del tool; la revisión de capturas sí obligó a corregir las gráficas de
  líneas y a agregar la detección de etiquetas encimadas. Ese cambio entró
  antes de fusionar, como pedía el criterio.
- Criterio 3: cumplido para el PR 1.

### Entrega

Dos PRs, por la regla de uno por tema: (1) `tools/deck_builder.py` con sus
pruebas; (2) skill, comando, documentación y bump. El #54 cierra con el
segundo. El #55 queda abierto y se actualiza con la decisión P1.

## Invariantes

- El modelo de lenguaje no calcula ni decide: las cifras salen del xlsx o de
  aritmética en código; el guion, la narrativa y cualquier recomendación son
  del analista, con gate.
- Cada artefacto tiene una dueña: `research/notes/` es de `research-note`.
- Nada se borra en la cobertura: cada corrida escribe archivos fechados.
- Sin navegador o sin recálculo, el estado es `[pendiente]`, nunca verde.
- Consola ASCII-only; sin reloj local en artefactos; sin dependencias nuevas.

## Límites conocidos

- El bloqueo solo atrapa números. Una fecha de evento escrita como año del eje
  ("entró a producción en 2025") pasa, y una afirmación sin número ("margen
  récord") también. Quedan para la tabla de afirmaciones y el gate 2.
- La evaluación de afirmaciones la hace el modelo sobre texto que él mismo
  redactó. El código garantiza que la lista está completa, no que el juicio
  sea correcto.
- `lit()` y las constantes dentro de derivados son válvulas: se pueden usar
  mal. Su única defensa es que aparecen listadas en el gate 2.
- El sentido alrededor de una cifra válida no se revisa: `-{num}`,
  `{num} mil` o "ventas de 2027 millones" pasan.
- `class="auto"` y `class="source"` escritos a mano en HTML propio son
  indistinguibles de los que genera el tool. `audit` detecta un deck
  desactualizado o un número retocado, no una falsificación deliberada del
  HTML.
- Atributos (`alt`, `title`) y la numeración de `<ol>` no se escanean.
- La regla es estricta con prosa normal: "IFRS 16", "S&P 500" o "top 3"
  exigen `lit()` o escribirse de otra forma. En los dos dogfoods no bloqueó
  texto legítimo, pero quien los escribió conocía la regla.
- Graficar hechos exige armar la serie a mano, no hay control de densidad en
  las tablas y `Deck` no da acceso por coordenada. Pendiente para una
  iteración posterior.
- El umbral de 90 px salió de un solo deck.
- La medición dice que no hay hueco ni encimado, no que la lámina se vea bien.

## Pendiente de quien mantiene el repo

La enmienda de 8 a 9 skills (CONTRIBUTING §4.6). Conviene registrarla en el
#54 antes de abrir el primer PR.

## Fuera de alcance

Nota .docx (#55), escenarios, contraste contra consenso, gráfica de precio
contra índice, exportar a PPTX, auto-ajuste de tamaño de gráficas.
