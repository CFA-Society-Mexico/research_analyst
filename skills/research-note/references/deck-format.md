# Deck de cobertura — formato, API del tool y verificación

Material de `research-note` para el formato deck. La implementación única es
`tools/deck_builder.py`; si este documento y el código discrepan, manda el
código (y se corrige este documento).

Doctrina: **en un deck un número no es texto.** Sale del modelo recalculado, de
un hecho citado o de aritmética en código entre ambos, y viaja con su origen
hasta el HTML. Un número suelto impide generar el deck.

## 1. Guion (`<ticker>-deck-outline_YYYY-MM-DD.md`)

El orden no cambia: qué es la empresa, qué la mueve, qué tan rentable es, cuánto
vale. Los actos sí: se adaptan a la emisora (una fundidora se explica por nodo y
wafer; un banco de consumo, por cartera y costo de riesgo). Un acto que el
modelo no sostiene se omite: un guion sin valuación es válido.

Una fila por lámina:

| # | Acto | Tipo | Hallazgo (será el título) | Visual | Del modelo | Hechos |
|---|---|---|---|---|---|---|
| 5 | Crecimiento | `slide` | Las ventas crecen menos de lo observado hasta el último año proyectado | barras, observado contra proyectado | `Annual` · Ingresos netos | — |
| 6 | Crecimiento | `slide` | La emisora concentra la mayor parte de su mercado | tabla, siete filas | — | `foundry_share` y seis más |
| 9 | Rentabilidad | `slide` | La depreciación pesa más sobre las ventas cada año proyectado | líneas | derivado: `Annual` · Depreciación entre Ingresos netos | — |
| 11 | Valuación | `slide` | El valor por acción queda arriba del precio que usa el modelo | football + tabla | celda: `Summary` · Valor por acción; rango: `Sensitivity` C9:E11 | — |

Reglas del guion:

- **El hallazgo es una afirmación, no un tema.** "Márgenes" no es un hallazgo;
  "El margen bruto proyectado queda debajo del observado" sí.
- **Ninguna cifra se teclea en el guion.** La cifra se nombra por su origen
  (hoja y etiqueta, o clave de hecho) y la pone el script. El analista aprueba
  el argumento; el número lo trae el modelo.
- Un concepto por lámina. Cada término técnico (ROIC, WACC, valor terminal, la
  unidad que vende la emisora) se explica la primera vez que aparece: la
  audiencia tiene conocimiento promedio de finanzas.
- Todo dato que no esté en el modelo se anota en la columna Hechos. En el paso 2
  consigue fuente o sale del guion.
- **El rótulo dice lo que es la fila del modelo.** Si la celda se llama "DCF -
  múltiplo de salida", la lámina no la presenta como "múltiplo de comparables";
  si la serie es ventas totales entre obleas, no es "precio por oblea". Un
  nombre más claro para la audiencia se explica en la entradilla, no sustituye
  al concepto.
- **Planea la densidad.** Una tabla de cuatro filas, una cuadrícula de
  sensibilidad de tres por tres o un football field de un método no llenan una
  lámina: la medición la marcará vacía (P6). Desde el guion: más filas, o dos
  visuales que sostengan el mismo hallazgo (football + tabla de supuestos). No
  dos hallazgos en una lámina.
- Tipos: `cover` y `section` (plantillas fijas), `slide` (anatomía completa),
  `raw` (diagrama o tabla a medida de la emisora), `legal` (aviso al cierre).

Anatomía de una lámina `slide`:

| Bloque | Regla |
|---|---|
| Título | El hallazgo, con su número |
| Entradilla | Una o dos frases con la consecuencia |
| Visual | Gráfica, tabla o diagrama |
| Nota | 2 o 3 frases bajo el visual: qué pasó, qué supone el modelo, por qué importa |
| Cierre | Una línea concreta. Opcional |
| Fuente | Obligatoria |

Sin Nota o sin fuente, el tool no acepta la lámina.

## 2. Registro de hechos (`<ticker>-facts_YYYY-MM-DD.json`)

Todo dato externo al modelo: participación de mercado, capacidad instalada,
número de clientes, una fecha de evento, guidance de la dirección.

```json
{"as_of": "2026-11-26",
 "facts": [{"key": "foundry_share", "value": 72.5, "unit": "%",
            "tag": "observado",
            "doc": "research/industry/industry-report_2026-11-20.md",
            "locator": "sec. 2, tabla 1", "quote": "72.5%"}]}
```

| Campo | Regla |
|---|---|
| `key` | Letras, dígitos y `_`. Sin `.` ni `-`. Única en el archivo |
| `value` | Número. El script lo formatea; aquí va sin redondear |
| `unit` | Texto libre (`%`, `USD mm`, `obleas`). Informativo: el script escribe la unidad |
| `tag` | `observado` o `guidance`. Un `supuesto` del analista no es un hecho: vive en el modelo |
| `doc` | Ruta relativa a la cobertura; el archivo debe existir. No puede salir de la cobertura |
| `locator` | Página, sección, nota o tabla |
| `quote` | Fragmento literal de la fuente, de al menos 2 caracteres |

- **La cita se copia del texto visible de la fuente**, lo que una persona lee en
  pantalla. Un filing en iXBRL parte cada cifra con etiquetas: el tool quita las
  etiquetas y normaliza espacios antes de cotejar, así que una cita copiada del
  HTML crudo no coincide. Cita corta y exacta (la cifra con su unidad) antes que
  una frase larga. Ese texto visible se obtiene con `facts --text` (abajo).
- Fuentes de texto (`.md`, `.txt`, `.htm`, `.html`, `.csv`, `.json`): si la cita
  no aparece, el tool no carga los hechos.
- Otras fuentes (PDF): el hecho se acepta y sale como "verificación manual"
  (check P4). Se coteja a mano en el gate 2.
- Dato sin fuente: sale del guion, o el analista aporta el documento a la
  cobertura. No se sustituye por memoria ni por una búsqueda sin archivar.
- Opinión o estimación de un tercero: lleva su hecho con ese documento, `tag`
  `observado` (lo observado es que la fuente lo publica), y la frase la
  atribuye ("<fuente> estima"). Nunca se presenta como dato de la emisora.
- El registro solo admite dos etiquetas. Si el industry report marca un dato
  como fuente secundaria sin verificar, o como estimación que el analista no
  adoptó, ese dato queda fuera del deck hasta que el documento primario esté en
  la cobertura.
- Guidance que ya vive dentro del modelo (un año anclado en la guía de la
  dirección): sin el documento en la cobertura no hay hecho `guidance` que
  citar. La lámina lo dice como supuesto del modelo ("el modelo supone"), sin
  atribuir cifra a la dirección.
- **Cita de una tabla:** una celda, no la fila. Al quitar las etiquetas, las
  celdas de una fila quedan separadas por un espacio (`2,192,931 58%`); una
  cita que cruza celdas es frágil.
- **Valida temprano, sin construir un deck.** Antes de escribir láminas:

  ```
  python tools/deck_builder.py facts <hechos.json> --root <cobertura>
  ```

  Lista cada hecho con `cita encontrada`, `manual` (aviso: fuente no textual,
  se coteja a mano) o el motivo por el que no carga; todos en una corrida, no
  solo el primero. Sale en 1 si algún hecho no carga, y en 3 (`[pendiente]`,
  nunca verde) si la lista está vacía. Sin `--root` toma la cobertura del
  registro cuando éste vive en `<cobertura>/research/notes/`; en otra carpeta,
  `--root` es obligatorio.
- **Cita que no aparece:** se copia del texto que el tool coteja, no del HTML.

  ```
  python tools/deck_builder.py facts <hechos.json> --root <cobertura> --text <doc> <salida.txt>
  ```

  `<doc>` es la ruta de la fuente relativa a la cobertura; `<salida.txt>` va a
  una carpeta de trabajo, no a `filings/`. El volcado sale aunque el registro
  no valide. Son líneas de hasta 100 caracteres (una palabra más larga no se
  parte); una cita puede cruzar un salto de línea, porque los espacios se
  normalizan al cotejar. Si `<salida.txt>` ya existe con otro contenido, el
  tool no lo pisa: se elige otro nombre.

## 3. El script (`research/notes/build_deck.py`)

Esqueleto. `PLUGIN` es la carpeta del plugin en este runtime (la raíz del clon,
o la del plugin instalado).

```python
"""Deck de cobertura. Correr: python research/notes/build_deck.py"""
import os
import sys
from pathlib import Path

PLUGIN = Path(os.environ.get("RA_PLUGIN_DIR", r"<ruta del plugin>"))
sys.path.insert(0, str(PLUGIN / "tools"))
from deck_builder import Deck, DeckError, lit  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]   # carpeta de la cobertura
AS_OF = "2026-11-26"                         # fijo: el tool no lee el reloj


def main() -> int:
    deck = Deck(root=ROOT, ticker="TSM", as_of=AS_OF,
                model="model/tsm-model_2026-11-25_v3.xlsx",
                facts=f"research/notes/tsm-facts_{AS_OF}.json",
                design="brand/DESIGN.md")
    rev = deck.series("Annual", "Ingresos netos")
    g = rev.growth()
    deck.cover("Emisora", "Inicio de cobertura",
               figures=[("Ventas 2025", f"{rev[2025]:,.0f}")])
    deck.section("Crecimiento", "¿De dónde sale el crecimiento?")
    deck.slide(
        title=f"Las ventas crecen {g[2027]:.0%} en 2027",
        lede="El ritmo proyectado es menor al observado.",
        visual=deck.bars(rev, show=(2025, 2027), h=330),
        note=f"Entre 2022 y 2025 crecieron {rev.cagr(2022, 2025):.1%} al año.",
        source="Fuente: modelo de la cobertura; reporte anual.",
        takeaway="El crecimiento se modera.")
    deck.legal()
    print(f"[ok] deck -> {deck.build().name}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except DeckError as exc:
        print(f"[x] {exc}")
        sys.exit(1)
```

Las rutas de `model`, `facts` y `design` son relativas a la cobertura. `model`
es el vigente de `model/` (fecha y versión más recientes), con el nombre que
tenga en esa cobertura. `facts` solo se pasa si el archivo existe; `design`
puede pasarse siempre (sin archivo, estilo neutro; los slots están en
`templates/coverage-tree.md`). La salida es
`research/notes/<ticker>-deck_<as_of>.html`.

### Conocer el modelo antes de escribir

El script necesita la etiqueta EXACTA de la columna B, la banda de sección
cuando la etiqueta se repite y la letra de columna cuando la fila tiene varias
celdas. El tool lo vuelca a un archivo (las etiquetas llevan acentos y la
consola es ASCII-only):

```
python tools/deck_builder.py labels <modelo.xlsx> <salida.md>
```

Por hoja da los periodos del eje con las claves que usa el script (`2027`,
`3Q2026`), sus columnas y dónde empieza lo proyectado. Por etiqueta: su fila,
si es banda de sección (negrita de 16 pt o más: lo que acepta `section=`),
cuántas celdas numéricas tiene y entre qué columnas. La columna Nota marca:

- `repetida: usa section=`: la etiqueta aparece en más de una banda.
- `repetida en su sección` o `repetida y fuera de sección`: `section=` no las
  separa; esas filas se leen con `deck.range(hoja, "C10:H10")`.
- `salto de línea`: la etiqueta trae un salto dentro de la celda; en el script
  se escribe con `\n`, como aparece en el archivo.
- fórmulas sin valor calculado: el archivo no está recalculado (preflight de
  `/present-coverage`).

Es solo lectura sobre el modelo. La salida va a una carpeta de trabajo, no a
`model/`; el tool solo reemplaza un archivo que sea un mapa anterior.

Para inspeccionar un `Num`, `print(n)` escribe valor y origen en ASCII
(`240.0 [Annual!H6]`). Es para depurar: `str(n)` y `repr(n)` no atan la cifra,
y dentro de una lámina bloquean como número suelto. Lo que ata la cifra a su
origen es el f-string, con o sin formato (`f"{n:,.0f}"`, `f"{n}"`).

### Leer del modelo

| Llamada | Devuelve | Notas |
|---|---|---|
| `deck.series(hoja, etiqueta, section=None)` | `Series` | Fila cuya columna B es exactamente la etiqueta. Se indexa por periodo: `rev[2027]`, `rev["3Q2026"]` |
| `deck.cell(hoja, etiqueta, column=None, section=None)` | `Num` | La única celda numérica de la fila; si hay varias, pide `column="C"` |
| `deck.range(hoja, "C6:G10")` | lista de listas de `Num` | Cuadrículas (sensibilidad) |
| `deck.fact(clave)` | `Num` | Del registro de hechos |

- **Etiqueta duplicada** (la misma línea en el estado de resultados y en el DCF):
  el tool falla y pide `section=`. Se compara por prefijo de la banda de sección,
  sin distinguir mayúsculas: `section="Assumptions"` encuentra "Assumptions
  (única zona de input del libro)".
- **Fórmula sin valor calculado**: el tool no corre. El modelo se recalcula antes
  (preflight de `/present-coverage`); el script no lo arregla.
- El corte entre observado y proyectado sale del encabezado del modelo. Ningún
  año de corte se escribe en el script.

### Derivar

`Num` con `+ - * / **`, entre sí o con constantes, devuelve otro `Num` que
conserva la expresión completa (`Annual!H14/Annual!G14-1`). Es código
determinista calculando; no es el asistente estimando.

| Atajo | Qué da |
|---|---|
| `serie.growth()` | Crecimiento contra el periodo anterior |
| `serie.over(otra)` | Razón periodo a periodo (un margen) |
| `serie.cagr(2022, 2027)` | Crecimiento compuesto entre dos años del eje |
| `serie.apply(fn, etiqueta)` | Serie derivada punto a punto (índice base, cambio de escala). `fn` devuelve un `Num` |
| `serie.keys`, `serie.is_actual(k)` | Periodos, y si el periodo es observado |

Formato: solo signo, separador de miles, decimales y tipo (`+,.1f`, `.0%`,
`,.0f`). Relleno, ancho y alineación se rechazan.

### Láminas

| Llamada | Uso |
|---|---|
| `deck.cover(title, subtitle="", figures=())` | Portada. `figures`: pares (rótulo, valor ya formateado) |
| `deck.section(title, question="")` | Separador de acto, numerado por el tool |
| `deck.slide(title, lede, visual, note, source, takeaway="")` | Lámina de contenido |
| `deck.raw(content, source, takeaway="")` | HTML o SVG propio: diagrama de la cadena de valor, tabla a medida |
| `deck.legal(text=...)` | Aviso legal; por defecto, uno genérico |

El contenido de una lámina no admite `<style>`, `<script>`, `<iframe>`,
`<object>` ni `<embed>`. Sí admite atributos `style=`. Una lámina `raw` pasa por
el mismo bloqueo de cifras y la misma medición que las demás; no tiene `title`:
su título va dentro del contenido, como `<h2>`.

### Gráficas (devuelven el HTML del visual; son opcionales)

| Llamada | Notas |
|---|---|
| `deck.bars(serie, keys=None, show=(), fmt=",.0f", scale=1.0, tick_fmt=None, w=1100, h=320)` | Observado y proyectado en colores distintos, corte rotulado. `show`: periodos con etiqueta de valor. `scale` divide (1e6 para millones) |
| `deck.lines([(nombre, serie), ...], keys=None, show=(), fmt=".0%", tick_fmt=None, w=1100, h=320)` | Varias series, un eje. La primera fija periodos y corte |
| `deck.football([(método, mín, máx, centro), ...], marker=(rótulo, valor), fmt=",.0f", w=1100, h=None)` | Rangos de valor; `marker` es la línea del precio. `h` reparte las filas en el alto; sin `h`, 60 por fila más 60 |
| `deck.waterfall([(rótulo, valor), ...], total=(rótulo, valor), fmt=",.0f", w=1100, h=320)` | Puente. El total se lee del modelo, no se suma aquí |
| `deck.heatmap(grid, rows, cols, fmt=",.0f", row_fmt=".1%", col_fmt=".1%", corner="", center=None)` | Sensibilidad. `center=(fila, columna)` marca el caso base |
| `deck.table(header, rows)` | Celdas de texto; los números llegan formateados desde un `Num`. Columna sin cifras, a la izquierda; con cifras, a la derecha |

- `keys=` recorta los periodos (una lista del eje). Etiquetar pocos puntos: los
  que sostienen el hallazgo.
- Los rótulos de serie, de método y de paso también se escanean: un dígito
  suelto ahí bloquea igual que en la prosa.
- Un color de serie con poco contraste contra el fondo obliga al tool a
  etiquetar todos los valores de esa gráfica.
- **Escala:** solo `bars` tiene `scale=`. En las demás se divide el `Num` antes
  de pasarlo (`valor / 1e6`); la constante queda en la expresión de origen.
- **Componer visuales:** `visual=` es HTML; dos visuales se concatenan con `+`.
  Quedan pegados (las marcas de eje del primero parecen una fila de la tabla
  que sigue): envuelve el segundo en `<div style="margin-top:16px">…</div>`.
- **`bars`, `lines`, `waterfall` y `football` crecen con `h=`.** En `football`
  el alto se reparte entre las filas: más separación y barras más gruesas. Un
  `h` menor a 54 por fila más 60 se rechaza (las etiquetas se tocarían). `table` y
  `heatmap` miden lo que midan sus filas. Subir `h=` llena el espacio, no
  agrega contenido: un football de un solo método sigue siendo de un solo
  método (regla de densidad del guion).
- **Alineación en `table`:** la primera columna va a la izquierda. De las
  demás, la que no trae cifras (ningún `Num` ni `lit()` en sus celdas) es de
  texto y también va a la izquierda; las de cifras, a la derecha. Se decide por
  columna completa: no mezcles prosa y cifras en la misma columna.

## 4. Regla de bloqueo y cuándo usar `lit()`

Tras renderizar, el tool recorre todo el texto visible, incluido el de cualquier
SVG escrito a mano. `build()` falla y no escribe el archivo si encuentra un
número sin origen.

| En el texto | Resultado |
|---|---|
| Número que no viene de un `Num`, con o sin unidad o moneda: `38%`, `11x`, `180pbs`, `US$5`, `2,150` | Bloquea |
| El mismo con símbolos o en rango: `~38%`, `>38%`, `n=30`, `10%-12%` | Bloquea |
| Año del eje del modelo, escrito solo: `2027`, `(2027)` | Pasa |
| Año fuera del eje, o con símbolo: `1987`, `$2027` | Bloquea |
| Dígitos pegados a letras que no son unidad: `N2`, `A14`, `3nm`, `IFRS16`, `1Q2026E`, `2T26`, `FY2026` | Pasa |
| Dígitos, guion, letras: `20-F`, `10-K`, `COVID-19` | Pasa |
| Nombre con el número separado: `IFRS 16`, `S&P 500`, `top 3`, `26 de noviembre` | Bloquea |
| `lit("4", why="pasos del método")` | Pasa y queda listado (P3) |
| La línea de fuente | No se escanea: cita años y páginas |

Cómo escribir sin pelear con la regla:

- **Cifra del negocio**: siempre por f-string de un `Num`. Si no está en el
  modelo ni en los hechos, no va en el deck.
- **Conteo editorial** ("cuatro pasos", "tres condiciones"): en letra. Con
  dígito, `lit("4", why="pasos del método")`. Un conteo sobre el negocio
  ("durante cinco años", "tres clientes") no es editorial aunque vaya en letra:
  es una afirmación y lleva evidencia en la verificación.
- **Nombre propio con número** ("IFRS 16", "S&P 500"): `lit("16", why="nombre de
  la norma")`, o se reescribe.
- **Año fuera del eje o fecha de evento**: es una afirmación sobre la emisora.
  Va como hecho con su cita (`deck.fact("fundacion")` con formato `.0f`), no como
  `lit()`.
- **Unidad**: se escribe junto al número formateado (`f"{x:,.0f} millones"`,
  `f"{deck.fact('share'):.1f}%"`). La unidad no lleva dígitos; pasa.

`lit()` no es salida para una cifra del negocio. Cada uso, con su porqué, sale
en el reporte de verificación y el analista lo ve en el gate 2. Lo mismo las
constantes dentro de un derivado (`* 100`, `/ 1e6`): quedan en la expresión de
origen.

## 5. Construir, medir, ajustar

`measure` no ajusta nada: reporta por lámina y el script se corrige. Lo normal
son dos o tres rondas. Mientras se itera basta la consola; `--report` es para
el cierre (sección 7).

| Falla | Qué hacer |
|---|---|
| P6, lámina vacía (el detalle da los px libres y la altura del visual) | En `bars`, `lines`, `waterfall` y `football`: subir `h=` en esa cantidad. En `table` y `heatmap`: más filas o un segundo visual del mismo hallazgo |
| P5, contenido encimado o cortado | Bajar `h=`, acortar la Nota o la entradilla, o partir la lámina |
| P7, desborde horizontal | Acortar el título o el rótulo; en tablas, menos columnas |
| P8, etiqueta fuera de la gráfica o encimada con otra | Menos periodos en `show=`, otro `fmt`, rótulos de serie más cortos |
| P8, nombres de serie encimados (`encimadas 'A / B'`) | Dos líneas que terminan en el mismo valor apilan sus nombres: recorta `keys=` a un tramo donde difieran, o grafica niveles (`apply`) en lugar de tasas |

**La medición no decide el contenido.** Si para pasar P6 o P8 hace falta otro
visual, más filas con hechos nuevos o un hallazgo que el guion no tenía, el
cambio se anota en el outline (sección "Cambios posteriores al gate 1") y se
presenta como excepción en el gate 2. Un hallazgo nuevo que cambie el argumento
vuelve al analista antes de seguir.

El SVG se escala al ancho útil de la lámina (`1280 - 2 * deck_margin`): con
`w=1100` y margen 56, cada 10 unidades de `h` son unos 10.6 px en pantalla.

Las gráficas se dimensionan para el espacio que deja la Nota, no al revés.

## 6. Checks P

Mismos códigos de salida que `tools/model_check.py`: 0 verde, 1 con `[x]`, 2 uso
incorrecto, 3 con `[pendiente]`. Sin navegador (Edge, Chrome o Chromium; la
variable `RA_BROWSER` fija la ruta) la medición queda `[pendiente]`, nunca
verde. Los IDs no se reciclan.

| # | Check | Cómo se detecta | Comando | Efecto |
|---|---|---|---|---|
| P1 | Cifras iguales a su origen | Relee el modelo y los hechos, re-evalúa cada expresión de origen con su formato y la compara contra el texto mostrado. Detecta un deck viejo contra un modelo nuevo y un número retocado | `audit` (y `build`) | Bloquea |
| P2 | Sin números sueltos | Escaneo del texto visible con la regla de la sección 4 | `audit` (y `build`) | Bloquea |
| P3 | Literales declarados con `lit()` | Lista cada uso con su porqué | `audit` | Avisa: se revisan en el gate 2 |
| P4 | Hechos con verificación manual | Hechos usados cuya fuente no es texto (PDF) | `audit` | Avisa: se cotejan a mano en el gate 2 |
| P5 | Sin contenido encimado | Distancia entre el fin del contenido y el cierre mayor que cero; ninguna lámina (tampoco portada, acto o aviso legal) con contenido cortado por su marco | `measure` | Bloquea |
| P6 | Sin lámina vacía | Esa distancia menor que `deck_gap_max` (90 px por defecto) | `measure` | Bloquea |
| P7 | Sin desborde horizontal | Ningún bloque de texto más ancho que su caja | `measure` | Bloquea |
| P8 | Etiquetas dentro de su gráfica | Ninguna etiqueta SVG fuera del área de su gráfica ni dos con sus cajas encimadas | `measure` | Bloquea |
| P9 | Contraste del texto | Razón WCAG del texto contra el fondo, mínimo 4.5 | `audit --design` | Avisa |
| P10 | Series distinguibles con daltonismo | Distancia entre cada par de colores de serie bajo deuteranopia | `audit --design` | Avisa |

P9 y P10 evalúan la paleta con la que se generó el deck. `audit --design
brand/DESIGN.md` se pasa siempre, igual que `Deck(design=...)`, y la ruta
puede ser relativa a la cobertura: con el archivo evalúa la marca del
analista; sin él, la paleta neutra del tool (que pasa ambos por prueba del
repo), y el título del grupo lo dice ("paleta neutra: no existe ..."). El deck
lleva anotados sus colores: si no son los de `--design` (ruta equivocada, o
`DESIGN.md` cambió después de generar), `audit` falla en lugar de evaluar una
paleta que el deck no tiene. Una carpeta en lugar de un archivo es error.

Subcomandos que no emiten checks P: `labels` (mapa del modelo, sección 3),
`facts` (registro de hechos, sección 2; mismos códigos de salida) y `claims`
(tabla de afirmaciones, sección 7).

## 7. Verificación del gate 2

Cuatro partes. Las dos primeras las hace el código; las otras dos, quien
verifica, sobre listas que el código garantiza completas.

**Cifras y espacio.** Hasta que ambos salgan en verde (o `measure` en
`[pendiente]` por falta de navegador, que se reporta como tal):

```
python tools/deck_builder.py audit <deck.html> <modelo.xlsx> --facts <hechos.json> --design brand/DESIGN.md --report research/notes/<ticker>-deck-verify_<fecha>.md
python tools/deck_builder.py measure <deck.html> --shots research/notes/shots_<fecha> --report research/notes/<ticker>-deck-verify_<fecha>.md
```

`--facts` solo si hay registro. El reporte se reescribe las veces que haga
falta sin perder lo ya trabajado:

- `audit --report` conserva la columna Evidencia de cada afirmación que sigue
  igual: misma lámina, mismo origen (columna Dónde) y mismo texto, cifras
  incluidas. Si la lámina cambió de número, la conserva solo cuando ese texto
  es único en la tabla; repetido y renumerado (un encabezado "Valor" en dos
  tablas) no se adivina. Una frase que cambió pierde la suya: era evidencia de
  otra afirmación. La consola dice cuántas se conservaron y cuántas quedaron
  sin frase.
- **Marca `[revisar]`.** El reporte anota la huella de los valores del modelo y
  de los hechos. Si cambian entre corridas, la evidencia conservada sale con
  `[revisar]` al frente: se escribió contra otros números, aunque la frase no
  tenga ninguno ("el ritmo proyectado es menor al observado"). Se coteja de
  nuevo y se le quita la marca a mano. Ninguna marca llega al gate 2.
- `measure --report` agrega o reemplaza la sección "Espacio" (P5 a P8) sin
  tocar el resto, con la huella del deck. `audit --report` la conserva
  mientras el deck sea el mismo; si el deck cambió, la deja en `[pendiente]`
  hasta volver a medir. Al cerrar: primero `audit --report`, después `measure
  --report`.
- `python tools/deck_builder.py claims <deck.html> <salida.md>` da la misma
  tabla en un archivo aparte. No arrastra evidencia ni pisa un archivo que ya
  la tenga: la evidencia vive en el reporte.

**Evidencia por frase.** La sección "Afirmaciones" lista todo el texto visible
que trae palabras, y la columna Dónde dice de dónde viene:

| Dónde | Qué lista |
|---|---|
| `texto` | Títulos, entradillas, Notas, cierres y el texto de una lámina a medida, frase por frase |
| `tabla` | Cada celda o encabezado con texto. Una celda que solo trae su cifra no: ya está en "Cifras y su origen" |
| `portada` | Cada cifra de portada con su rótulo |
| `gráfica` | Rótulos de serie, de método y de paso, el rótulo del marcador y los textos de un SVG a medida |

No lista lo que pone el tool (marcas de eje, periodos, "Proyectado", número de
página) ni la línea de fuente. Cada fila recibe evidencia: la celda del modelo
(ya está en "Cifras y su origen"), o el documento con su localizador. En un
rótulo, la evidencia es que dice lo mismo que la etiqueta de la fila del
modelo. Una frase sin evidencia se elimina del script, se regenera el deck y
se repite la corrida.

Quien redactó una frase tiende a darla por buena. Si el runtime permite
subagentes, esta parte la hace uno que no redactó el deck, con el deck, el
modelo, los hechos y los documentos fuente, y la consigna de encontrar lo que no
se sostiene, tablas y rótulos incluidos. En el dogfood, la pasada propia cambió
tres frases y el revisor independiente encontró trece más. Sin subagentes: la
pasada se hace igual y el reporte dice que no hubo revisor independiente.

Lo que el bloqueo no ve y esta pasada sí debe ver. Llevan evidencia aunque no
tengan número:

- "récord", "primero", "único", "siempre", "cada año";
- conteos en letra sobre el negocio ("durante cinco años": se cuenta año por
  año en la serie);
- fechas de eventos escritas con un año del eje;
- comparaciones contra competidores;
- un rótulo que renombra la fila del modelo (el deck debe decir lo mismo que la
  etiqueta de la celda, o explicar la diferencia);
- un título que mezcla periodos (un dato trimestral junto a uno anual);
- "el mayor", "el principal" cuando la identidad cambia entre periodos (el
  cliente mayor de un año no es el del siguiente);
- moneda ambigua ("dólares" en un deck que reporta en otra moneda);
- una cifra de un tercero que no cuadra con las demás cifras de su propia
  fuente: se resuelve o se retira, no se presenta con reserva en letra chica.

**Capturas, una por una.** `measure` en verde dice que no hay hueco ni
encimado; no dice que la lámina se vea bien. En el dogfood hubo defectos
visibles con la medición en verde. Se abre cada `slide-NN.png` y se revisa:

- el título dice el hallazgo y cabe sin partirse mal;
- la gráfica se lee: escala razonable, etiquetas que no se tocan aunque no se
  encimen, rótulo "Proyectado" donde corresponde;
- las series se distinguen y cada una tiene nombre;
- la tabla no está apretada ni tiene columnas que sobran;
- la Nota narra lo que muestra el visual, no otra cosa.

Sin visión en el runtime, las capturas las revisa el analista en el gate.

**Texto.** Errores que el tool no atrapa y que aparecieron en decks reales:

- Diferencia entre dos tasas escrita como porcentaje. Va en puntos
  porcentuales: `f"{(m27 - m25) * 100:+.1f} puntos porcentuales"`, nunca
  `f"{m27 - m25:+.1%}"`.
- Contrastes "no es X, sino Y". Se escribe lo que es.
- Frases de efecto al cierre y metadiscurso ("lo importante es", "cabe
  destacar").
- Signo o escala alrededor de una cifra válida: `-{x}` o "millones" donde la
  serie está en miles. El tool valida el número, no el sentido que lo rodea.

## 8. Límites conocidos

- El bloqueo solo atrapa números. Una afirmación sin número pasa: queda para la
  evidencia por frase.
- La evidencia la evalúa el asistente, a veces sobre texto que él mismo redactó.
  El código lista todo el texto visible con palabras (prosa, tablas, portada y
  gráficas); no garantiza que el juicio sea correcto.
- La evidencia se conserva por texto exacto de la frase: cambiar una coma o una
  cifra la suelta y hay que volver a escribirla. La marca `[revisar]` se
  dispara con cualquier cambio de valores del modelo o de los hechos, toque o
  no a esa frase: el tool no sabe de qué celdas depende una afirmación.
- `facts --text` aplana la fuente: no conserva párrafos ni la forma de las
  tablas. Sirve para copiar la cita, no para leer el documento.
- `table` decide la alineación por columna completa: una columna que mezcla
  prosa y cifras queda a la derecha.
- Atributos (`alt`, `title`) y la numeración de `<ol>` no se escanean ni se
  listan como afirmaciones.
- Graficar hechos exige armar la `Series` a mano (`from deck_builder import
  Series`); no hay control de densidad en las tablas.
- El umbral de 90 px de lámina vacía salió de un solo deck; se cambia con
  `deck_gap_max` en `brand/DESIGN.md`.
- macOS: medición sin verificar.
