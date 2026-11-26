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
| 6 | Crecimiento | `slide` | La emisora concentra la mayor parte de su mercado | tabla | — | `foundry_share` |

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
  una frase larga.
- Fuentes de texto (`.md`, `.txt`, `.htm`, `.html`, `.csv`, `.json`): si la cita
  no aparece, el tool no carga los hechos.
- Otras fuentes (PDF): el hecho se acepta y sale como "verificación manual"
  (check P4). Se coteja a mano en el gate 2.
- Dato sin fuente: sale del guion, o el analista aporta el documento a la
  cobertura. No se sustituye por memoria ni por una búsqueda sin archivar.
- Opinión o estimación de un tercero: se atribuye en la frase ("<fuente> estima")
  y lleva su hecho con ese documento.

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

Las rutas de `model`, `facts` y `design` son relativas a la cobertura. `facts`
solo se pasa si el archivo existe; `design` puede pasarse siempre (sin archivo,
estilo neutro). La salida es `research/notes/<ticker>-deck_<as_of>.html`.

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
el mismo bloqueo de cifras y la misma medición que las demás.

### Gráficas (devuelven el HTML del visual; son opcionales)

| Llamada | Notas |
|---|---|
| `deck.bars(serie, keys=None, show=(), fmt=",.0f", scale=1.0, tick_fmt=None, w=1100, h=320)` | Observado y proyectado en colores distintos, corte rotulado. `show`: periodos con etiqueta de valor. `scale` divide (1e6 para millones) |
| `deck.lines([(nombre, serie), ...], keys=None, show=(), fmt=".0%", tick_fmt=None, w=1100, h=320)` | Varias series, un eje. La primera fija periodos y corte |
| `deck.football([(método, mín, máx, centro), ...], marker=(rótulo, valor), fmt=",.0f", w=1100, h=None)` | Rangos de valor; `marker` es la línea del precio |
| `deck.waterfall([(rótulo, valor), ...], total=(rótulo, valor), fmt=",.0f", w=1100, h=320)` | Puente. El total se lee del modelo, no se suma aquí |
| `deck.heatmap(grid, rows, cols, fmt=",.0f", row_fmt=".1%", col_fmt=".1%", corner="", center=None)` | Sensibilidad. `center=(fila, columna)` marca el caso base |
| `deck.table(header, rows)` | Celdas de texto; los números llegan formateados desde un `Num` |

- `keys=` recorta los periodos (una lista del eje). Etiquetar pocos puntos: los
  que sostienen el hallazgo.
- Los rótulos de serie, de método y de paso también se escanean: un dígito
  suelto ahí bloquea igual que en la prosa.
- Un color de serie con poco contraste contra el fondo obliga al tool a
  etiquetar todos los valores de esa gráfica.

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
| Dígitos pegados a letras que no son unidad: `N2`, `A14`, `3nm`, `IFRS16`, `1Q2026E`, `FY2026` | Pasa |
| Dígitos, guion, letras: `20-F`, `10-K`, `COVID-19` | Pasa |
| Nombre con el número separado: `IFRS 16`, `S&P 500`, `top 3`, `26 de noviembre` | Bloquea |
| `lit("4", why="pasos del método")` | Pasa y queda listado (P3) |
| La línea de fuente | No se escanea: cita años y páginas |

Cómo escribir sin pelear con la regla:

- **Cifra del negocio**: siempre por f-string de un `Num`. Si no está en el
  modelo ni en los hechos, no va en el deck.
- **Conteo editorial** ("cuatro pasos", "tres condiciones"): en letra. Con
  dígito, `lit("4", why="pasos del método")`.
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
son dos o tres rondas.

| Falla | Qué hacer |
|---|---|
| P6, lámina vacía (el detalle da los px libres y la altura del visual) | Subir `h=` del visual en esa cantidad, o partir el contenido de otra forma |
| P5, contenido encimado o cortado | Bajar `h=`, acortar la Nota o la entradilla, o partir la lámina |
| P7, desborde horizontal | Acortar el título o el rótulo; en tablas, menos columnas |
| P8, etiqueta fuera de la gráfica o encimada con otra | Menos periodos en `show=`, otro `fmt`, rótulos de serie más cortos |

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

## 7. Verificación del gate 2

Cuatro partes. Las dos primeras las hace el código; las otras dos, quien
verifica, sobre listas que el código garantiza completas.

**Cifras y espacio.** Hasta que ambos salgan en verde (o `measure` en
`[pendiente]` por falta de navegador, que se reporta como tal):

```
python tools/deck_builder.py audit <deck.html> <modelo.xlsx> --facts <hechos.json> --design brand/DESIGN.md --report research/notes/<ticker>-deck-verify_<fecha>.md
python tools/deck_builder.py measure <deck.html> --shots research/notes/shots_<fecha>
```

`--design` solo si el archivo existe. `--report` reescribe el reporte completo
en cada corrida: la columna de evidencia se llena sobre el de la última.

**Evidencia por frase.** La sección "Afirmaciones" del reporte lista cada frase
visible del deck. Cada fila recibe evidencia: la celda del modelo (ya está en
"Cifras y su origen"), o el documento con su localizador. Una frase sin
evidencia se elimina del script, se regenera el deck y se repite la corrida.
Para iterar sin perder lo ya resuelto,
`python tools/deck_builder.py claims <deck.html> <salida.md>` da la misma tabla
en un archivo de trabajo. Si el runtime permite subagentes, esta parte la hace
uno que no redactó el deck: quien escribió una frase tiende a darla por buena.

Lo que el bloqueo no ve y esta pasada sí debe ver: "récord", "primero", "único",
"siempre", fechas de eventos escritas con un año del eje, y comparaciones contra
competidores. Llevan evidencia aunque no tengan número.

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
- La evidencia la evalúa el asistente sobre texto que él mismo redactó. El
  código garantiza que la lista está completa, no que el juicio sea correcto.
- Atributos (`alt`, `title`) y la numeración de `<ol>` no se escanean.
- Graficar hechos exige armar la `Series` a mano (`from deck_builder import
  Series`); no hay control de densidad en las tablas.
- El umbral de 90 px de lámina vacía salió de un solo deck; se cambia con
  `deck_gap_max` en `brand/DESIGN.md`.
- macOS: medición sin verificar.
