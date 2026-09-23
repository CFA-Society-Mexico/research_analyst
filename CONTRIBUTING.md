# Contribuir a research_analyst

Gracias por querer aportar. Este plugin lo usa gente que publica research con su
firma: un bug aquí no rompe un build, contamina una nota de inversión. Por eso el
protocolo es estricto en lo que importa (invariantes, trazabilidad, citas) y
ligero en lo demás.

Antes de escribir una línea, lee las dos páginas que gobiernan el diseño:
[`docs/architecture.md`](docs/architecture.md) (el porqué de cada decisión) y
[`AGENTS.md`](AGENTS.md) (cómo se enruta el trabajo).

---

## 1. Las reglas que ningún PR puede romper

Si tu cambio viola una de estas, se cierra sin importar qué tan bien esté escrito.

| # | Invariante | Qué significa en la práctica |
|---|---|---|
| R1 | **El modelo nunca calcula ni decide** | Ninguna cifra puede nacer del texto de una skill. Sale de una fórmula de Excel o de código determinista. Si tu cambio hace que el asistente "estime", "promedie" o "ajuste" un número, está mal. |
| R2 | **Trazabilidad de 4 etiquetas** | Todo dato es `observado` (con documento y página), `guidance` (management), `supuesto` (analista) u `output` (fórmula). Un artefacto nuevo sin etiquetas no se acepta. |
| R3 | **Citas normativas exactas o `[VERIFICAR]`** | Solo `framework-mapper` emite citas (ASC, NIC/NIIF, NIF, CNBV). Ninguna otra skill, ningún tool. Sin fuente primaria verificada, se marca `[VERIFICAR]` — nunca se aproxima. |
| R4 | **Cero recomendaciones de inversión** | El plugin estructura y cuestiona. Comprar/vender/mantener y el precio objetivo llevan la firma del analista, no la del modelo. |
| R5 | **Nada se borra** | En una carpeta de cobertura solo se agrega o se archiva fechado. Retención mínima 7 años. Ningún código del plugin puede tener una ruta que borre un filing, un journal o una versión del modelo. |
| R6 | **Ownership único de artefactos** | Cada archivo tiene exactamente UNA skill que lo escribe (tabla en `docs/architecture.md`). Las demás leen. Excepción diseñada: `thesis-journal.md`, append-only multi-skill. |
| R7 | **Sin APIs de pago ni keys privadas en el repo** | Fuentes públicas y gratuitas (EDGAR, XBRL, FRED). Las keys las pone el usuario en su workspace (`macro/fred.key`, variable de entorno), jamás en el repo, jamás impresas en consola. |

---

## 2. Mapa del repo — qué toca cada cosa

```
commands/            orquestadores: preflight + secuencia de skills + gates. CERO lógica propia.
skills/*/SKILL.md    unidades de conocimiento: una skill hace UNA cosa y no llama a otra.
skills/*/references/ material largo que la skill carga bajo demanda (checks, mapeo normativo).
templates/           CONTRATOS entre skills (yaml y specs). Cambiarlos rompe consumidores: ver 4.4.
tools/*.py           lo determinista: descarga de datos y construcción/audit del xlsx.
docs/architecture.md el porqué del diseño.
AGENTS.md            router para runtimes sin skills nativas (Cursor, Codex).
```

Regla de composición: **skill nunca llama a skill**. Si dos skills se necesitan,
quien las secuencia es un comando.

---

## 3. Setup local

Requisitos: Python 3.10+ y `openpyxl` (única dependencia de terceros; el resto es
stdlib a propósito, para que los tools corran en cualquier sandbox).

```bash
git clone https://github.com/CFA-Society-Mexico/research_analyst && cd research_analyst
```

```bash
python -m pip install openpyxl
```

Prueba de humo del builder (escribe un libro demo y corre encima el audit de
formato completo):

```bash
python tools/xlsx_builder.py demo ./scratch/demo.xlsx
```

Para probar cambios de skills y comandos: instala el repo local como plugin en
Claude Code (`/plugin marketplace add <ruta local>`) o, en Cursor/Codex, abre el
repo y deja que `AGENTS.md` enrute. **Haz dogfood contra una emisora real** — la
mayoría de los bugs de este plugin solo aparecen con un filing de verdad.

---

## 4. Tipos de contribución y su protocolo

### 4.1 Editar una skill (`skills/*/SKILL.md`)

- Mantén el tamaño: las skills viven en ~100 líneas. Si tu cambio la infla, el
  material largo va a `references/` y la skill lo carga bajo demanda.
- El frontmatter (`name`, `description`) es lo que dispara la skill. Si tocas la
  `description`, agrega los disparadores en el lenguaje natural que usa un
  analista ("archiva este 10-Q", "corre los checks"), no jerga interna.
- No dupliques doctrina que ya vive en otra skill o en un template: enlaza.
- Si la skill empieza a hacer dos cosas, no la crezcas — abre issue para discutir
  el split.

### 4.2 Agregar o cambiar un check de integridad

Es la contribución más valiosa y la más regulada. Cuatro familias:

| Familia | Dónde vive la lógica | Cómo se verifica |
|---|---|---|
| **S** (estructurales) | escaneo por código sobre el libro | debe ser detectable sin abrir Excel |
| **C** (contables) | fórmulas en la tab `Checks` del modelo | la fórmula la escribe el builder, no el analista |
| **D** (doctrina) | reglas del plugin (drivers, staleness, reconciliación) | avisa o bloquea: declara cuál de los dos |
| **F** (formato) | `tools/xlsx_builder.py audit` | debe fallar en un libro malo y pasar en el `demo` |

Un check nuevo llega completo o no llega:

1. Fila en [`skills/model-standards/references/integrity-checks.md`](skills/model-standards/references/integrity-checks.md)
   con ID, enunciado y **cómo se detecta**.
2. Implementación (código para S/F, fórmula del builder para C, regla escrita
   para D).
3. Declara si **bloquea** o solo **avisa**. Un check que bloquea con falsos
   positivos es peor que no tenerlo: si el caso legítimo existe (año de
   transición, serie descontinuada), la excepción va escrita en la misma fila.
4. Prueba en los dos sentidos: un libro que debe fallarlo y el `demo` que debe
   pasarlo.

Los IDs no se reciclan. Un check derogado se marca derogado, no se borra ni se
reasigna su número.

### 4.3 Tocar un tool (`tools/*.py`)

- **stdlib por default.** Una dependencia nueva necesita justificación explícita
  en el PR; `openpyxl` es la única aceptada hoy.
- **stdout ASCII-only.** Sin flechas, sin acentos, sin emojis en `print()`, en la
  ayuda de argparse ni en mensajes de error: la consola de Windows es cp1252 y
  revienta con `UnicodeEncodeError`. Usa `->`, `[ok]`, `[x]`, `[aviso]`. El
  contenido de los archivos generados (Markdown, CSV) sí puede llevar acentos.
- **Determinismo.** Mismo input, mismo output byte a byte hasta donde se pueda.
  Nada de reloj local metido en un artefacto (usa la fecha de la última
  observación del dato, como hace `fred_fetch`).
- **Red bloqueada es un caso esperado, no un traceback.** Sigue el patrón de
  `sec_fetch` / `xbrl_fetch` / `fred_fetch`: mensaje accionable que nombra el
  dominio a permitir y qué hacer si no se puede.
- **Las keys jamás se imprimen**, ni siquiera dentro de una URL en un mensaje de
  error. Redáctalas.
- **El xlsx solo se escribe vía `ModelStyler` de `tools/xlsx_builder.py`.**
  openpyxl crudo para estructura o formato es rechazo automático: el formato es
  código, no criterio.

### 4.4 Cambiar un template (`templates/*`)

Los templates son **contratos**, no ejemplos. Cambiar un campo rompe a quien lo
lee. Un PR de template incluye:

1. El cambio del template.
2. La actualización de TODA skill que lo lee o lo escribe (grep del nombre del
   campo).
3. Qué pasa con las coberturas ya existentes: campo nuevo opcional con default, o
   ruta de migración escrita. Nunca dejes un workspace viejo ilegible.

### 4.5 Agregar un comando (`commands/*.md`)

Un comando es un orquestador: preflight, secuencia de skills, gate de usuario
entre pasos y protocolo de debate al cierre de cada gate
([`templates/debate-protocol.md`](templates/debate-protocol.md)). Si tu comando
contiene lógica de análisis, esa lógica pertenece a una skill.

Preflight que falla = **detente**. Nunca corras parcial: media cobertura es peor
que ninguna.

### 4.6 Agregar una skill nueva

Alto costo, gate alto. El repo tiene 8 skills por diseño (el tope original era 6
y la enmienda está documentada en `docs/architecture.md`). Abre issue ANTES de
escribir código y responde: de qué artefacto es dueña, por qué no cabe en una
existente, qué gate propio tiene. Si la respuesta es "para no inflar otra skill",
el problema real probablemente es que la otra skill necesita `references/`.

### 4.7 Citas normativas

Solo entran a
[`skills/framework-mapper/references/ifrs-asc-nif-line-map.md`](skills/framework-mapper/references/ifrs-asc-nif-line-map.md),
con **fuente primaria** (el estándar mismo o el boletín del emisor — no un blog
de una Big 4 resumiendo el estándar). Sin fuente primaria: `[VERIFICAR]`, que es
una contribución válida y útil. Una cita inventada es el peor bug posible de este
repo: convierte un error del modelo en un error de cumplimiento del analista.

---

## 5. Estilo de escritura (documentación y salida al usuario)

- **Documentación al usuario en español; nombres de archivos, skills y código en
  inglés.**
- **Sin emojis ni iconos decorativos** en la salida al usuario. Estados: `[ok]`,
  `[x]`, `[aviso]`, `[n/a]`.
- Tono de nota de research: directo, con fuente, sin celebraciones ni relleno.
- Cifras con separador de miles y unidades en el encabezado, no en cada celda.
- Español con acentos en Markdown; ASCII estricto en stdout (4.3) y en mensajes
  de commit.

---

## 6. Verificación antes de abrir el PR

No hay suite de tests todavía (construirla es una contribución abierta, sección
8). Hoy la verificación es esta lista, y se corre completa.

Sintaxis de todos los tools:

```bash
python -c "import ast,glob;[ast.parse(open(f,encoding='utf-8').read()) for f in glob.glob('tools/*.py')];print('[ok] sintaxis')"
```

Builder y audit de formato en verde:

```bash
python tools/xlsx_builder.py demo ./scratch/demo.xlsx
```

Y a mano:

- [ ] Si tocaste un tool: corrió contra un caso real (un ticker de verdad) y
      también por el camino de error (sin red, sin key) sin traceback.
- [ ] Si tocaste el builder o un check: el `demo` pasa en verde Y un libro
      deliberadamente roto falla en el check nuevo.
- [ ] Si tocaste una skill o comando: lo ejecutaste end-to-end en un workspace de
      dogfood, no solo lo leíste.
- [ ] `grep` de los nombres que cambiaste: ninguna referencia colgada en
      `skills/`, `commands/`, `templates/`, `AGENTS.md`, `README.md`.
- [ ] Ningún archivo de una cobertura real (filings, modelos, journals, keys) se
      coló en el commit. **Nunca subas datos de emisoras ni transcripts.**
- [ ] Consola ASCII-only.

---

## 7. Commits, versionado y PRs

**Commits** — Conventional Commits, sujeto en ASCII puro (sin acentos), 72
caracteres o menos, en español:

```
feat: F19 roll de caja cerrado + xbrl_fetch captura el CF completo
fix: D2 excluye el anio de transicion (trimestres observados parciales)
docs: regla UDM/LTM = exactamente 4 trimestres (t-3..t)
chore: bump a v0.5.2 - doctrina completa de la auditoria
```

Prefijos en uso: `feat`, `fix`, `docs`, `chore`. El cuerpo explica el **porqué**
cuando no es obvio del sujeto; el qué ya está en el diff.

**Versionado** — semver en `.claude-plugin/plugin.json`. El bump va en su
**propio commit** (`chore: bump a vX.Y.Z - <resumen>`), después del commit
funcional. Criterio: `patch` = fix o refinamiento de un check; `minor` = check
nuevo, tool nuevo, capacidad nueva; `major` = cambio de contrato que rompe
coberturas existentes.

**PRs** — uno por tema. Un PR que arregla un check y de paso reordena tres skills
no se puede revisar. Incluye:

1. Qué problema resuelve (con el caso real que lo destapó, si lo hubo).
2. Qué invariante de la sección 1 podría rozar y por qué no la rompe.
3. Cómo lo verificaste (la lista de la sección 6, con el resultado).
4. Si tocaste contratos: qué pasa con las coberturas ya existentes.

---

## 8. Fronteras abiertas — dónde más falta ayuda

Cada una es un PR bienvenido; las tres primeras son las que más duelen hoy.

| Frontera | Qué falta |
|---|---|
| **Bancos** | Mapeo Anexo 33 (criterios CNBV) vs IFRS 9, línea por línea. Hoy el plugin detecta la emisora financiera y se detiene, a propósito. |
| **Aseguradoras** | Mapeo CNSF vs IFRS 17. Mismo estado. |
| **Suite de evals** | No existe. Casos dorados por skill: un filing de entrada, la captura esperada, el veredicto esperado de cada check. |
| **Citas NIF** | Varias marcadas `[VERIFICAR]` esperando fuente primaria del CINIF. |
| **Conector Banxico SIE** | Series MX (token gratuito) bajando a `macro/series/` como CSV + manifest — mismo patrón que `tools/fred_fetch.py`, que ya cubre FRED. |
| **AFFO de FIBRAs** | Convención de ajuste no cerrada. |
| **Comparabilidad entre marcos** | Hoy las 7 diferencias más comunes están verificadas; el resto se marca `[VERIFICAR]`. |

---

## 9. Reportar bugs

Abre un issue con: qué esperabas, qué pasó, el comando exacto y **datos
anonimizados**. Nunca pegues filings privados, transcripts, modelos de clientes,
API keys ni nada que pueda ser información privilegiada (MNPI) — el plugin tiene
un guard de MNPI en el preflight precisamente porque este riesgo es real. Si el
bug solo se reproduce con un documento sensible, descríbelo y ofrece reproducirlo
tú con guía.

Si el bug produce **una cifra mal** o **una cita normativa mal**, márcalo como
crítico en el título: esos dos son los únicos que llegan al reporte publicado del
analista.

---

## 10. Licencia

Al contribuir aceptas que tu aporte se publique bajo [MIT](LICENSE). Proyecto
educativo y de productividad — **no constituye asesoría de inversión**.
