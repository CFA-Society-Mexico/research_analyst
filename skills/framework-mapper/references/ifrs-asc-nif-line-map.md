# IFRS vs US GAAP vs NIF — mapa norma→línea

Base: deep research con verificación adversarial (2026-08-25; 25 claims verificados
3 votos c/u, 0 refutados; fuentes primarias: texto NIF B-10, Anexo 33 CUB, SEC
33-8879/FRM; guías Big 4 2025-2026 solo como apoyo secundario para ubicar
párrafos). **Regla: toda cita fuera de esta tabla, o
marcada [VERIFICAR], se verifica en fuente primaria antes de usarse en una nota.**

## Diferencias que cambian una línea del modelo

### 1. Inventarios — ASC 330 · IAS 2 · NIF C-4 [VERIFICAR]
- LIFO aceptable SOLO bajo US GAAP; prohibido bajo IAS 2.
- Write-downs a VNR: base nueva irreversible bajo US GAAP (ASC 330-10-35-14);
  IAS 2 párr. 33 EXIGE reversa (limitada al monto original).
- **Líneas:** COGS, margen bruto, inventario.
- **Ajuste:** re-expresar a FIFO con la reserva LIFO revelada al comparar cruzado;
  nunca proyectar recuperaciones de write-downs en modelo anual US GAAP.

### 2. Deterioro de larga duración — ASC 360/350 · IAS 36 · NIF C-15 [VERIFICAR]
- US GAAP, activos de larga duración en uso (ASC 360-10): dos pasos (filtro de
  recuperabilidad con flujos NO descontados; si falla, pérdida a valor razonable).
  Goodwill (ASC 350-20): UN paso desde ASU 2017-04 — pérdida = exceso del valor
  en libros de la unidad de reporte sobre su valor razonable, topada al goodwill
  asignado; se eliminó el antiguo "paso 2" (valor razonable implícito del
  goodwill). Obligatorio para emisoras SEC en ejercicios que inician después
  del 15-dic-2019; conserva la evaluación cualitativa opcional.
  IAS 36: un paso (valor en libros vs mayor de FV−costos de disposición y valor en uso).
- Reversas: prohibidas US GAAP (salvo held-for-sale); IAS 36 obliga evaluarlas cada
  periodo (excepto goodwill).
- **Ajuste:** deterioros IFRS disparan antes; modelo IFRS puede traer upside por
  reversas que un modelo US GAAP no debe contener.

### 3. Arrendamientos — ASC 842 · IFRS 16 · NIF D-5 [VERIFICAR]
- ASC 842: modelo dual; el operativo es UN gasto lineal DENTRO de EBITDA.
  IFRS 16: modelo único (depreciación ROU + interés, ambos FUERA de EBITDA).
- **Ajuste clave de comps:** emisor IFRS 16 muestra EBITDA más alto que gemelo
  US GAAP con contratos idénticos — normalizar antes de comparar (fila de ajuste
  en Val_Comps cuando el universo mezcla marcos).
- IFRS 16: exención de bajo valor (~US$5,000, BC100) sin equivalente ASC 842.

### 4. Revaluación y desarrollo — ASC 730 · IAS 16.29-31 · IAS 38.57 · NIF C-8 [VERIFICAR]
- Revaluación PP&E: prohibida US GAAP; elección de política bajo IAS 16 (clase completa).
- Desarrollo: gasto US GAAP (ASC 730; software: ASC 985-20 / 350-40); IAS 38.57
  OBLIGA capitalizar con los 6 criterios de factibilidad.
- **Ajuste:** des-capitalizar desarrollo IFRS o capitalizar R&D US GAAP para
  comparar; superávit de revaluación en ORI de emisores IFRS/NIF.

### 5. Inflación — NIF B-10 · IAS 29 · ASC 830 ⚠ crítico para México
- **NIF B-10: re-expresión integral con acumulado trienal ≥ 26%** (Mejoras 2023
  quitaron el paréntesis "8% anual", NO el umbral). IAS 29: hiperinflación ~100%
  trienal. ASC 830: highly inflationary >100% trienal.
- Zona 26%-100%: reportante NIF re-expresa (no monetarios, capital, REPOMO) mientras
  gemelo IFRS/US GAAP presenta nominal — comparabilidad rota.
- Hoy: desconexión (trienal MX ~15-20%; IPTF nov-2025: México NO highly inflationary).
  Bases re-expresadas hasta dic-2007 se conservan como costo.
- **Reconexión = retrospectiva (NIF B-1): step-change en PP&E, inventarios, capital,
  depreciación.** Regla del plugin (check D7): emisora NIF con trienal ≥ 22% en
  macro-view ⇒ escenario de reconexión pre-modelado obligatorio.

### 6. PTU diferida — NIF D-3 · IAS 19 ⚠ exclusivo México
- NIF: PTU diferida por método de activos y pasivos (ref. NIF D-4).
  IFRS: PTU = beneficio a empleados IAS 19; SIN diferido (IFRIC: método de
  activos/pasivos de IAS 12 aplica solo a impuestos a la utilidad).
- **Ajuste:** al migrar NIF→IFRS eliminar PTU diferida del balance; tasa efectiva
  "impuestos + PTU" difiere entre marcos.

### 7. Pérdida crediticia esperada — ASC 326 · IFRS 9 · NIF C-16
- NIF C-16 (vigente 2018) ≈ IFRS 9: modelo por incremento de riesgo/etapas.
  ASC 326 (CECL): pérdida de por vida desde día uno — provisión front-loaded.
- **Ajuste:** reportante US GAAP carga provisión antes que gemelo IFRS/NIF.

## Sin claims verificados — [VERIFICAR] antes de citar

| Área | Normas |
|---|---|
| Ingresos (convergencia declarada D-1/D-2 ↔ IFRS 15) | ASC 606 · IFRS 15 · NIF D-1 |
| Impuestos diferidos | ASC 740 · IAS 12 · NIF D-4 |
| Clasificación intereses/dividendos en flujos (+ IFRS 18 desde 2027) | ASC 230 · IAS 7 · NIF B-2 |
| Presentación (utilidad de operación, ORI) | NIF B-3 |
| Pensiones más allá de PTU | ASC 715 · IAS 19 · NIF D-3 |

## Mapa regulatorio — quién reporta bajo qué

| Emisora | Marco | Base legal |
|---|---|---|
| BMV no financiera | IFRS obligatorio desde ejercicio 2012 | CNBV Boletín 056/2008; CUE Art. 78 (reformas 2009) |
| Banco / casa de bolsa | Criterios CNBV: Anexo 33 CUB — NIF con overrides (Criterio A-1 ¶3); supletoriedad A-4: CNBV→NIF→NIIF→US GAAP (ASC 105)→otros | Anexo 33 CUB (versión dic-2025) — **fuera de alcance v1** |
| Aseguradora | Criterios CNSF | fuera de alcance v1 |
| Privada no listada | NIF (CINIF) | — |
| FPI con ADRs | IFRS-IASB aceptado por SEC SIN reconciliación; condiciones: declaración explícita y sin reservas + dictamen que opine sobre ese cumplimiento. Variantes jurisdiccionales NO califican; GAAP local (incl. NIF) ⇒ reconciliación Item 18 | SEC Release 33-8879 (2007, vigente per FRM 6310.1/6310.3); Release 33-8959 |

## Fuentes primarias

Texto NIF B-10 (PDF IFT) · SEC 33-8879 + FRM Topic 6 · Anexo 33 CUB
(cnbv.gob.mx) · IAS 2/16/29/36/38 · perfil IFRS Foundation México · FASB ASU
2017-04 (https://storage.fasb.org/ASU2017-04.pdf).

## Fuentes secundarias (para ubicar el párrafo, nunca para citar)

EY US GAAP vs IFRS (ene-2026) · KPMG IFRS/US GAAP (nov-2025) · PwC (2026) ·
IPTF/PwC Q2-2025 y Q4-2025. Por CONTRIBUTING §4.7, una guía de Big 4 que
resume el estándar no es fuente primaria: la cita se verifica en el estándar
o en el boletín del emisor.

El dossier completo con evidencia y votos vive en el vault privado del autor
(`raw/ifrs-usgaap-nif-line-differences-2026-08-25.md`) y no está en este repo.
