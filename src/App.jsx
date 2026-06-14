import { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ═══════════════════════════════════════════════════════════════════
// CONSTANTES E DATOS BASE
// ═══════════════════════════════════════════════════════════════════

const VERSIÓN = "3.0";
const STORAGE_KEY = "fgc_grella_sistema";

// ─── NIVEIS DE COMPETICIÓN ─────────────────────────────────────────
// FGC: probas do calendario autonómico galego
// RFEC .5: probas nacionais propostas por federacións autonómicas (H.5, I.5, K.5, L.5, P.5, G.5)
//          Celebradas en Galicia con normativa RFEC pero designación autonómica
// RFEC Copa/Camp: probas clase H.1, I.1, K.1, L.1, P.1, G.1 e Campionatos de España
// UCI: calendario UCI (C1, C2, C3, S1, S2...)
const NIVEIS_COMPETICIÓN = [
  { id: "FGC_LOCAL",   label: "Proba FGC / Copa Galicia (autonómica)",      org: "FGC"  },
  { id: "FGC_MINI",    label: "Mini BTT / Mini DH – Escolas FGC",           org: "FGC"  },
  { id: "FGC_CAMP",    label: "Campionato Galego",                           org: "FGC"  },
  { id: "RFEC_PUNTCO", label: "Proba .5 RFEC en Galicia (H.5 / I.5 / K.5 / L.5 / P.5 / G.5)", org: "RFEC" },
  { id: "RFEC_COPA",   label: "Copa de España (H.1 / I.1 / K.1 / L.1 / P.1 / G.1)",            org: "RFEC" },
  { id: "RFEC_CAMP",   label: "Campionato de España (CN)",                   org: "RFEC" },
  { id: "UCI_SERIE",   label: "Serie UCI / Copa do Mundo (C1/C2/C3/S1/S2)", org: "UCI"  },
];

const MODALIDADES = [
  { id: "XCO",     label: "BTT – Cross Country Olímpico (XCO)",  mini: false },
  { id: "XCM",     label: "BTT – Cross Country Maratón (XCM)",   mini: false },
  { id: "DH",      label: "BTT – Descenso (DH)",                 mini: false },
  { id: "RR",      label: "Ruta – Carreira en Liña",             mini: false },
  { id: "CR",      label: "Ruta – Contrarreloxo (CR)",           mini: false },
  { id: "CX",      label: "Ciclocross (CX)",                     mini: false },
  { id: "PISTA",   label: "Pista",                               mini: false },
  { id: "MINIBTT", label: "Mini BTT (escolas / base FGC)",       mini: true  },
  { id: "MINIDH",  label: "Mini DH (escolas / base FGC)",        mini: true  },
];

// Categorías estándar (Elite, Sub23, etc.)
const CATEGORÍAS_STD = [
  "Elite", "Sub-23", "Júnior", "Cadete", "Infantil", "Escolar",
  "Máster 30", "Máster 40", "Máster 50", "Máster 60+",
  "Feminino Elite", "Feminino Sub-23", "Feminino Júnior",
  "Feminino Cadete", "Feminino Infantil",
];

// Categorías Mini BTT / Mini DH (FGC – base e escolas)
// Idades orientativas segundo normativa FGC:
//   Promesa A/B  → 6-7 anos    | Promesa Fem → feminino 6-7
//   Benjamín A/B → 8-9 anos    | Benjamín Fem → feminino 8-9
//   Alevín A/B   → 10-11 anos  | Alevín Fem  → feminino 10-11
//   Principiante A/B → 12-13 anos | Principiante Fem
//   Infantil / Infantil Fem → 14-15 anos (comparten con BTT estándar)
const CATEGORÍAS_MINI = [
  "Promesa A", "Promesa B", "Promesa Fem.",
  "Benjamín A", "Benjamín B", "Benjamín Fem.",
  "Alevín A",  "Alevín B",  "Alevín Fem.",
  "Principiante A", "Principiante B", "Principiante Fem.",
  "Infantil", "Infantil Fem.",
];

// Función auxiliar para obter as categorías segundo modalidade
function getCategorías(modalidadeId) {
  const mod = MODALIDADES.find(m => m.id === modalidadeId);
  return mod?.mini ? CATEGORÍAS_MINI : CATEGORÍAS_STD;
}

const CATEGORÍAS = [...CATEGORÍAS_STD, ...CATEGORÍAS_MINI];

// ═══════════════════════════════════════════════════════════════════
// NORMATIVA REAL – FGC (13/12/2025) + RFEC (BTT: 8.04.26 · CX: 9.09.25 · DX: 2.06.26)
// ═══════════════════════════════════════════════════════════════════

// ── XERARQUÍAS POR NIVEL ─────────────────────────────────────────
// Cada array é a orde de prioridade para construír a grella de saída.
// Claves: "Campión_Galicia" | "Campión_España" | "Copa_Galicia" | "Copa_España" |
//         "UCI" | "RFEC" | "FGC" | "Inscrición"
const XERARQUÍAS_BASE = {

  // ── NIVEL AUTONÓMICO FGC ────────────────────────────────────────
  // Copa Galicia XCO · Art. III-H.4.1 FGC
  // 1º Líder Copa Galicia · 2º Pts UCI (non aplica Infantil) · 3º Pts RFEC ·
  // 4º Clasificación Copa Galicia (1ª proba: copa ano anterior) · 5º Inscrición
  FGC_LOCAL:   ["Copa_Galicia", "UCI", "RFEC", "FGC", "Inscrición"],

  // Mini BTT-XC · Art. III-N.1.6.3 / Mini DH · Art. III-N.2.1 FGC
  // Sen puntos UCI nin RFEC. Só ranking FGC mini ou inscrición.
  // Mini DH: orde INVERSA ao ranking (o mellor sae o último na manga clasificatoria)
  FGC_MINI:    ["FGC", "Inscrición"],

  // Campionato Galego XCO · Art. III-F.5.1 FGC
  // 1º Campión Galicia ano anterior (se mantén categoría) · resto = Copa Galicia
  FGC_CAMP:    ["Campión_Galicia", "Copa_Galicia", "UCI", "RFEC", "FGC", "Inscrición"],

  // ── NIVEL NACIONAL RFEC – PROBAS .5 (autonómicas no calendario RFEC) ─
  // Clases H.5-XCO, I.5-DH, K.5-XCM, Ku.5-XCUM, L.5-Enduro, P.5-PumpTrack, G.5-Gravel
  // Aplícase a mesma normativa que Copa España da disciplina correspondente
  // XCO Art. IV-G.6 (páx 15 PDF RFEC): UCI vigente → Copa España XCO → Ranking RFEC → Inscrición
  // DH Art. IV-H manga clasif.: orde inversa UCI → Copa España → Ranking RFEC → Inscrición
  RFEC_PUNTCO: ["UCI", "Copa_España", "RFEC", "Inscrición"],

  // ── COPA DE ESPAÑA RFEC ─────────────────────────────────────────
  // XCO cat. UCI (Elite/Sub23/Júnior) · Art. IV-G.6 RFEC (8.04.26)
  //   Ranking UCI vigente → Copa España XCO vigente → (1ª proba: Ranking RFEC) → Inscrición
  // XCO cat. non-UCI (Cadete/Máster):
  //   Copa España XCO vigente → (1ª proba: Ranking RFEC) → Inscrición
  // DH manga clasif. · Art. IV-H.5 RFEC: orde inversa por categorías:
  //   Inf.M/F → por orde inversa inscrición excepto: Copa España provisional inversa →
  //   Ranking RFEC DHI → Pts UCI (menor a maior) → Campión España ano anterior (o último)
  // CX Copa España · Art. V-J.9 RFEC (9.09.25):
  //   1º Pts UCI · 2º Clasificación Copa España CX (1ª proba: final ano anterior) · 3º Ranking RFEC CX · 4º Inscrición
  // XCM Copa España · Art. IV-J.1 RFEC:
  //   Grupo1: Elite/Sub23 con pts UCI XCM → pts UCI XCO → resto Elite/Sub23
  //   Grupo2: Masters 25 primeiros Copa España XCM · Grupo3: resto a discreción
  RFEC_COPA:   ["UCI", "Copa_España", "RFEC", "Inscrición"],

  // ── CAMPIONATO DE ESPAÑA RFEC ───────────────────────────────────
  // XCO · Art. IV-F.5 RFEC (8.04.26):
  //   1º Campión España ano anterior (mantén categoría) · 2º Pts UCI (non aplica Cadete) ·
  //   3º Ranking RFEC BTT XCO (ou Copa España XCO ano en curso) · 4º Seleccións FFAA · 5º Inscrición
  // DH manga clasif. · Art. IV-F.6.2 RFEC (13.11.25):
  //   Categorías pola clasificación Copa España/Ranking RFEC en orde inversa.
  //   Protección UCI en manga final: os que non entran en top-10M / top-5F sairán "protexidos" detrás.
  // CX Camp. España · Art. V-H.3 RFEC (9.09.25):
  //   1º Campión España ano anterior · 2º Pts UCI (non aplica Cadete) · 3º Ranking RFEC CX · 4º Inscrición
  RFEC_CAMP:   ["Campión_España", "UCI", "Copa_España", "RFEC", "Inscrición"],

  // ── SERIE UCI / COPA DO MUNDO ───────────────────────────────────
  // Regulamento UCI exclusivo. RFEC e FGC como referencia subsidiaria.
  UCI_SERIE:   ["UCI", "RFEC", "FGC", "Inscrición"],
};

// ── NOTAS COMPLETAS POR NIVEL ────────────────────────────────────
const NOTAS_BASE = {

  FGC_LOCAL:
    "NORMATIVA FGC (act. 13/12/2025) · Art. III-H.4.1 (XCO) · III-I.5.1.1 (DH) · III-J.1.5 (XCM).\n" +
    "Orde na grella:\n" +
    "  1º Líder da Copa Galicia.\n" +
    "  2º Corredores con puntos UCI, de maior a menor (NON aplica categoría Infantil).\n" +
    "  3º Corredores sen puntos UCI → Ranking RFEC da disciplina.\n" +
    "  4º Clasificación xeral Copa Galicia (1ª proba: Copa do ano anterior).\n" +
    "  5º Resto por orde de inscrición.\n" +
    "XCO: chamada á TOTALIDADE dos corredores.\n" +
    "XCM: chamada aos 50 primeiros (Art. III-J.1.5).\n" +
    "DH manga clasificatoria: orde por categorías especificada en Art. III-I.5.1 (Infantil → Máster → Elite o últimos).",

  FGC_MINI:
    "NORMATIVA FGC (act. 13/12/2025) · Art. III-N.1.6 (Mini BTT-XC) · III-N.2.1 (Mini DH).\n" +
    "Carácter FORMATIVO. Sen puntos UCI nin RFEC.\n" +
    "Mini BTT-XC:\n" +
    "  · 1ª proba tempada → SORTEO DE CLUBS. O director deportivo reparte postos entre os seus corredores.\n" +
    "  · Resto de probas → ranking FGC MiniBTT, de maior a menor puntos.\n" +
    "Mini DH:\n" +
    "  · 1ª proba → ORDE INVERSA á data de inscrición.\n" +
    "  · Resto → ORDE INVERSA ao ranking FGC MiniDH (o mellor sae o ÚLTIMO na manga clasificatoria).\n" +
    "Licenzas dun día: sitúanse DESPOIS dos federados. Non obteñen puntos para o orde de saída.\n" +
    "Corredores non figurados no orde de saída: sitúanse inmediatamente despois dos nomeados.\n" +
    "Grellas SEPARADAS por sexo en cada categoría. Nº de corredores por fila: decide o organizador.\n" +
    "PENALIZACIÓN: cambio de dorsal ou placa = −20 puntos no ranking de orde de saída.",

  FGC_CAMP:
    "NORMATIVA FGC (act. 13/12/2025) · Art. III-F.5.1 (XCO) · III-F.5.2 (DH).\n" +
    "Campionato Galego:\n" +
    "  · O Campión/a de Galicia do ano anterior OCUPA A 1ª POSICIÓN, sempre que manteña a categoría.\n" +
    "  · O resto da grella realízase IGUAL QUE A COPA GALICIA da disciplina correspondente.\n" +
    "DH manga final: orde inversa por tempos da manga clasificatoria. O Campión/a Galicia do ano anterior sae o ÚLTIMO na manga clasificatoria dentro da súa categoría.",

  RFEC_PUNTCO:
    "NORMATIVA RFEC (BTT: 8.04.26 · CX: 9.09.25) · Clases H.5 / I.5 / K.5 / Ku.5 / L.5 / P.5 / G.5.\n" +
    "Probas nacionais propostas pola FGC no calendario RFEC. Aplícase a mesma normativa que a Copa de España da disciplina:\n" +
    "XCO (Art. IV-G.6): Categorías UCI → Ranking UCI → Copa España XCO → (1ª proba: Ranking RFEC) → Inscrición.\n" +
    "                   Categorías non-UCI → Copa España XCO → (1ª proba: Ranking RFEC) → Inscrición.\n" +
    "DH (Art. IV-H.5 manga clasif.): orde inversa por categorías: Infantil primeiro, Elite/Sub23 os últimos.\n" +
    "   Dentro de cada cat.: orde inversa á Copa España provisional → Ranking RFEC DHI → Pts UCI (menor a maior) → Campión España (o último).\n" +
    "CX (Art. V-J.9): 1º Pts UCI · 2º Clasificación Copa España CX (1ª proba: final ano anterior) · 3º Ranking RFEC CX · 4º Inscrición.\n" +
    "PumpTrack (Art. IV-Q.6): 1ª e 2ª manga: orde inscripción (sen puntos Copa) / inversa Copa España. 3ª manga: scratch inverso mellor tempo 1ª ou 2ª manga.",

  RFEC_COPA:
    "NORMATIVA RFEC (BTT: 8.04.26 · CX: 9.09.25) · Clases H.1 / I.1 / K.1 / Ku.1 / L.1 / P.1 / G.1.\n" +
    "XCO Copa España (Art. IV-G.6):\n" +
    "  Cat. UCI (Elite/Sub23/Júnior): Ranking UCI → Copa España XCO → (1ª proba: Ranking RFEC) → Inscrición.\n" +
    "  Cat. non-UCI (Cadete/Máster): Copa España XCO → (1ª proba: Ranking RFEC) → Inscrición.\n" +
    "DH Copa España (Art. IV-H.5) manga clasificatoria:\n" +
    "  Orde de categorías: Infantil M/F → M60 → M50 → M40 → M30 → Cadete → Mulleres Máster → Mulleres Cadete → Mulleres Júnior → Mulleres Elite/Sub23 → Júnior Masc. → Elite/Sub23 Masc.\n" +
    "  Dentro de cada cat. (agás Elite/Sub23/Júnior): orde inversa Copa España provisional → Ranking RFEC DHI → Pts UCI menor a maior → Campión España o ÚLTIMO.\n" +
    "  Elite/Sub23/Júnior: orde inversa Ranking RFEC individual. Manga final: inversa por tempos manga clasif.; Elite+Sub23+Júnior como cat. única scratch. Top-10M e Top-5F con protección UCI.\n" +
    "CX Copa España (Art. V-J.9): 1º Pts UCI · 2º Copa España CX (1ª proba: final ano anterior) · 3º Ranking RFEC CX · 4º Inscrición. Probas en días consecutivos: manter orde do 1º día.\n" +
    "XCM Copa España (Art. IV-J.1.5): Grupo1: Elite/Sub23 con pts UCI XCM → pts UCI XCO → resto E/S23. Grupo2: Masters 25 primeiros Copa España. Grupo3: resto a discreción (ábrese 15 min antes).\n" +
    "PumpTrack Copa España (Art. IV-Q.6): 1ª/2ª manga 1ª proba → orde inversa Copa España ano anterior. Resto probas → inversa clasificación provisional Copa. 3ª manga → scratch inverso mellor tempo.",

  RFEC_CAMP:
    "NORMATIVA RFEC (BTT: 8.04.26 · CX: 9.09.25) · Clase CN – Campionato Nacional de España.\n" +
    "XCO (Art. IV-F.5):\n" +
    "  1º Campión/a España ano anterior (mantén categoría).\n" +
    "  2º Ciclistas con puntos UCI, de maior a menor (NON aplica Cadete).\n" +
    "  3º Ranking RFEC BTT XCO vixente (se non existe: Copa España XCO ano en curso).\n" +
    "  4º Resto seleccionados polas FFAA.\n" +
    "  5º Orde de inscrición.\n" +
    "DH (Art. IV-F.6.2 – 13.11.25):\n" +
    "  Orde cat. manga clasif.: M Infantil → H Infantil → H M60 → H M50 → H M40 → H M30 → H Cadete → M Máster → M Cadete → M Júnior → M Elite/Sub23 → H Júnior → H Elite/Sub23.\n" +
    "  Dentro de cada cat.: Infantil por orde inversa inscrición. Resto: orde inversa Copa España → Ranking RFEC DHI → Pts UCI (menor a maior) → Campión España (o último, sae ao final).\n" +
    "  Manga final: mesmo orde cat. + orde inversa tempos manga clasif. Elite+Sub23+Júnior como cat. única. Protección UCI top-10M / top-5F: saen 'protexidos' inmediatamente despois dos anteriores. Intervalo mínimo 1 min entre os últimos 10 corredores.\n" +
    "CX (Art. V-H.3 – 9.09.25):\n" +
    "  1º Campión/a España ano anterior (mantén categoría).\n" +
    "  2º Pts UCI de maior a menor (NON aplica Cadete).\n" +
    "  3º Ranking RFEC individual de ciclocross.\n" +
    "  4º Orde de inscrición.\n" +
    "  Varias categorías xuntas: intervalo mínimo 15\" entre elas.\n" +
    "Categorías RFEC 2026 (I-B.1): Promesa 0-8a · Principiante 9-10a · Alevín 11-12a · Infantil 13-14a · Cadete 15-16a · Júnior 17-18a · Sub23 19-22a · Elite 23-75a · Máster30/40/50/60.",

  UCI_SERIE:
    "NORMATIVA UCI – Regulamento do Deporte Ciclista.\n" +
    "Clases C1, C2, C3 (XCO/DH/CX), S1, S2 (Short Track), MS (XCM Maratón Series).\n" +
    "Orde de saída conforme ao Regulamento UCI. RFEC e FGC actúan como referencia subsidiaria.\n" +
    "En probas UCI celebradas en España: consideración internacional para Elite, Sub23 e Júnior; nacional para o resto.",
};

// ── DATOS MINI BTT (Art. III-N.1.1 FGC) ────────────────────────
// Anos de nacemento para 2026 (RFEC I-B.1: Promesa 0-8a = 2018-2026, Principiante 9-10 = 2016-17, etc.)
const MINIBTT_CATEGORÍAS_INFO = {
  "Benjamín":     { anos: "2020 en diante", duración: "5 min",  roda: "Libre (non indica límite 29\")" },
  "Promesa":      { anos: "2018 e 2019",    duración: "10 min", roda: "Prohibida roda 29\"" },
  "Principiante": { anos: "2016 e 2017",    duración: "15 min", roda: "Prohibida roda 29\"" },
  "Alevín":       { anos: "2014 e 2015",    duración: "20 min", roda: "Prohibida roda 29\"" },
  "Infantil":     { anos: "2012 e 2013",    duración: "30 min", roda: "Permitida roda 29\"" },
};

// Orde de categorías na proba Mini BTT-XC (Art. III-N.1.1 FGC)
const MINIBTT_ORDE_CATEGORÍAS = [
  "Benjamín Fem.", "Benjamín B (nados 2021+)", "Benjamín A (nados 2020)",
  "Promesa Fem.", "Promesa B (nados 2019)", "Promesa A (nados 2018)",
  "Principiante Fem.", "Principiante B (nados 2017)", "Principiante A (nados 2016)",
  "Alevín Fem.", "Alevín Masc.",
  "Infantil Masc. + Infantil Fem.",
];

// Orde de categorías na proba Mini DH (Art. III-N.2.1 FGC)
const MINIDH_ORDE_CATEGORÍAS = [
  "Benjamín Fem.", "Benjamín Masc.",
  "Promesa Fem.", "Promesa Masc.",
  "Principiante Fem.", "Principiante Masc.",
  "Alevín Fem.", "Alevín Masc.",
  "Infantil Fem.", "Infantil Masc.",
];

// Táboa de puntos Mini BTT/DH para orde de saída (Art. I-Q FGC – columna MINIBTT/MINIDH)
const TABLA_PUNTOS_MINI = {
  1:600, 2:540, 3:500, 4:465, 5:425, 6:400, 7:377, 8:362, 9:347, 10:332,
  11:322, 12:312, 13:302, 14:292, 15:282, 16:272, 17:262, 18:252, 19:242, 20:232,
  21:224, 22:216, 23:208, 24:200, 25:192, 26:184, 27:176, 28:168, 29:160, 30:152,
  31:146, 32:140, 33:134, 34:128, 35:122, 36:116, 37:110, 38:104, 39:98,  40:92,
  41:88,  42:84,  43:80,  44:76,  45:72,  46:68,  47:64,  48:60,  49:56,  50:52,
  51:50,  52:48,  53:46,  54:44,  55:42,  56:40,  57:38,  58:36,  59:34,  60:32,
  61:30,  62:29,  63:28,  64:27,  65:26,  66:25,  67:24,  68:23,  69:22,  70:21,
  71:20,  72:19,  73:18,  74:17,  75:16,  76:15,  77:14,  78:13,  79:12,  80:11,
  81:10,  82:9,   83:8,   84:7,   85:6,   86:5,   87:4,   88:3,   89:2,   90:1,
};

// Táboa de puntos Copa España XCO (Art. IV-G.7 RFEC) – para referencia
const TABLA_PUNTOS_COPA_ESP_XCO = {
  1:200, 2:175, 3:155, 4:140, 5:128, 6:120, 7:112, 8:104, 9:96, 10:88,
  11:80, 12:76, 13:72, 14:68, 15:64, 16:60, 17:56, 18:52, 19:48, 20:44,
  21:40, 22:36, 23:32, 24:28, 25:24,
  "26-30":22, "31-35":20, "36-40":18, "41-45":16, "46-50":14,
};

// Alias de columnas para auto-detección
const ALIAS = {
  dorsal:     ["dorsal","num_bib","bib","número","numero","nº","number"],
  nome:       ["nome","name","nombre","apellidos","corredor","atleta","rider","nombre_apellidos"],
  apelidos:   ["apelidos","apellidos","surname","lastname"],
  equipo:     ["equipo","team","clube","club","escuadra","team_name"],
  uci_id:     ["uci_id","uciid","uci id","codigo_uci","uci","uci_code"],
  licenza:    ["licenza","licencia","licence","license","lic","num_licencia"],
  categoría:  ["categoria","categoría","category","cat","grupo"],
  pts_uci:    ["pts_uci","puntos_uci","uci_points","ucipoints","puntos uci","uci pts","rankinguci"],
  pts_rfec:   ["pts_rfec","puntos_rfec","rfec_points","puntos rfec","rfec pts","rankingrfec"],
  pts_fgc:    ["pts_fgc","puntos_fgc","fgc_points","puntos fgc","fgc pts","rankingfgc"],
  inscricion: ["inscricion","inscripcion","inscription","orde","order","nº inscricion","num_orden"],
};

// ═══════════════════════════════════════════════════════════════════
// PERSISTENCIA LOCAL (simula base de datos offline)
// ═══════════════════════════════════════════════════════════════════

function gardaEstado(estado) {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(estado)); } catch(e) {}
}

function cargaEstado() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch(e) { return null; }
}

const ESTADO_INICIAL = {
  rankings: {},      // { "UCI_XCO_Elite_2026-06": [{uci_id, nome, pts}, ...] }
  normativas: {},    // { "XCO_FGC_LOCAL": { caixonsPorFila, xerarquía, notas, data } }
  metadatos: {
    ultima_actualizacion: null,
    version: VERSIÓN,
  }
};

// ═══════════════════════════════════════════════════════════════════
// UTILIDADES DE FICHEIRO
// ═══════════════════════════════════════════════════════════════════

function lerFicheiroExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
        resolve(rows);
      } catch(err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function lerFicheiroCSV(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const text = e.target.result;
        const sep = text.includes(";") ? ";" : ",";
        const lines = text.trim().split(/\r?\n/);
        if (lines.length < 2) { resolve([]); return; }
        const headers = lines[0].split(sep).map(h => h.trim().replace(/^["']|["']$/g,""));
        const rows = lines.slice(1).map(line => {
          const vals = line.split(sep).map(v => v.trim().replace(/^["']|["']$/g,""));
          const obj = {};
          headers.forEach((h,i) => { obj[h] = vals[i] ?? ""; });
          return obj;
        });
        resolve(rows);
      } catch(err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsText(file, "UTF-8");
  });
}

async function lerFicheiro(file) {
  if (!file) return [];
  const ext = file.name.split(".").pop().toLowerCase();
  if (ext === "xlsx" || ext === "xls") return lerFicheiroExcel(file);
  return lerFicheiroCSV(file);
}

function detectarCol(headers, tipo) {
  const al = ALIAS[tipo] || [];
  for (const h of headers) {
    const hl = h.toLowerCase().replace(/[\s_-]+/g,"_");
    if (al.some(a => hl === a || hl.includes(a))) return h;
  }
  return null;
}

function mapearRanking(rows, fonte) {
  if (!rows.length) return [];
  const headers = Object.keys(rows[0]);
  const colNome     = detectarCol(headers,"nome");
  const colUciId    = detectarCol(headers,"uci_id");
  const colLicenza  = detectarCol(headers,"licenza");
  const colPts      = fonte==="UCI" ? detectarCol(headers,"pts_uci")
                    : fonte==="RFEC" ? detectarCol(headers,"pts_rfec")
                    : detectarCol(headers,"pts_fgc");
  const colCat      = detectarCol(headers,"categoría");
  const colEquipo   = detectarCol(headers,"equipo");

  return rows.map((r,i) => ({
    pos: i+1,
    nome:    (colNome    ? r[colNome]    : "") || "–",
    uci_id:  (colUciId   ? r[colUciId]   : "") || "",
    licenza: (colLicenza ? r[colLicenza] : "") || "",
    pts:     Number(colPts ? r[colPts] : 0) || 0,
    categoría: (colCat   ? r[colCat]   : "") || "",
    equipo:  (colEquipo  ? r[colEquipo] : "") || "",
  })).filter(r => r.pts > 0 || r.uci_id || r.licenza);
}

function mapearInscritos(rows) {
  if (!rows.length) return [];
  const headers = Object.keys(rows[0]);
  const cols = {};
  for (const tipo of Object.keys(ALIAS)) cols[tipo] = detectarCol(headers,tipo);

  return rows.map((r,i) => ({
    id: i,
    dorsal:    String(r[cols.dorsal]   || i+1),
    nome:      r[cols.nome]    || r[cols.apelidos] || "Descoñecido",
    equipo:    r[cols.equipo]  || "–",
    categoría: r[cols.categoría] || "",
    uci_id:    String(r[cols.uci_id]  || ""),
    licenza:   String(r[cols.licenza] || ""),
    inscricion: Number(r[cols.inscricion]) || (i+1),
    pts_uci:  0, pts_rfec: 0, pts_fgc: 0,
    dns: false, bloqueado: false, pos_bloq: null,
    alertas: [],
  }));
}

// ═══════════════════════════════════════════════════════════════════
// MOTOR EN FERVENZA
// ═══════════════════════════════════════════════════════════════════

function cruzarPuntos(corredores, rankings, categoríaFiltro) {
  return corredores.map(c => {
    const busca = (lista) => {
      if (!lista) return null;
      const filt = categoríaFiltro
        ? lista.filter(r => !r.categoría || r.categoría.toLowerCase().includes(categoríaFiltro.toLowerCase()))
        : lista;
      if (c.uci_id) {
        const m = filt.find(r => r.uci_id && r.uci_id.trim()===c.uci_id.trim());
        if (m) return m;
      }
      if (c.licenza) {
        const m = filt.find(r => r.licenza && r.licenza.trim()===c.licenza.trim());
        if (m) return m;
      }
      return null;
    };

    const mUCI  = busca(rankings.UCI);
    const mRFEC = busca(rankings.RFEC);
    const mFGC  = busca(rankings.FGC);

    const alertas = [];
    if (!mUCI && !mRFEC && !mFGC) {
      if (!c.uci_id && !c.licenza) alertas.push("Sen UCI ID nin licenza");
      else alertas.push("Non localizado en ningún ránking");
    }

    return {
      ...c,
      pts_uci:  mUCI  ? mUCI.pts  : 0,
      pts_rfec: mRFEC ? mRFEC.pts : 0,
      pts_fgc:  mFGC  ? mFGC.pts  : 0,
      alertas,
    };
  });
}

function ordenarFervenza(corredores, xerarquía) {
  const libres    = corredores.filter(c => !c.dns && !c.bloqueado);
  const bloqueados = corredores.filter(c => !c.dns && c.bloqueado);

  libres.sort((a,b) => {
    for (const regra of xerarquía) {
      let da=0, db=0;
      if (regra==="UCI")       { da=a.pts_uci;  db=b.pts_uci; }
      else if (regra==="RFEC") { da=a.pts_rfec; db=b.pts_rfec; }
      else if (regra==="FGC")  { da=a.pts_fgc;  db=b.pts_fgc; }
      else if (regra==="Inscrición") { da=-a.inscricion; db=-b.inscricion; }
      if (db!==da) return db-da;
    }
    return a.inscricion-b.inscricion;
  });

  const resultado = [...libres];
  const bOrd = [...bloqueados].sort((a,b)=>(a.pos_bloq||999)-(b.pos_bloq||999));
  for (const b of bOrd) {
    const pos = Math.min((b.pos_bloq||1)-1, resultado.length);
    resultado.splice(pos,0,b);
  }
  return resultado;
}

function construirMatriz(lista, n) {
  const filas=[];
  for(let i=0;i<lista.length;i+=n) filas.push(lista.slice(i,i+n));
  return filas;
}

// ═══════════════════════════════════════════════════════════════════
// EXPORTACIÓN LOG
// ═══════════════════════════════════════════════════════════════════

function xerarLog(grella, cfg) {
  const ts  = new Date().toISOString();
  const sep = "═".repeat(70);
  const lin = "─".repeat(70);
  let txt = `${sep}\nGRELLA OFICIAL DE SAÍDA – FEDERACIÓN GALEGA DE CICLISMO\n${sep}\n`;
  txt += `Modalidade : ${cfg.modalidade}\n`;
  txt += `Nivel      : ${cfg.nivel}\n`;
  txt += `Categoría  : ${cfg.categoría||"Todas"}\n`;
  txt += `Caix./Fila : ${cfg.porFila}\n`;
  txt += `Xerarquía  : ${cfg.xerarquía.join(" › ")}\n`;
  txt += `Xerado     : ${ts}\n`;
  txt += `Rankings   : ${Object.keys(cfg.rankingsUsados||{}).join(", ")||"–"}\n`;
  txt += `${lin}\n`;
  txt += `Pos  Dor   Nome                            Equipo                    UCI    RFEC   FGC\n`;
  txt += `${lin}\n`;
  grella.forEach((c,i) => {
    txt += `${String(i+1).padStart(3)}  ${String("D"+c.dorsal).padStart(5)}  `;
    txt += `${c.nome.padEnd(30)}  ${c.equipo.padEnd(24)}  `;
    txt += `${String(c.pts_uci).padStart(5)}  ${String(c.pts_rfec).padStart(5)}  ${String(c.pts_fgc).padStart(5)}`;
    if(c.bloqueado) txt += "  🔒";
    txt += "\n";
  });
  txt += `${sep}\n[FIN DO LOG DE AUDITORÍA – ${ts}]\n`;
  return txt;
}

function descargarTexto(contido, nome) {
  const a = document.createElement("a");
  a.href = "data:text/plain;charset=utf-8,"+encodeURIComponent(contido);
  a.download = nome;
  a.click();
}

// ═══════════════════════════════════════════════════════════════════
// COMPOÑENTES UI BASE
// ═══════════════════════════════════════════════════════════════════

const C = {
  bg:       "#0d0d1f",
  panel:    "#12122a",
  panel2:   "#1a1a38",
  border:   "#2a2a50",
  gold:     "#f0e040",
  goldDim:  "#c8bc30",
  text:     "#e0e0f0",
  muted:    "#8080a0",
  danger:   "#e74c3c",
  warn:     "#e67e22",
  ok:       "#27ae60",
  blue:     "#3498db",
};

const css = {
  app: {
    minHeight:"100vh", background:C.bg,
    fontFamily:"Verdana,Geneva,sans-serif", color:C.text,
    fontSize:14, lineHeight:1.5,
  },
  header: {
    background:`linear-gradient(135deg,${C.panel} 0%,${C.panel2} 100%)`,
    borderBottom:`3px solid ${C.gold}`,
    padding:"1rem 1.5rem",
    display:"flex", alignItems:"center", justifyContent:"space-between",
    flexWrap:"wrap", gap:10,
  },
  card: {
    background:C.panel, border:`1px solid ${C.border}`,
    borderRadius:10, padding:"1.25rem",
    marginBottom:"1rem",
  },
  label: {
    fontSize:10, color:C.muted, textTransform:"uppercase",
    letterSpacing:1.5, marginBottom:6, display:"block",
  },
  input: {
    background:C.panel2, border:`1px solid ${C.border}`,
    borderRadius:6, padding:"8px 12px", color:C.text,
    fontFamily:"Verdana,sans-serif", fontSize:13, width:"100%",
    boxSizing:"border-box",
  },
  select: {
    background:C.panel2, border:`1px solid ${C.border}`,
    borderRadius:6, padding:"8px 12px", color:C.text,
    fontFamily:"Verdana,sans-serif", fontSize:13, width:"100%",
    boxSizing:"border-box", cursor:"pointer",
  },
};

function Btn({children,onClick,variant="primary",disabled=false,small=false,style={}}) {
  const v = {
    primary:   {background:C.gold,   color:"#0d0d1f", border:"none"},
    secondary: {background:C.panel2, color:C.text,    border:`1px solid ${C.border}`},
    danger:    {background:C.danger, color:"#fff",    border:"none"},
    ghost:     {background:"transparent", color:C.gold, border:`1px solid ${C.gold}`},
    ok:        {background:C.ok,     color:"#fff",    border:"none"},
  };
  return (
    <button onClick={onClick} disabled={disabled} style={{
      ...v[variant],
      fontFamily:"Verdana,sans-serif", fontWeight:700,
      fontSize: small?11:13, borderRadius:6,
      padding: small?"5px 10px":"9px 18px",
      cursor:disabled?"not-allowed":"pointer",
      opacity:disabled?0.4:1, transition:"all .15s",
      ...style
    }}>{children}</button>
  );
}

function Tag({children, color=C.muted}) {
  return (
    <span style={{
      display:"inline-block", padding:"2px 9px", borderRadius:20,
      background:color+"22", color, border:`1px solid ${color}44`,
      fontSize:11, fontWeight:700,
    }}>{children}</span>
  );
}

function Alert({tipo="warn",children}) {
  const cols = {warn:{bg:"#2a1a00",bord:C.warn,txt:C.warn}, err:{bg:"#2a0000",bord:C.danger,txt:C.danger}, ok:{bg:"#002a10",bord:C.ok,txt:C.ok}};
  const c = cols[tipo]||cols.warn;
  return (
    <div style={{background:c.bg,border:`1px solid ${c.bord}`,borderRadius:8,padding:"10px 14px",marginBottom:10}}>
      <span style={{color:c.txt,fontSize:12}}>{children}</span>
    </div>
  );
}

function DropZone({onFile, accept=".xlsx,.xls,.csv", label="Arrastra o ficheiro ou preme para explorar"}) {
  const ref = useRef();
  const [over,setOver] = useState(false);
  return (
    <div
      onDragOver={e=>{e.preventDefault();setOver(true);}}
      onDragLeave={()=>setOver(false)}
      onDrop={e=>{e.preventDefault();setOver(false);onFile(e.dataTransfer.files[0]);}}
      onClick={()=>ref.current.click()}
      style={{
        border:`2px dashed ${over?C.gold:C.border}`, borderRadius:10,
        padding:"2rem", textAlign:"center", cursor:"pointer",
        background: over?C.panel2:C.bg, transition:"all .2s",
      }}>
      <div style={{fontSize:32,marginBottom:8}}>📂</div>
      <div style={{color:C.gold,fontWeight:700,marginBottom:4}}>{label}</div>
      <div style={{fontSize:11,color:C.muted}}>Formatos: {accept}</div>
      <input ref={ref} type="file" accept={accept} style={{display:"none"}}
        onChange={e=>onFile(e.target.files[0])} />
    </div>
  );
}

function Stepper({pasos, actual}) {
  return (
    <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center",marginBottom:"1rem"}}>
      {pasos.map((p,i) => (
        <div key={i} style={{display:"flex",alignItems:"center",gap:8}}>
          <div style={{
            width:28,height:28,borderRadius:"50%",
            background: i<actual?C.gold : i===actual?C.gold+"33":"transparent",
            border:`2px solid ${i<=actual?C.gold:C.border}`,
            color: i<actual?"#0d0d1f":i===actual?C.gold:C.muted,
            display:"flex",alignItems:"center",justifyContent:"center",
            fontSize:12,fontWeight:700,flexShrink:0,
          }}>
            {i<actual?"✓":i+1}
          </div>
          <span style={{fontSize:12,color:i<=actual?C.gold:C.muted,fontWeight:i===actual?700:400}}>
            {p}
          </span>
          {i<pasos.length-1 && <span style={{color:C.border,fontSize:16}}>›</span>}
        </div>
      ))}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// MÓDULO ADMIN – Xestión de Rankings
// ═══════════════════════════════════════════════════════════════════

function PanelRankings({estado,setEstado}) {
  const [fonte,setFonte]       = useState("FGC");
  const [modalidade,setMod]    = useState("XCO");
  const [categoría,setCat]     = useState("Elite");
  const [tempada,setTemp]      = useState(new Date().getFullYear().toString());
  const [msg,setMsg]           = useState(null);
  const [preview,setPreview]   = useState([]);

  const chave = `${fonte}_${modalidade}_${categoría}_${tempada}`;
  const rankActual = estado.rankings[chave];

  const onFich = async (file) => {
    if(!file) return;
    try {
      const rows = await lerFicheiro(file);
      const mapeados = mapearRanking(rows, fonte);
      setPreview(mapeados.slice(0,8));
      const novo = {...estado, rankings:{...estado.rankings,[chave]:mapeados},
        metadatos:{...estado.metadatos,ultima_actualizacion:new Date().toISOString()}};
      setEstado(novo); gardaEstado(novo);
      setMsg({tipo:"ok",txt:`✓ ${mapeados.length} corredores cargados en ${chave}`});
    } catch(e) {
      setMsg({tipo:"err",txt:"Erro ao ler o ficheiro: "+e.message});
    }
  };

  const eliminar = () => {
    const {[chave]:_,...resto} = estado.rankings;
    const novo = {...estado,rankings:resto};
    setEstado(novo); gardaEstado(novo);
    setPreview([]); setMsg({tipo:"ok",txt:`Ranking ${chave} eliminado`});
  };

  const ranksCargados = Object.keys(estado.rankings);

  return (
    <div>
      <div style={css.card}>
        <span style={css.label}>Cargar / Actualizar Ranking</span>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:10,marginBottom:16}}>
          <div>
            <span style={css.label}>Fonte</span>
            <select value={fonte} onChange={e=>setFonte(e.target.value)} style={css.select}>
              {["FGC","RFEC","UCI"].map(f=><option key={f}>{f}</option>)}
            </select>
          </div>
          <div>
            <span style={css.label}>Modalidade</span>
            <select value={modalidade} onChange={e=>setMod(e.target.value)} style={css.select}>
              <option value="TODAS">– Todas as modalidades –</option>
              {MODALIDADES.map(m=>(
                <option key={m.id} value={m.id}>
                  {m.mini?"🟡 ":""}{m.id}
                </option>
              ))}
            </select>
          </div>
          <div>
            <span style={css.label}>Categoría</span>
            <select value={categoría} onChange={e=>setCat(e.target.value)} style={css.select}>
              <option value="TODAS">– Todas as categorías –</option>
              {getCategorías(modalidade==="TODAS"?"XCO":modalidade).map(c=><option key={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <span style={css.label}>Tempada</span>
            <input value={tempada} onChange={e=>setTemp(e.target.value)} style={css.input} placeholder="2026"/>
          </div>
        </div>
        {MODALIDADES.find(m=>m.id===modalidade)?.mini && fonte!=="FGC" && (
          <Alert tipo="warn">
            ⚠ As categorías Mini BTT/DH só teñen ránking FGC. Non existen rankings RFEC nin UCI para estas categorías de base.
          </Alert>
        )}

        <div style={{marginBottom:8,padding:"8px 12px",background:C.panel2,borderRadius:6,
          fontSize:12,color:C.muted}}>
          Clave: <strong style={{color:C.gold}}>{chave}</strong>
          {rankActual && <span style={{marginLeft:10,color:C.ok}}>✓ {rankActual.length} rexistros</span>}
        </div>

        <DropZone onFile={onFich} label={`Subir ranking ${fonte} – ${modalidade} – ${categoría}`}/>

        {msg && <div style={{marginTop:10}}>
          <Alert tipo={msg.tipo==="ok"?"ok":"err"}>{msg.txt}</Alert>
        </div>}

        {preview.length>0 && (
          <div style={{marginTop:12}}>
            <span style={css.label}>Vista previa (primeiros 8)</span>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead>
                  <tr style={{background:C.panel2}}>
                    {["Pos","Nome","UCI ID","Licenza","Puntos","Categoría"].map(h=>(
                      <th key={h} style={{padding:"6px 10px",textAlign:"left",color:C.muted,
                        borderBottom:`1px solid ${C.border}`,fontWeight:700}}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {preview.map((r,i)=>(
                    <tr key={i} style={{background:i%2===0?C.bg:C.panel}}>
                      <td style={{padding:"5px 10px",color:C.gold}}>{r.pos}</td>
                      <td style={{padding:"5px 10px"}}>{r.nome}</td>
                      <td style={{padding:"5px 10px",color:C.muted,fontSize:11}}>{r.uci_id||"–"}</td>
                      <td style={{padding:"5px 10px",color:C.muted,fontSize:11}}>{r.licenza||"–"}</td>
                      <td style={{padding:"5px 10px",fontWeight:700,color:C.gold}}>{r.pts}</td>
                      <td style={{padding:"5px 10px",color:C.muted}}>{r.categoría||"–"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>

      {/* Rankings cargados */}
      <div style={css.card}>
        <span style={css.label}>Rankings dispoñibles no sistema ({ranksCargados.length})</span>
        {ranksCargados.length===0
          ? <div style={{color:C.muted,fontSize:13}}>Non hai rankings cargados aínda.</div>
          : <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
              {ranksCargados.map(k=>(
                <div key={k} style={{
                  background:C.panel2,border:`1px solid ${C.border}`,borderRadius:8,
                  padding:"8px 14px",display:"flex",alignItems:"center",gap:10
                }}>
                  <span style={{fontSize:12,color:C.gold,fontWeight:700}}>{k}</span>
                  <span style={{fontSize:11,color:C.muted}}>{estado.rankings[k].length} corr.</span>
                  {k===chave && (
                    <button onClick={eliminar} style={{background:"none",border:"none",color:C.danger,cursor:"pointer",fontSize:14}}>✕</button>
                  )}
                </div>
              ))}
            </div>
        }
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// MÓDULO ADMIN – Xestión de Normativas
// ═══════════════════════════════════════════════════════════════════

function PanelNormativas({estado,setEstado}) {
  const [modalidade,setMod]  = useState("XCO");
  const [nivel,setNivel]     = useState("FGC_LOCAL");
  const [caixons,setCaixons] = useState(8);
  const [xerarquía,setXer]   = useState([...XERARQUÍAS_BASE.FGC_LOCAL]);
  const [notas,setNotas]     = useState("");
  const [msg,setMsg]         = useState(null);

  const chave = `${modalidade}_${nivel}`;
  const normActual = estado.normativas[chave];

  useEffect(()=>{
    if(normActual){
      setCaixons(normActual.caixonsPorFila||8);
      setXer(normActual.xerarquía||[...(XERARQUÍAS_BASE[nivel]||["FGC","Inscrición"])]);
      setNotas(normActual.notas||"");
    } else {
      setXer([...(XERARQUÍAS_BASE[nivel]||["FGC","Inscrición"])]);
      setNotas(NOTAS_BASE[nivel]||"");
    }
  },[chave]);

  const gardar = () => {
    const nova = {...estado, normativas:{...estado.normativas,[chave]:{
      caixonsPorFila:caixons, xerarquía, notas,
      data: new Date().toISOString(),
    }}};
    setEstado(nova); gardaEstado(nova);
    setMsg({tipo:"ok",txt:`✓ Normativa ${chave} gardada`});
  };

  const moverRegra = (i,dir) => {
    const nx=[...xerarquía];
    const j=i+dir;
    if(j<0||j>=nx.length) return;
    [nx[i],nx[j]]=[nx[j],nx[i]];
    setXer(nx);
  };

  const normsCargadas = Object.keys(estado.normativas);

  return (
    <div>
      <div style={css.card}>
        <span style={css.label}>Configurar Normativa</span>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:10,marginBottom:16}}>
          <div>
            <span style={css.label}>Modalidade</span>
            <select value={modalidade} onChange={e=>setMod(e.target.value)} style={css.select}>
              <option value="TODAS">– Todas as modalidades –</option>
              {MODALIDADES.map(m=>(
                <option key={m.id} value={m.id}>
                  {m.mini ? "🟡 " : ""}{m.id} – {m.label.split("–")[1]?.trim()||m.label}
                </option>
              ))}
            </select>
          </div>
          <div>
            <span style={css.label}>Nivel de competición</span>
            <select value={nivel} onChange={e=>{setNivel(e.target.value);setXer([...(XERARQUÍAS_BASE[e.target.value]||["FGC","Inscrición"])]);setNotas(NOTAS_BASE[e.target.value]||"");}} style={css.select}>
              {NIVEIS_COMPETICIÓN.map(n=><option key={n.id} value={n.id}>{n.label}</option>)}
            </select>
          </div>
          <div>
            <span style={css.label}>Caixóns por fila</span>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              <button onClick={()=>setCaixons(p=>Math.max(1,p-1))}
                style={{width:32,height:32,borderRadius:6,background:C.panel2,border:`1px solid ${C.border}`,color:C.text,cursor:"pointer",fontSize:16}}>−</button>
              <span style={{width:36,textAlign:"center",fontWeight:700,color:C.gold,fontSize:18}}>{caixons}</span>
              <button onClick={()=>setCaixons(p=>Math.min(20,p+1))}
                style={{width:32,height:32,borderRadius:6,background:C.panel2,border:`1px solid ${C.border}`,color:C.text,cursor:"pointer",fontSize:16}}>+</button>
            </div>
          </div>
        </div>

        <div style={{marginBottom:16}}>
          <span style={css.label}>Xerarquía de desempate (arrastra para reordenar)</span>
          <div style={{display:"flex",flexDirection:"column",gap:6}}>
            {xerarquía.map((r,i)=>(
              <div key={r} style={{
                display:"flex",alignItems:"center",gap:10,
                background:C.panel2,border:`1px solid ${C.border}`,
                borderRadius:8,padding:"8px 14px",
              }}>
                <span style={{width:24,height:24,borderRadius:"50%",
                  background:i===0?C.gold:C.panel,color:i===0?"#0d0d1f":C.muted,
                  display:"flex",alignItems:"center",justifyContent:"center",
                  fontWeight:700,fontSize:12,flexShrink:0}}>{i+1}</span>
                <span style={{flex:1,fontWeight:i===0?700:400,color:i===0?C.gold:C.text}}>{r}</span>
                <button onClick={()=>moverRegra(i,-1)} disabled={i===0}
                  style={{background:"none",border:"none",color:i===0?C.border:C.muted,cursor:i===0?"default":"pointer",fontSize:16}}>▲</button>
                <button onClick={()=>moverRegra(i,1)} disabled={i===xerarquía.length-1}
                  style={{background:"none",border:"none",color:i===xerarquía.length-1?C.border:C.muted,cursor:i===xerarquía.length-1?"default":"pointer",fontSize:16}}>▼</button>
              </div>
            ))}
          </div>
        </div>

        <div style={{marginBottom:16}}>
          <span style={css.label}>Notas / Observacións regulamentarias</span>
          <textarea value={notas} onChange={e=>setNotas(e.target.value)}
            style={{...css.input,height:80,resize:"vertical"}}
            placeholder="Ex: En caso de empate en puntos FGC aplicarase sorteo público..."/>
        </div>

        <div style={{display:"flex",gap:10,alignItems:"center"}}>
          <Btn onClick={gardar}>Gardar normativa</Btn>
          {normActual && <Tag color={C.ok}>Gardada {new Date(normActual.data).toLocaleDateString("gl-ES")}</Tag>}
        </div>
        {msg && <div style={{marginTop:10}}><Alert tipo={msg.tipo}>{msg.txt}</Alert></div>}
      </div>

      <div style={css.card}>
        <span style={css.label}>Normativas no sistema ({normsCargadas.length})</span>
        {normsCargadas.length===0
          ? <div style={{color:C.muted,fontSize:13}}>Non hai normativas gardadas.</div>
          : <div style={{display:"grid",gap:8}}>
              {normsCargadas.map(k=>{
                const n=estado.normativas[k];
                return (
                  <div key={k} style={{background:C.panel2,border:`1px solid ${C.border}`,
                    borderRadius:8,padding:"10px 14px",display:"flex",
                    alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
                    <div>
                      <span style={{color:C.gold,fontWeight:700,fontSize:13}}>{k}</span>
                      <span style={{marginLeft:10,fontSize:11,color:C.muted}}>
                        {n.caixonsPorFila} caix./fila · {n.xerarquía.join(" › ")}
                      </span>
                    </div>
                    <Tag color={C.blue}>{new Date(n.data).toLocaleDateString("gl-ES")}</Tag>
                  </div>
                );
              })}
            </div>
        }
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════
// MÓDULO ÁRBITRO – Xeración de Grella
// ═══════════════════════════════════════════════════════════════════

function PanelArbitro({estado}) {
  const [paso,setPaso]         = useState(0);
  const [modalidade,setMod]    = useState("");
  const [nivel,setNivel]       = useState("");
  const [categoría,setCat]     = useState("");
  const [tempada,setTemp]      = useState(new Date().getFullYear().toString());
  const [corredores,setCorr]   = useState([]);
  const [fichNome,setFichNome] = useState("");
  const [grella,setGrella]     = useState([]);
  const [matriz,setMatriz]     = useState([]);
  const [drag,setDrag]         = useState(null);
  const [dragOver,setDragOver] = useState(null);
  const [vistaTab,setVistaTab] = useState("visual"); // visual | lista | log
  const [alertas,setAlertas]   = useState([]);
  const [xerado,setXerado]     = useState(false);

  const chaveNorm = `${modalidade}_${nivel}`;
  const normativa = estado.normativas[chaveNorm];
  const porFila   = normativa?.caixonsPorFila || 8;
  const xerarquía = normativa?.xerarquía || (nivel ? XERARQUÍAS_BASE[nivel] : []);

  // Rankings dispoñibles para este contexto
  const getRanking = (fonte) => {
    // Se modalidade ou categoría é "TODAS", busca calquera ranking dispoñible para esa fonte e tempada
    if (modalidade === "TODAS" || !modalidade) {
      const prefix = `${fonte}_`;
      const chaves = Object.keys(estado.rankings).filter(k =>
        k.startsWith(prefix) && k.endsWith(`_${tempada}`)
      );
      if (!chaves.length) return null;
      // Devolve array combinado de todos
      return chaves.flatMap(k => estado.rankings[k]);
    }
    if (categoría === "" || categoría === "TODAS" || !categoría) {
      // Busca todos os rankings desta modalidade e tempada independentemente da categoría
      const prefix = `${fonte}_${modalidade}_`;
      const chaves = Object.keys(estado.rankings).filter(k =>
        k.startsWith(prefix) && k.endsWith(`_${tempada}`)
      );
      if (!chaves.length) return null;
      return chaves.flatMap(k => estado.rankings[k]);
    }
    const chave = `${fonte}_${modalidade}_${categoría}_${tempada}`;
    return estado.rankings[chave] || null;
  };

  const rankingsActuais = {
    UCI:  getRanking("UCI"),
    RFEC: getRanking("RFEC"),
    FGC:  getRanking("FGC"),
  };

  const rankingsDisp = Object.entries(rankingsActuais)
    .filter(([,v])=>v).map(([k])=>k);

  const onFicheiroInscritos = async (file) => {
    if(!file) return;
    setFichNome(file.name);
    try {
      const rows = await lerFicheiro(file);
      const mapped = mapearInscritos(rows);
      setCorr(mapped);
      setAlertas([]);
      setGrella([]);
      setMatriz([]);
      setXerado(false);
      setPaso(2);
    } catch(e) {
      setAlertas([{tipo:"err",txt:"Erro ao ler inscritos: "+e.message}]);
    }
  };

  const xerarGrella = () => {
    const cruzados = cruzarPuntos(corredores, rankingsActuais, categoría);
    setCorr(cruzados);
    const al = cruzados
      .filter(c=>c.alertas.length>0)
      .map(c=>({tipo:"warn",txt:`D${c.dorsal} ${c.nome}: ${c.alertas.join(", ")}`}));
    setAlertas(al);
    const ord = ordenarFervenza(cruzados, xerarquía);
    const mat = construirMatriz(ord, porFila);
    setGrella(ord);
    setMatriz(mat);
    setXerado(true);
    setPaso(3);
  };

  const toggleDNS = (id) => {
    const novo = corredores.map(c=>c.id===id?{...c,dns:!c.dns}:c);
    setCorr(novo);
    const ord = ordenarFervenza(novo, xerarquía);
    setGrella(ord);
    setMatriz(construirMatriz(ord,porFila));
  };

  const toggleBloq = (id) => {
    const c = corredores.find(x=>x.id===id);
    if(!c) return;
    if(c.bloqueado){
      const novo=corredores.map(x=>x.id===id?{...x,bloqueado:false,pos_bloq:null}:x);
      setCorr(novo);
      const ord=ordenarFervenza(novo,xerarquía);
      setGrella(ord); setMatriz(construirMatriz(ord,porFila));
    } else {
      const posActual = grella.findIndex(x=>x.id===id)+1;
      const pos = window.prompt(`Fixar "${c.nome}" na posición número:`,String(posActual));
      if(!pos) return;
      const novo=corredores.map(x=>x.id===id?{...x,bloqueado:true,pos_bloq:parseInt(pos)}:x);
      setCorr(novo);
      const ord=ordenarFervenza(novo,xerarquía);
      setGrella(ord); setMatriz(construirMatriz(ord,porFila));
    }
  };

  const onDragEnd = () => {
    if(drag===null||dragOver===null||drag===dragOver){setDrag(null);setDragOver(null);return;}
    const nova=[...grella];
    const [m]=nova.splice(drag,1);
    nova.splice(dragOver,0,m);
    setGrella(nova); setMatriz(construirMatriz(nova,porFila));
    setDrag(null); setDragOver(null);
  };

  const logTexto = xerado ? xerarLog(grella,{
    modalidade,nivel,categoría,porFila,xerarquía,
    rankingsUsados: Object.fromEntries(rankingsDisp.map(k=>[k,true]))
  }) : "";

  const pasos = ["Configuración","Inscritos","Grella"];

  return (
    <div>
      <Stepper pasos={pasos} actual={paso}/>

      {/* PASO 0: Config */}
      {paso===0 && (
        <div style={css.card}>
          <span style={css.label}>Paso 1 · Configuración da proba</span>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:12,marginBottom:16}}>
            <div>
              <span style={css.label}>Modalidade</span>
              <select value={modalidade} onChange={e=>setMod(e.target.value)} style={css.select}>
                <option value="">– Escolle –</option>
                <option value="TODAS">⭐ Todas as modalidades</option>
                {MODALIDADES.map(m=><option key={m.id} value={m.id}>{m.label}</option>)}
              </select>
            </div>
            <div>
              <span style={css.label}>Nivel de competición</span>
              <select value={nivel} onChange={e=>setNivel(e.target.value)} style={css.select}>
                <option value="">– Escolle –</option>
                {NIVEIS_COMPETICIÓN.map(n=><option key={n.id} value={n.id}>{n.label}</option>)}
              </select>
            </div>
            <div>
              <span style={css.label}>Categoría</span>
              <select value={categoría} onChange={e=>setCat(e.target.value)} style={css.select}>
                <option value="">⭐ Todas as categorías</option>
                {(modalidade && modalidade!=="TODAS"
                  ? getCategorías(modalidade)
                  : CATEGORÍAS
                ).map(c=><option key={c}>{c}</option>)}
              </select>
              {MODALIDADES.find(m=>m.id===modalidade)?.mini && (
                <div style={{fontSize:10,color:C.warn,marginTop:4}}>
                  ⚠ Mini BTT/DH: categorías de base FGC (idades 6-15 anos). Sen puntos UCI/RFEC.
                </div>
              )}
            </div>
            <div>
              <span style={css.label}>Tempada</span>
              <input value={tempada} onChange={e=>setTemp(e.target.value)} style={css.input}/>
            </div>
          </div>

          {/* Info normativa */}
          {modalidade && nivel && (
            <div style={{
              background:C.panel2,border:`1px solid ${normativa?C.ok:C.warn}`,
              borderRadius:8,padding:"12px 16px",marginBottom:16
            }}>
              {normativa
                ? <>
                    <div style={{color:C.ok,fontWeight:700,marginBottom:6}}>✓ Normativa cargada: {chaveNorm}</div>
                    <div style={{fontSize:12,color:C.muted}}>
                      Xerarquía: <strong style={{color:C.gold}}>{normativa.xerarquía.join(" › ")}</strong>
                      {" · "}{normativa.caixonsPorFila} caixóns/fila
                    </div>
                    {normativa.notas && <div style={{fontSize:11,color:C.muted,marginTop:4}}>{normativa.notas}</div>}
                  </>
                : <div style={{color:C.warn,fontSize:13}}>
                    ⚠ Non hai normativa gardada para <strong>{chaveNorm}</strong>. Usaranse valores por defecto.
                  </div>
              }
            </div>
          )}

          {/* Aviso especial Mini BTT/DH */}
          {MODALIDADES.find(m=>m.id===modalidade)?.mini && (
            <div style={{
              background:"#1a1200",border:`1px solid ${C.warn}`,
              borderRadius:8,padding:"14px 16px",marginBottom:16,fontSize:12
            }}>
              <div style={{color:C.warn,fontWeight:700,marginBottom:8}}>
                🟡 {modalidade==="MINIBTT"?"Mini BTT – XC (Art. III-N.1)":"Mini DH (Art. III-N.2)"} · Normativa FGC 13/12/2025
              </div>
              <div style={{color:C.muted,lineHeight:1.9}}>
                <strong style={{color:C.text}}>Categorías por ano de nacemento:</strong>{" "}
                Benjamín (2020+) · Promesa (2018-19) · Principiante (2016-17) · Alevín (2014-15) · Infantil (2012-13)<br/>
                <strong style={{color:C.text}}>Orde de saída:</strong>{" "}
                {modalidade==="MINIBTT"
                  ? "1ª proba → sorteo de clubs. Resto de probas → ránking FGC MiniBTT (maior a menor puntos)."
                  : "1ª proba → orde INVERSA á data de inscrición. Resto → orde INVERSA ao ránking FGC MiniDH (o mellor sae último na manga clasificatoria)."
                }<br/>
                <strong style={{color:C.gold}}>Sen puntos UCI nin RFEC.</strong>{" "}
                Licenzas dun día: despois dos federados, sen puntos de orde.<br/>
                <strong style={{color:C.text}}>Orde de categorías na proba:</strong>{" "}
                <span style={{color:C.gold}}>
                  {(modalidade==="MINIBTT"?MINIBTT_ORDE_CATEGORÍAS:MINIDH_ORDE_CATEGORÍAS).join(" → ")}
                </span><br/>
                <strong style={{color:C.text}}>Grellas:</strong>{" "}
                Separadas por sexo en cada categoría. Nº de corredores por fila: decide o organizador.
              </div>
            </div>
          )}

          {/* Info rankings */}
          {modalidade && (
            <div style={{marginBottom:16}}>
              <span style={css.label}>Rankings dispoñibles para esta configuración</span>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                {["UCI","RFEC","FGC"].map(f=>{
                  const r=getRanking(f);
                  return (
                    <div key={f} style={{
                      padding:"6px 14px",borderRadius:8,
                      background:r?C.ok+"22":C.panel2,
                      border:`1px solid ${r?C.ok:C.border}`,
                      fontSize:12
                    }}>
                      <span style={{color:r?C.ok:C.muted,fontWeight:700}}>{f}</span>
                      {r
                        ? <span style={{color:C.ok,marginLeft:6}}>✓ {r.length} corr.</span>
                        : <span style={{color:C.muted,marginLeft:6}}>non cargado</span>
                      }
                    </div>
                  );
                })}
              </div>
              {rankingsDisp.length===0 && (
                <Alert tipo="warn">Sen rankings cargados. A grella xerarase só pola orde de inscrición.</Alert>
              )}
            </div>
          )}

          <Btn onClick={()=>setPaso(1)} disabled={(!modalidade)||!nivel}>
            Continuar →
          </Btn>
        </div>
      )}

      {/* PASO 1: Inscritos */}
      {paso===1 && (
        <div style={css.card}>
          <span style={css.label}>Paso 2 · Cargar listaxe de inscritos</span>
          <div style={{fontSize:12,color:C.muted,marginBottom:12}}>
            O sistema cruzará automaticamente co UCI ID ou Licenza nos rankings <strong style={{color:C.gold}}>{rankingsDisp.join(", ")||"(ningún cargado)"}</strong>.
          </div>
          <DropZone onFile={onFicheiroInscritos} label="Arrastra o ficheiro de inscritos (.xlsx ou .csv)"/>
          {fichNome && <div style={{marginTop:8,fontSize:12,color:C.ok}}>✓ {fichNome}</div>}
          {alertas.map((a,i)=><Alert key={i} tipo={a.tipo}>{a.txt}</Alert>)}

          {/* Demo */}
          <div style={{marginTop:12,fontSize:12,color:C.muted}}>
            Sen ficheiro? {" "}
            <span style={{color:C.gold,cursor:"pointer",textDecoration:"underline"}} onClick={()=>{
              const filaDemo=[
                {dorsal:"1",nome:"Xan Díaz Fonte",equipo:"Galicia Pro Team",uci_id:"GAL19961130",licenza:"G-010",inscricion:1},
                {dorsal:"2",nome:"Anxo Fernández Rei",equipo:"Galicia Pro Team",uci_id:"GAL19900101",licenza:"G-001",inscricion:2},
                {dorsal:"3",nome:"Fernando Cid Brea",equipo:"Lugo Riders",uci_id:"GAL19910210",licenza:"G-006",inscricion:3},
                {dorsal:"4",nome:"Hugo Blanco Seoane",equipo:"Ferrol MTB",uci_id:"GAL19930825",licenza:"G-008",inscricion:4},
                {dorsal:"5",nome:"Brais Nogueira López",equipo:"Vigo CC",uci_id:"GAL19920315",licenza:"G-002",inscricion:5},
                {dorsal:"6",nome:"Emilio Varela Costa",equipo:"Santiago SC",uci_id:"",licenza:"G-005",inscricion:6},
                {dorsal:"7",nome:"Carlos Míguez Paz",equipo:"A Coruña MTB",uci_id:"GAL19880720",licenza:"G-003",inscricion:7},
                {dorsal:"8",nome:"Iván Pardo Neto",equipo:"Vigo CC",uci_id:"GAL19870415",licenza:"G-009",inscricion:8},
                {dorsal:"9",nome:"David Ramos Otero",equipo:"Ourense Bici",uci_id:"GAL19950601",licenza:"G-004",inscricion:9},
                {dorsal:"10",nome:"Gael Torres Rivas",equipo:"Pontevedra CC",uci_id:"",licenza:"",inscricion:10},
                {dorsal:"11",nome:"Laura Moure Soto",equipo:"A Coruña MTB",uci_id:"GAL20000305",licenza:"G-011",inscricion:11},
                {dorsal:"12",nome:"Marcos Salgado Piñeiro",equipo:"Ourense Bici",uci_id:"",licenza:"",inscricion:12},
              ];
              const mapped=filaDemo.map((r,i)=>({
                id:i,dorsal:r.dorsal,nome:r.nome,equipo:r.equipo,
                uci_id:r.uci_id,licenza:r.licenza,categoría:"Elite",
                inscricion:r.inscricion,pts_uci:0,pts_rfec:0,pts_fgc:0,
                dns:false,bloqueado:false,pos_bloq:null,alertas:[],
              }));
              setCorr(mapped); setFichNome("demo_inscritos.csv");
              setGrella([]); setMatriz([]); setXerado(false);
              setPaso(2);
            }}>Cargar datos de demostración</span>
          </div>
          <div style={{marginTop:16}}>
            <Btn variant="secondary" onClick={()=>setPaso(0)}>← Volver</Btn>
          </div>
        </div>
      )}

      {/* PASO 2+: Inscritos cargados + Accións */}
      {paso>=2 && (
        <>
          {/* Panel de accións */}
          <div style={{...css.card,display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
            <Btn onClick={xerarGrella}>🏁 {xerado?"Rexenerar":"Xerar"} grella</Btn>
            {xerado && <>
              <Btn variant="secondary" onClick={()=>setVistaTab("visual")} style={{opacity:vistaTab==="visual"?1:.6}}>Visual</Btn>
              <Btn variant="secondary" onClick={()=>setVistaTab("lista")} style={{opacity:vistaTab==="lista"?1:.6}}>Lista chamada</Btn>
              <Btn variant="secondary" onClick={()=>setVistaTab("log")} style={{opacity:vistaTab==="log"?1:.6}}>📄 LOG</Btn>
              <Btn variant="ghost" onClick={()=>descargarTexto(logTexto,`grella_${modalidade}_${nivel}_${Date.now()}.log`)}>
                ⬇ Exportar LOG
              </Btn>
            </>}
            <Btn variant="secondary" onClick={()=>{setPaso(0);setGrella([]);setMatriz([]);setXerado(false);}}>↺ Nova proba</Btn>
          </div>

          {alertas.length>0 && (
            <div style={css.card}>
              <span style={css.label}>⚠ {alertas.length} alertas de cruzamento</span>
              <div style={{maxHeight:150,overflowY:"auto"}}>
                {alertas.map((a,i)=><Alert key={i} tipo={a.tipo}>{a.txt}</Alert>)}
              </div>
              <div style={{fontSize:11,color:C.muted,marginTop:6}}>
                Estes corredores irán ao final da grella. Podes asignarlles puntos manualmente no futuro cargando o ranking correspondente.
              </div>
            </div>
          )}

          {/* Xestión DNS / Bloqueo */}
          <div style={css.card}>
            <span style={css.label}>Xestión de corredores · {corredores.length} inscritos · {fichNome}</span>
            <div style={{fontSize:11,color:C.muted,marginBottom:10}}>
              DNS = ausencia confirmada · 🔒 = fixar posición na grella
            </div>
            <div style={{maxHeight:280,overflowY:"auto",border:`1px solid ${C.border}`,borderRadius:8}}>
              {corredores.map((c,i)=>{
                const posGrella = xerado ? grella.findIndex(x=>x.id===c.id)+1 : null;
                return (
                  <div key={c.id} style={{
                    display:"flex",alignItems:"center",gap:8,padding:"7px 10px",
                    background:i%2===0?C.panel2:C.bg,
                    borderBottom:`1px solid ${C.border}`,
                    opacity:c.dns?.5:1,
                  }}>
                    <span style={{width:38,color:C.gold,fontWeight:700,fontSize:13,flexShrink:0}}>D{c.dorsal}</span>
                    <span style={{flex:1,fontSize:12,textDecoration:c.dns?"line-through":"none",minWidth:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                      {c.nome}
                    </span>
                    <span style={{width:90,fontSize:11,color:C.muted,flexShrink:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.equipo}</span>
                    {xerado && <span style={{width:36,fontSize:11,color:C.gold,textAlign:"center",flexShrink:0}}>P{posGrella||"–"}</span>}
                    <div style={{display:"flex",gap:4,flexShrink:0}}>
                      {["UCI","RFEC","FGC"].map(f=>{
                        const pts=f==="UCI"?c.pts_uci:f==="RFEC"?c.pts_rfec:c.pts_fgc;
                        return pts>0?<Tag key={f} color={C.muted}>{f}:{pts}</Tag>:null;
                      })}
                    </div>
                    <button onClick={()=>toggleDNS(c.id)} style={{
                      padding:"3px 8px",borderRadius:4,fontSize:11,fontWeight:700,
                      cursor:"pointer",border:"none",flexShrink:0,
                      background:c.dns?C.danger:C.panel,color:c.dns?"#fff":C.muted
                    }}>DNS</button>
                    <button onClick={()=>toggleBloq(c.id)} style={{
                      padding:"3px 8px",borderRadius:4,fontSize:11,
                      cursor:"pointer",border:"none",flexShrink:0,
                      background:c.bloqueado?C.gold:C.panel,
                      color:c.bloqueado?"#0d0d1f":C.muted,fontWeight:c.bloqueado?700:400
                    }}>
                      {c.bloqueado?`🔒P${c.pos_bloq}`:"🔒"}
                    </button>
                  </div>
                );
              })}
            </div>
          </div>

          {/* VISTA GRELLA VISUAL */}
          {xerado && vistaTab==="visual" && (
            <div style={css.card}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
                <span style={css.label}>Grella visual · {modalidade} · {nivel} · {categoría||"Todas"}</span>
                <span style={{fontSize:11,color:C.muted}}>Arrastra para reordenar manualmente</span>
              </div>
              <div style={{overflowX:"auto"}}>
                {matriz.map((fila,fi)=>(
                  <div key={fi} style={{marginBottom:14}}>
                    <div style={{fontSize:10,color:fi===0?C.gold:C.muted,textTransform:"uppercase",
                      letterSpacing:1.5,marginBottom:6,fontWeight:fi===0?700:400}}>
                      Fila {fi+1}{fi===0?" – DIANTEIRA":""}
                    </div>
                    <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                      {fila.map(c=>{
                        const gi=grella.indexOf(c);
                        const isOver=dragOver===gi;
                        return (
                          <div key={c.id}
                            draggable
                            onDragStart={()=>setDrag(gi)}
                            onDragEnter={()=>setDragOver(gi)}
                            onDragOver={e=>e.preventDefault()}
                            onDragEnd={onDragEnd}
                            style={{
                              width:105,padding:"10px 8px",borderRadius:8,cursor:"grab",
                              userSelect:"none",transition:"transform .1s,box-shadow .1s",
                              background:c.bloqueado?"linear-gradient(135deg,#1e1800,#2a2200)"
                                :fi%2===0?C.panel:C.panel2,
                              border:c.bloqueado?`2px solid ${C.gold}`
                                :isOver?`2px solid ${C.blue}`:`1px solid ${C.border}`,
                              boxShadow:isOver?`0 0 14px ${C.blue}66`:"none",
                              transform:isOver?"scale(1.04)":"scale(1)",
                            }}>
                            <div style={{fontWeight:700,fontSize:20,color:C.gold,
                              textAlign:"center",fontFamily:"Verdana,monospace",marginBottom:3}}>
                              {c.dorsal}
                            </div>
                            <div style={{fontSize:10,color:C.text,textAlign:"center",
                              lineHeight:1.3,marginBottom:3}}>
                              {c.nome.split(" ").slice(0,2).join(" ")}
                            </div>
                            <div style={{fontSize:9,color:C.muted,textAlign:"center"}}>
                              {c.equipo.slice(0,15)}
                            </div>
                            <div style={{fontSize:9,color:C.muted,textAlign:"center",marginTop:4}}>
                              {[c.pts_uci&&`U${c.pts_uci}`,c.pts_rfec&&`R${c.pts_rfec}`,c.pts_fgc&&`F${c.pts_fgc}`]
                                .filter(Boolean).join(" ")||"–"}
                            </div>
                            <div style={{fontSize:9,color:C.border,textAlign:"center",marginTop:2}}>
                              P{gi+1}{c.bloqueado?" 🔒":""}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* VISTA LISTA CHAMADA */}
          {xerado && vistaTab==="lista" && (
            <div style={css.card}>
              <span style={css.label}>Listaxe de chamada · Para difusión WhatsApp/Telegram</span>
              <div style={{background:C.bg,borderRadius:8,border:`1px solid ${C.border}`,
                overflow:"hidden",maxHeight:500,overflowY:"auto"}}>
                <div style={{padding:"8px 12px",background:C.panel2,
                  fontWeight:700,fontSize:12,color:C.gold,
                  borderBottom:`1px solid ${C.border}`}}>
                  {modalidade} · {NIVEIS_COMPETICIÓN.find(n=>n.id===nivel)?.label} · {categoría||"Todas"} · {new Date().toLocaleDateString("gl-ES")}
                </div>
                {grella.map((c,i)=>(
                  <div key={c.id} style={{
                    display:"flex",gap:10,padding:"6px 12px",
                    background:i%2===0?C.panel2:C.bg,
                    borderBottom:`1px solid ${C.border}`,
                    fontFamily:"Verdana,monospace",fontSize:12,
                  }}>
                    <span style={{width:28,color:C.muted,flexShrink:0}}>{i+1}.</span>
                    <span style={{width:42,color:C.gold,fontWeight:700,flexShrink:0}}>D{c.dorsal}</span>
                    <span style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.nome}</span>
                    <span style={{width:120,color:C.muted,fontSize:11,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.equipo}</span>
                    <span style={{width:60,fontSize:10,color:C.muted,textAlign:"right",flexShrink:0}}>
                      {[c.pts_uci&&`U:${c.pts_uci}`,c.pts_rfec&&`R:${c.pts_rfec}`,c.pts_fgc&&`F:${c.pts_fgc}`].filter(Boolean).join(" ")||"–"}
                    </span>
                    {c.bloqueado&&<span style={{color:C.gold,flexShrink:0}}>🔒</span>}
                  </div>
                ))}
              </div>
              <div style={{marginTop:10,display:"flex",gap:8}}>
                <Btn small variant="ghost" onClick={()=>{
                  const txt=grella.map((c,i)=>`${i+1}. D${c.dorsal} ${c.nome} (${c.equipo})`).join("\n");
                  navigator.clipboard?.writeText(txt).then(()=>alert("Copiado!"));
                }}>📋 Copiar para portapapeis</Btn>
              </div>
            </div>
          )}

          {/* VISTA LOG */}
          {xerado && vistaTab==="log" && (
            <div style={css.card}>
              <span style={css.label}>LOG de auditoría · Proba xurídica do procedemento</span>
              <pre style={{
                fontFamily:"monospace",fontSize:11,color:"#90ee90",
                background:"#030310",padding:14,borderRadius:8,
                overflowX:"auto",whiteSpace:"pre",lineHeight:1.6,
                border:`1px solid #1a3a1a`,maxHeight:400,overflowY:"auto"
              }}>
                {logTexto}
              </pre>
            </div>
          )}
        </>
      )}
    </div>
  );
}

export default App;