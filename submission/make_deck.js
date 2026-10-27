// FreshGuard external-presentation deck (pptxgenjs).
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const FIG = path.join(__dirname, "figures");
const fig = (f) => path.join(FIG, f);

const C = {
  ink: "14323A", navy: "1E2761", teal: "028090", blue: "2B6CB0",
  green: "2F855A", amber: "B7791F", seafoam: "00A896", grey: "5A6B78",
  light: "EDF2F7", card: "F3F7F8", white: "FFFFFF", muted: "7A8894",
};
const HEAD = "Cambria", BODY = "Calibri";

const pres = new pptxgen();
pres.defineLayout({ name: "W", width: 13.333, height: 7.5 });
pres.layout = "W";
pres.author = "FreshGuard Project Team";
pres.title = "FreshGuard — Freshness-Aware Adaptive RAG";

const FOOT = "FreshGuard  ·  Freshness-Aware Adaptive RAG with a Gated Verification Chain  ·  Anurag University";

function footer(slide, n) {
  slide.addText(FOOT, { x: 0.5, y: 7.12, w: 10.8, h: 0.3, fontFace: BODY, fontSize: 9,
    color: C.muted, align: "left", isTextBox: true, margin: 0 });
  slide.addText(String(n), { x: 12.4, y: 7.12, w: 0.5, h: 0.3, fontFace: BODY, fontSize: 10,
    color: C.muted, align: "right", isTextBox: true, margin: 0 });
}

function title(slide, text, color = C.teal) {
  slide.addText(text, { x: 0.6, y: 0.42, w: 12.1, h: 0.9, fontFace: HEAD, fontSize: 30,
    bold: true, color, align: "left", isTextBox: true, margin: 0, valign: "middle" });
}

// bullet list block
function bullets(slide, items, { x = 0.7, y = 1.5, w = 12, fontSize = 16, color = C.ink } = {}) {
  const runs = items.map((it, i) => {
    const t = typeof it === "string" ? it : it.t;
    const lvl = typeof it === "string" ? 0 : (it.lvl || 0);
    const bold = typeof it === "object" && it.bold;
    return { text: t, options: { bullet: { indent: 16 }, indentLevel: lvl, bold: !!bold,
      color: lvl ? C.grey : color, fontSize: lvl ? fontSize - 2 : fontSize, fontFace: BODY,
      paraSpaceAfter: 8, breakLine: true } };
  });
  slide.addText(runs, { x, y, w, h: 5.3, align: "left", isTextBox: true, valign: "top", margin: 0 });
}

let N = 0;
function content(text, items, opts) {
  N += 1;
  const s = pres.addSlide();
  title(s, text);
  if (items) bullets(s, items, opts);
  footer(s, N);
  return s;
}

// ---------------- 1. TITLE ----------------
(function () {
  N += 1;
  const s = pres.addSlide();
  s.background = { color: C.navy };
  s.addText("Major Project External Presentation", { x: 0.9, y: 0.7, w: 11.5, h: 0.5,
    fontFace: BODY, fontSize: 18, color: C.seafoam, isTextBox: true, margin: 0 });
  s.addText("FreshGuard", { x: 0.9, y: 1.45, w: 11.5, h: 1.0, fontFace: HEAD, fontSize: 54,
    bold: true, color: C.white, isTextBox: true, margin: 0 });
  s.addText("Freshness-Aware Adaptive Retrieval-Augmented Generation with a Learned Retrieval Gate and Statement Verification",
    { x: 0.9, y: 2.5, w: 11.2, h: 1.0, fontFace: HEAD, fontSize: 22, italic: true, color: C.light, isTextBox: true, margin: 0 });
  const team = [
    "Team members", "J. Sai Teja — 23EG106C22", "P. Supriya — 23EG106C52",
    "Sk. Mubeena — 23EG106C58", "U. Venkatesh — 23EG106C63",
  ];
  s.addText(team.map((t, i) => ({ text: t, options: { bold: i === 0, color: i === 0 ? C.seafoam : C.white,
    fontSize: i === 0 ? 15 : 15, fontFace: BODY, breakLine: true, paraSpaceAfter: 5 } })),
    { x: 0.95, y: 4.0, w: 6.0, h: 2.4, isTextBox: true, margin: 0, valign: "top" });
  s.addText([
    { text: "Guide", options: { bold: true, color: C.seafoam, fontSize: 15, fontFace: BODY, breakLine: true, paraSpaceAfter: 5 } },
    { text: "Mohammad Arshad", options: { color: C.white, fontSize: 15, fontFace: BODY, breakLine: true, paraSpaceAfter: 2 } },
    { text: "Assistant Professor", options: { color: C.light, fontSize: 14, fontFace: BODY, breakLine: true } },
    { text: "Department of Artificial Intelligence, Anurag University", options: { color: C.light, fontSize: 13, fontFace: BODY } },
  ], { x: 7.1, y: 4.0, w: 5.3, h: 2.4, isTextBox: true, margin: 0, valign: "top" });
})();

// ---------------- 2. CONTENTS ----------------
(function () {
  N += 1;
  const s = pres.addSlide();
  title(s, "Contents");
  const left = ["Abstract", "Introduction", "Literature Survey", "Research Gap", "Objectives", "Problem Statement"];
  const right = ["Methodology", "Implementation", "Results", "Discussion", "Conclusion & Future Work", "References"];
  const mk = (arr, start) => arr.map((t, i) => ({ text: `${start + i}.  ${t}`, options: {
    color: C.ink, fontSize: 18, fontFace: BODY, breakLine: true, paraSpaceAfter: 12 } }));
  s.addText(mk(left, 1), { x: 1.1, y: 1.7, w: 5.3, h: 5, isTextBox: true, margin: 0, valign: "top" });
  s.addText(mk(right, 7), { x: 7.0, y: 1.7, w: 5.3, h: 5, isTextBox: true, margin: 0, valign: "top" });
  footer(s, N);
})();

// ---------------- 3. ABSTRACT ----------------
content("Abstract", [
  { t: "LLMs answer fluently but also confidently when their knowledge is out of date; static RAG retrieves for every query and never checks the answer it writes.", bold: false },
  "FreshGuard — a freshness-aware, adaptive RAG system that verifies every answer before showing it.",
  { t: "Learned freshness gate: a logistic-regression estimator over TF-IDF + temporal features predicts P(knowledge is stale); retrieve only when stale or low-confidence.", lvl: 1 },
  { t: "Retrieval over a FAISS / TF-IDF store with relevance × recency ranking and a year gate.", lvl: 1 },
  { t: "Live ingestion on a KB miss: fetch from Tavily web search + Grok, embed, persist, re-retrieve and verify.", lvl: 1 },
  { t: "Statement-level verification (supported / unsupported / contradicted) + a calibrated confidence.", lvl: 1 },
  "Results: freshness-decision accuracy 1.00, factuality 1.00, hallucination F1 1.00, ECE 0.10; 45 tests pass identically across 10 runs.",
  "Main weakness: the gate is an uncalibrated default and live ingestion needs a key — calibration is the next step.",
], { y: 1.45, fontSize: 15 });

// ---------------- 4. INTRODUCTION ----------------
content("Introduction", [
  "LLMs answer from a fixed training snapshot — for anything that has changed since, they can be confidently wrong, with no freshness signal.",
  "RAG grounds answers in retrieved passages, but it has two weaknesses this project targets:",
  { t: "It retrieves for every query — wasting latency and cost on timeless facts the model already knows.", lvl: 1 },
  { t: "It does not check the final answer: the model can misread, mis-combine, or state what the evidence does not support.", lvl: 1 },
  "Our goal: two learned gates around generation — a freshness gate (when to retrieve) and a verification gate (is the answer supported) — plus live ingestion to fill knowledge-base gaps and honest abstention when it cannot.",
], { y: 1.55, fontSize: 17 });

// ---------------- 5. LITERATURE SURVEY ----------------
content("Literature Survey", [
  { t: "RAG — Lewis et al.: retrieve, then generate; no check on the final answer, retrieves every time.", bold: false },
  "Knowledge freshness / temporal QA — fixed cutoff or time-decay rules; query-independent, cannot learn what is time-sensitive.",
  "Adaptive / Self-RAG — let the model decide when to retrieve; decision hidden inside the LLM, not inspectable or tunable.",
  "Hybrid retrieval & ranking — dense + BM25, recency weighting, reciprocal rank fusion (Cormack et al.).",
  "Hallucination detection — NLI entailment (MultiNLI), FActScore atomic facts, SelfCheckGPT consistency; used offline, not as a gate.",
  "Calibration — expected calibration error (ECE); reported, rarely acted on inside a running system.",
], { y: 1.5, fontSize: 16 });

// ---------------- 6. RESEARCH GAP ----------------
content("Literature Survey — Research Gap", [
  "Freshness estimation, the retrieval decision, live augmentation and verification are studied separately; the first three rarely act on the query, the last is rarely deployed.",
  "Gaps this project addresses:",
  { t: "A learned, query-specific freshness estimator — a staleness probability, not a fixed time rule.", lvl: 1 },
  { t: "An explicit, tunable retrieval decision — skip for fresh, trigger for stale / low-confidence.", lvl: 1 },
  { t: "Live ingestion that fills a KB miss from Tavily + Grok, then embeds, persists, re-retrieves and verifies.", lvl: 1 },
  { t: "A year gate + honest abstention — never answer a 2026 question from a 2022 document.", lvl: 1 },
  { t: "Per-statement verification + a shown, calibrated confidence, with ECE measured.", lvl: 1 },
  { t: "A fully offline, deterministic default — reproducibly testable without any network or key.", lvl: 1 },
], { y: 1.45, fontSize: 15 });

// ---------------- 7. OBJECTIVES ----------------
content("Objectives", [
  "Build query preprocessing and a learned freshness estimator that outputs a calibrated staleness probability.",
  "Implement an adaptive retrieval decision (retrieve if stale or low-confidence) with admin-configurable thresholds.",
  "Implement FAISS / TF-IDF retrieval with relevance × recency ranking, plus live ingestion from Tavily + Grok on a KB miss.",
  "Implement statement-level verification, corrective action, and a three-signal confidence score.",
  "Provide a web UI with a freshness badge, confidence meter, per-statement verification, citations, history, feedback and admin.",
  "Evaluate retrieval decisions, factuality, hallucination detection and calibration — and report results honestly, including failures.",
], { y: 1.55, fontSize: 17 });

// ---------------- 8. PROBLEM STATEMENT ----------------
content("Problem Statement", [
  "Existing LLM / basic-RAG systems, when freshness matters:",
  { t: "Stale answers with no signal — time-sensitive questions answered from a fixed snapshot.", lvl: 1 },
  { t: "Retrieval for every query — latency and cost even for timeless facts.", lvl: 1 },
  { t: "Answers shown unchecked — nothing verifies the answer against the evidence.", lvl: 1 },
  { t: "Guessing instead of abstaining — a wrong-period document is used as if it answered.", lvl: 1 },
  "Need: a system that (1) estimates per-query staleness, (2) retrieves only when needed, (3) fetches + ingests fresh evidence on a miss, (4) verifies every statement with a calibrated confidence, and (5) abstains honestly when no evidence can be found.",
], { y: 1.5, fontSize: 16 });

// ---------------- 9. METHODOLOGY — ARCHITECTURE ----------------
(function () {
  N += 1;
  const s = pres.addSlide();
  title(s, "Methodology — System Architecture");
  s.addImage({ path: fig("architecture.png"), x: 0.7, y: 1.4, w: 11.9, h: 4.9 });
  s.addText("Web UI → FastAPI backend → pipeline stages → FAISS index + SQLite; LLM and live sources via adapters; everything offline and deterministic by default.",
    { x: 0.7, y: 6.4, w: 11.9, h: 0.6, fontFace: BODY, fontSize: 13, color: C.grey, isTextBox: true, margin: 0 });
  footer(s, N);
})();

// ---------------- 10. FRESHNESS GATE ----------------
(function () {
  N += 1;
  const s = pres.addSlide();
  title(s, "Methodology — Freshness Gate & Adaptive Retrieval");
  bullets(s, [
    { t: "Learned estimator: TF-IDF (word n-grams) + LogisticRegression over the query → P(stale).", bold: false },
    "Temporal cues carry the signal (years, 'latest', 'current', 'now'), but the model generalises beyond them.",
    "Model self-confidence derived monotonically from the freshness score.",
    { t: "Decision: retrieve if freshness ≥ 0.55 OR confidence < 0.60; else answer directly.", bold: true },
    "Thresholds are admin-configurable; the decision is explicit and inspectable, not hidden in the LLM.",
  ], { x: 0.7, y: 1.55, w: 12, fontSize: 16 });
  s.addImage({ path: fig("gate.png"), x: 1.4, y: 5.05, w: 10.6, h: 1.85 });
  footer(s, N);
})();

// ---------------- 11. RETRIEVAL + LIVE INGESTION ----------------
(function () {
  N += 1;
  const s = pres.addSlide();
  title(s, "Methodology — Retrieval, Ranking & Live Ingestion");
  bullets(s, [
    "FAISS index over TF-IDF embeddings (offline, deterministic); relevance × recency ranking; min-relevance floor.",
    { t: "Year gate: if the query names a year no retrieved doc covers, local evidence is insufficient.", bold: true },
    "On a KB miss → Tavily web search + Grok supply fresh passages (keyed by env vars).",
    "Each passage is embedded into FAISS, persisted to SQLite, re-retrieved, generated from and verified like a seed doc.",
    "No key configured → abstain honestly instead of citing a wrong-period document.",
  ], { x: 0.7, y: 1.55, w: 12, fontSize: 16 });
  s.addImage({ path: fig("live_ingest.png"), x: 1.6, y: 5.1, w: 10.1, h: 1.8 });
  footer(s, N);
})();

// ---------------- 12. GENERATION + VERIFICATION ----------------
content("Methodology — Generation & Verification", [
  { t: "Generation is pluggable: default offline grounded extractive LLM; OpenAI / Anthropic / Grok via key.", bold: false },
  { t: "Parametric memory holds only timeless facts (models a knowledge cutoff) — time-sensitive queries need retrieval.", lvl: 1 },
  "Verification: split the answer into statements; label each against the evidence:",
  { t: "supported — covered, no conflict;  unsupported — absent or an unseen entity;", lvl: 1 },
  { t: "contradicted — negation flip, numeric mismatch, or entity swap;  review — partial support.", lvl: 1 },
  "Correction: keep supported / review, drop unsupported (with a warning), flag contradicted; abstain if nothing is supported.",
  { t: "Confidence = 0.50·support + 0.30·evidence-relevance + 0.20·freshness-alignment, scaled to 0–100.", bold: true },
], { y: 1.5, fontSize: 15 });

// ---------------- 13. IMPLEMENTATION — TECH ----------------
content("Implementation — Technologies & Modules", [
  "Python 3.11 · FastAPI + Uvicorn · Pydantic · SQLite — offline-first backend.",
  "scikit-learn (TF-IDF, LogisticRegression) · FAISS (faiss-cpu) via LangChain vector store · joblib · NumPy / SciPy.",
  "Live sources: Tavily web-search REST API · xAI Grok (OpenAI-compatible) via requests · python-dotenv for keys.",
  "Single-page HTML/CSS/JS frontend served by FastAPI (Ask · History · Admin).",
  "Modular pipeline — preprocess, freshness, decision, retrieval, live_sources, llm, verify, confidence, pipeline, storage, evaluation.",
  { t: "Testing: 45 pytest unit / API / end-to-end / evaluation tests; live path tested with a mocked HTTP layer; no key needed.", bold: true },
], { y: 1.5, fontSize: 16 });

// ---------------- 14. IMPLEMENTATION — CONFIG ----------------
(function () {
  N += 1;
  const s = pres.addSlide();
  title(s, "Implementation — Pipeline Configuration");
  const rows = [
    [{ text: "Parameter", options: { bold: true, color: C.white } }, { text: "Value", options: { bold: true, color: C.white } }],
    ["Freshness / confidence threshold", "0.55 / 0.60 (admin-configurable)"],
    ["Embeddings / index", "TF-IDF → FAISS (inner product = cosine)"],
    ["Top-k / min-relevance / half-life", "4  /  0.05  /  2.0 years"],
    ["Verification support hi / lo", "0.32 supported  /  0.15 unsupported"],
    ["Confidence weights", "support 0.50 · relevance 0.30 · freshness 0.20"],
    ["LLM provider (default)", "offline grounded; OpenAI / Anthropic / Grok via key"],
    ["Live sources", "Tavily (TAVILY_API_KEY) · Grok (XAI_API_KEY)"],
  ];
  s.addTable(rows, { x: 0.8, y: 1.7, w: 11.7, colW: [5.2, 6.5], fontFace: BODY, fontSize: 15,
    color: C.ink, border: { type: "solid", color: "CBD5E0", pt: 1 }, align: "left", valign: "middle",
    rowH: 0.5, fill: { color: C.white }, margin: 6,
    // header + zebra via per-row would need manual; keep header fill:
  });
  // header fill overlay
  s.addShape(pres.ShapeType.rect, { x: 0.8, y: 1.7, w: 11.7, h: 0.5, fill: { color: C.navy } });
  s.addText([{ text: "Parameter", options: { bold: true, color: C.white, fontSize: 15, fontFace: BODY } }],
    { x: 0.86, y: 1.7, w: 5.2, h: 0.5, valign: "middle", isTextBox: true, margin: 6 });
  s.addText([{ text: "Value", options: { bold: true, color: C.white, fontSize: 15, fontFace: BODY } }],
    { x: 6.06, y: 1.7, w: 6.4, h: 0.5, valign: "middle", isTextBox: true, margin: 6 });
  footer(s, N);
})();

// ---------------- 15. RESULTS — METRICS ----------------
(function () {
  N += 1;
  const s = pres.addSlide();
  title(s, "Results — Controlled Benchmark");
  // stat callouts
  const stats = [
    ["1.00", "Freshness-decision accuracy", C.teal],
    ["1.00", "Factuality", C.blue],
    ["1.00", "Hallucination F1", C.green],
    ["0.10", "Calibration error (ECE)", C.amber],
  ];
  stats.forEach(([v, l, col], i) => {
    const x = 0.7 + i * 3.08;
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.5, w: 2.85, h: 1.5, rectRadius: 0.08,
      fill: { color: C.card }, line: { color: col, width: 1.25 }, shadow: { type: "outer", color: "BBBBBB", blur: 4, offset: 2, angle: 90, opacity: 0.4 } });
    s.addText(v, { x, y: 1.6, w: 2.85, h: 0.85, fontFace: HEAD, fontSize: 40, bold: true, color: col, align: "center", isTextBox: true, margin: 0 });
    s.addText(l, { x: x + 0.1, y: 2.42, w: 2.65, h: 0.5, fontFace: BODY, fontSize: 12, color: C.grey, align: "center", isTextBox: true, margin: 0 });
  });
  s.addImage({ path: fig("metrics.png"), x: 2.0, y: 3.2, w: 9.3, h: 3.6 });
  footer(s, N);
})();

// ---------------- 16. RESULTS — LIVE INGESTION ----------------
(function () {
  N += 1;
  const s = pres.addSlide();
  title(s, "Results — Future-Event Live Ingestion");
  bullets(s, [
    { t: "Before: a 2026 World Cup query was answered from the 2022 document — a stale, year-mismatched source.", bold: false },
    "With the year gate, no key: the system no longer cites the wrong year — it abstains and names the key to enable a fetch.",
    { t: "With Tavily + Grok: fetch → embed → persist → re-retrieve → verify → a 2026 answer, both statements supported and cited.", bold: true },
  ], { x: 0.7, y: 1.5, w: 12, fontSize: 16 });
  s.addImage({ path: fig("future_event.png"), x: 1.9, y: 3.55, w: 9.5, h: 3.3 });
  footer(s, N);
})();

// ---------------- 17. DISCUSSION ----------------
content("Discussion — Limitations & Responsible AI", [
  "Uncalibrated gate: the 0.55 / 0.60 thresholds are defaults; ECE 0.10 shows room to tighten confidence vs. correctness.",
  "Lightweight verification: deterministic rules can be conservative on paraphrase, and check faithfulness, not completeness.",
  "Small evaluation: 12 labelled queries, 8 hallucination pairs, a demonstrated live-ingestion flow — observations, not rates.",
  "Live ingestion depends on an external key; answer quality depends on what Tavily / Grok return.",
  "Responsible AI:",
  { t: "Research prototype, not professional advice; the system shows its evidence and abstains when it is weak.", lvl: 1 },
  { t: "Index and history stay local; queries go to the live services only when enabled; keys are git-ignored.", lvl: 1 },
  { t: "Failures are reported alongside successes — the limits are part of the record.", lvl: 1 },
], { y: 1.5, fontSize: 15 });

// ---------------- 18. CONCLUSION ----------------
content("Conclusion & Future Work", [
  { t: "A small, mostly-offline pipeline can make an LLM freshness-aware and auditable: a learned freshness gate + adaptive retrieval with live ingestion + statement verification kept stale and unsupported answers away from the user.", bold: true },
  "Benchmark: freshness-decision / factuality / hallucination F1 = 1.00; ECE 0.10; 45 tests identical across 10 runs.",
  "Live ingestion turned a wrongly-answered future-event question into a verified, cited answer — or an honest abstention.",
  "Future work:",
  { t: "Calibrate the thresholds and the shown confidence on a held-out set; add char n-grams + entity-recency features.", lvl: 1 },
  { t: "Add a transformer encoder + BM25 with rank fusion; cache and dedup live-ingested docs.", lvl: 1 },
  { t: "NLI entailment, atomic-fact and completeness checks; local-LLM option; larger, harder evaluation set.", lvl: 1 },
], { y: 1.5, fontSize: 15 });

// ---------------- 19. REFERENCES ----------------
(function () {
  N += 1;
  const s = pres.addSlide();
  title(s, "References (selected)");
  const refs = [
    "[1] P. Lewis et al., Retrieval-Augmented Generation for Knowledge-Intensive NLP Tasks, NeurIPS, 2020.",
    "[2] A. Asai et al., Self-RAG: Learning to Retrieve, Generate and Critique through Self-Reflection, ICLR, 2024.",
    "[3] A. Mallen et al., When Not to Trust Language Models, ACL, 2023.",
    "[4] S. Jin et al., FreshLLMs: Refreshing LLMs with Search-Engine Augmentation, Findings of ACL, 2024.",
    "[5] J. Johnson, M. Douze, H. Jégou, Billion-Scale Similarity Search with GPUs (FAISS), IEEE TBD, 2019.",
    "[6] G. Cormack, C. Clarke, S. Büttcher, Reciprocal Rank Fusion, ACM SIGIR, 2009.",
    "[7] S. Min et al., FActScore, EMNLP, 2023.",
    "[8] P. Manakul et al., SelfCheckGPT, EMNLP, 2023.",
    "[9] C. Guo et al., On Calibration of Modern Neural Networks, ICML, 2017.",
    "[10] F. Pedregosa et al., Scikit-learn: Machine Learning in Python, JMLR, 2011.",
    "[11] Tavily Search API docs, 2024; xAI Grok API docs, 2024; LangChain (H. Chase), 2022.",
  ];
  s.addText(refs.map((t) => ({ text: t, options: { color: C.ink, fontSize: 13, fontFace: BODY,
    breakLine: true, paraSpaceAfter: 7 } })), { x: 0.8, y: 1.6, w: 11.8, h: 5.2, isTextBox: true, margin: 0, valign: "top" });
  footer(s, N);
})();

pres.writeFile({ fileName: path.join(__dirname, "FreshGuard_Presentation.pptx") })
  .then((f) => console.log("wrote", f));
