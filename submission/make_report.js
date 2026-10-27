// FreshGuard project report generator (docx-js).
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, PageBreak,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, ImageRun,
  TableOfContents, PageNumber, Header, Footer, LevelFormat, Tab,
  PositionalTab, PositionalTabAlignment, PositionalTabLeader, convertInchesToTwip,
} = require("docx");

const FIG = path.join(__dirname, "figures");
const img = (f) => fs.readFileSync(path.join(FIG, f));

const NAVY = "1E2761";
const ACCENT = "2b6cb0";
const MUTE = "4a5568";

// ---------- helpers ----------
const H1 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 140 },
  children: [new TextRun({ text, bold: true })] });
const H2 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 },
  children: [new TextRun({ text, bold: true })] });
const H3 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 160, after: 80 },
  children: [new TextRun({ text, bold: true, italics: true })] });

const P = (text, opts = {}) => new Paragraph({
  spacing: { after: 120, line: 300 }, alignment: opts.align || AlignmentType.JUSTIFIED,
  children: Array.isArray(text) ? text : [new TextRun({ text, ...opts })],
});
const center = (runs, opts = {}) => new Paragraph({ alignment: AlignmentType.CENTER,
  spacing: { after: opts.after ?? 60, before: opts.before ?? 0 }, children: runs });

const chapTitle = (n, name) => [
  center([new TextRun({ text: `CHAPTER ${n}`, bold: true, size: 28 })], { before: 60 }),
  new Paragraph({ heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER, spacing: { after: 160 },
    children: [new TextRun({ text: name, bold: true, size: 28 })] }),
];

const bullet = (text, level = 0) => new Paragraph({
  spacing: { after: 70, line: 290 }, alignment: AlignmentType.JUSTIFIED,
  numbering: { reference: "bullets", level },
  children: Array.isArray(text) ? text : [new TextRun(text)],
});
const num = (text) => new Paragraph({
  spacing: { after: 70, line: 290 }, alignment: AlignmentType.JUSTIFIED,
  numbering: { reference: "nums", level: 0 },
  children: Array.isArray(text) ? text : [new TextRun(text)],
});

const figure = (file, caption, widthIn = 6.1) => {
  const w = Math.round(widthIn * 96);
  // keep aspect by reading nothing; approximate heights per figure set below
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 40 },
      children: [new ImageRun({ type: "png", data: img(file),
        transformation: { width: w, height: Math.round(w * (file === "metrics.png" ? 0.5 :
          file === "architecture.png" ? 0.565 : file === "future_event.png" ? 0.35 : 0.38)) } })] }),
    center([new TextRun({ text: caption, italics: true, size: 18 })], { after: 160 }),
  ];
};

// table builder
const shade = (fill) => ({ type: ShadingType.CLEAR, fill, color: "auto" });
const tcell = (text, { w, header = false, bg, bold = false, align } = {}) => new TableCell({
  width: { size: w, type: WidthType.DXA },
  shading: header ? shade(NAVY) : (bg ? shade(bg) : undefined),
  margins: { top: 60, bottom: 60, left: 90, right: 90 },
  children: [new Paragraph({ alignment: align || AlignmentType.LEFT, spacing: { after: 0, line: 260 },
    children: (Array.isArray(text) ? text : [text]).map((t) =>
      new TextRun({ text: String(t), bold: header || bold, color: header ? "FFFFFF" : "000000", size: 19 })) })],
});
const table = (widths, rows) => new Table({
  columnWidths: widths,
  width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
  rows: rows.map((r, i) => new TableRow({ tableHeader: i === 0,
    children: r.map((c, j) => tcell(c, { w: widths[j], header: i === 0, bg: i % 2 === 0 && i > 0 ? "EDF2F7" : undefined })) })),
  borders: {
    top: { style: BorderStyle.SINGLE, size: 4, color: "B0B7C3" },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: "B0B7C3" },
    left: { style: BorderStyle.SINGLE, size: 4, color: "B0B7C3" },
    right: { style: BorderStyle.SINGLE, size: 4, color: "B0B7C3" },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: "CBD5E0" },
    insideVertical: { style: BorderStyle.SINGLE, size: 2, color: "CBD5E0" },
  },
});
const tcap = (t) => center([new TextRun({ text: t, italics: true, size: 18 })], { before: 100, after: 60 });

const code = (lines) => new Paragraph({
  spacing: { before: 60, after: 120 }, shading: shade("F4F5F7"),
  border: { top: { style: BorderStyle.SINGLE, size: 2, color: "D9DEE6" }, bottom: { style: BorderStyle.SINGLE, size: 2, color: "D9DEE6" },
    left: { style: BorderStyle.SINGLE, size: 2, color: "D9DEE6" }, right: { style: BorderStyle.SINGLE, size: 2, color: "D9DEE6" } },
  children: lines.flatMap((l, i) => [new TextRun({ text: l, font: "Consolas", size: 16, break: i === 0 ? 0 : 1 })]),
});

const gap = (n = 1) => Array.from({ length: n }, () => new Paragraph({ children: [] }));
const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

// ---------- front matter ----------
const TITLE = "FreshGuard: Freshness-Aware Adaptive Retrieval-Augmented Generation with a Learned Retrieval Gate and Statement Verification";
const team = [
  ["J. Sai Teja", "23EG106C22"],
  ["P. Supriya", "23EG106C52"],
  ["Sk. Mubeena", "23EG106C58"],
  ["U. Venkatesh", "23EG106C63"],
];

function titlePage() {
  return [
    center([new TextRun({ text: "A Mini Project Report on", italics: true, size: 26 })], { before: 200, after: 220 }),
    center([new TextRun({ text: TITLE, bold: true, size: 30 })], { after: 260 }),
    center([new TextRun({ text: "Submitted to Anurag University in partial fulfilment of the requirements", italics: true, size: 22 })]),
    center([new TextRun({ text: "for the award of the Degree of", italics: true, size: 22 })], { after: 120 }),
    center([new TextRun({ text: "Bachelor of Technology in Artificial Intelligence", bold: true, size: 24 })], { after: 220 }),
    center([new TextRun({ text: "Submitted by", size: 22 })], { after: 120 }),
    ...team.map(([n, r]) => center([new TextRun({ text: `${n}  —  ${r}`, bold: true, size: 23 })], { after: 60 })),
    center([new TextRun({ text: "Under the Guidance of", italics: true, size: 22 })], { before: 200, after: 80 }),
    center([new TextRun({ text: "Mohammad Arshad", bold: true, size: 24 })]),
    center([new TextRun({ text: "Assistant Professor", size: 22 })], { after: 260 }),
    center([new TextRun({ text: "DEPARTMENT OF ARTIFICIAL INTELLIGENCE", bold: true, size: 24 })], { before: 160 }),
    center([new TextRun({ text: "SCHOOL OF ENGINEERING", bold: true, size: 24 })]),
    center([new TextRun({ text: "ANURAG UNIVERSITY", bold: true, size: 24 })]),
    center([new TextRun({ text: "2023–2027", bold: true, size: 24 })]),
    pageBreak(),
  ];
}

function certificate() {
  return [
    center([new TextRun({ text: "ANURAG UNIVERSITY", bold: true, size: 26 })], { before: 120 }),
    center([new TextRun({ text: "DEPARTMENT OF ARTIFICIAL INTELLIGENCE", bold: true, size: 24 })], { after: 160 }),
    center([new TextRun({ text: "CERTIFICATE", bold: true, size: 28 })], { after: 200 }),
    P([new TextRun("This is to certify that the project report titled "),
      new TextRun({ text: TITLE, bold: true }),
      new TextRun(" is being submitted by J. Sai Teja (23EG106C22), P. Supriya (23EG106C52), Sk. Mubeena (23EG106C58) and U. Venkatesh (23EG106C63), in IV B.Tech I semester, Artificial Intelligence, and is a record of bonafide work carried out by them. The results embodied in this report have not been submitted to any other University for the award of any degree.")]),
    ...gap(1),
    table([4600, 4600], [
      ["Student's Name", "Signature"],
      ["1. J. Sai Teja", "1."],
      ["2. P. Supriya", "2."],
      ["3. Sk. Mubeena", "3."],
      ["4. U. Venkatesh", "4."],
    ]),
    ...gap(2),
    table([3066, 3066, 3068], [
      ["Guide", "Head of Department", "External Examiner"],
      ["Mohammad Arshad\nAssistant Professor", "Dr. M. Trupti\nProfessor", " "],
    ]),
    pageBreak(),
  ];
}

function declaration() {
  return [
    center([new TextRun({ text: "DECLARATION", bold: true, size: 28 })], { before: 120, after: 180 }),
    P([new TextRun("We hereby declare that the Mini Project report entitled "),
      new TextRun({ text: `“${TITLE}”`, bold: true }),
      new TextRun(", submitted to the Department of Artificial Intelligence, Anurag University, in partial fulfilment of the requirements for the award of the degree of Bachelor of Technology in Artificial Intelligence, is a record of original work carried out by us under the guidance of Mohammad Arshad, Assistant Professor, Department of Artificial Intelligence. The results embodied in this report have not been submitted to any other university or institute for the award of any degree or diploma. Wherever the work of others, including datasets, pretrained models and publicly available code, has been used, it has been duly acknowledged and cited. The results reported are genuine and reproducible from the submitted code.")]),
    P([new TextRun({ text: "Use of AI tools: ", bold: true }),
      new TextRun("We declare that we have used an AI assistant (Claude) during this project to help with software development and with drafting and editing parts of this report. All experiments were run by us, and every result reported here was checked against the logged output of the system. We take full responsibility for the content of this report.")]),
    ...gap(1),
    P([new TextRun("Place: Hyderabad")]),
    P([new TextRun("Date: ______________")]),
    ...gap(1),
    table([4000, 2600, 2600], [
      ["Name of the Student", "Roll No.", "Signature"],
      ...team.map(([n, r]) => [n, r, " "]),
    ]),
    pageBreak(),
  ];
}

function acknowledgement() {
  return [
    center([new TextRun({ text: "ACKNOWLEDGEMENT", bold: true, size: 28 })], { before: 120, after: 180 }),
    P("We are grateful to Prof. Archana Mantri, Vice-Chancellor, Anurag University, for the facilities the University made available to us and for the environment in which this project could be carried out."),
    P("We thank Dr. V. Vijay Kumar, Dean, School of Engineering, for his encouragement throughout our course of study."),
    P("We sincerely thank Dr. M. Trupti, Professor and Head of the Department of Artificial Intelligence, for providing the resources and support we needed to complete this work."),
    P("We are especially indebted to our guide, Mohammad Arshad, Assistant Professor, Department of Artificial Intelligence, Anurag University. His questions during the reviews shaped how we tested the system, and his feedback on each draft improved this report."),
    P("We also thank the developers of the open-source libraries this project is built on, in particular scikit-learn, FAISS, LangChain, FastAPI, Uvicorn and the web-search and large-language-model providers Tavily and xAI (Grok), whose free tiers made the live-ingestion experiments possible."),
    P("An AI assistant (Claude) was used during this project to help with software development and with drafting and editing parts of this report. All experiments were run by us, and every result reported here was checked against the logged output of the system."),
    ...gap(1),
    ...team.map(([n, r]) => center([new TextRun({ text: `${n} (${r})`, italics: true, size: 21 })], { after: 40 })),
    pageBreak(),
  ];
}

function abstract() {
  return [
    center([new TextRun({ text: "ABSTRACT", bold: true, size: 28 })], { before: 120, after: 180 }),
    P("Large language models answer questions fluently, but they answer confidently even when their knowledge is out of date, and they can state facts that are no longer true. Static retrieval-augmented generation (RAG) is one remedy, but it retrieves for every question — wasting latency and cost on questions the model already knows — and it still does not check the answer the model finally writes."),
    P("This project builds FreshGuard, a freshness-aware, adaptive RAG system that decides when retrieval is needed, retrieves fresh evidence only when it is, generates a grounded answer, and then verifies every statement against that evidence before the user sees it. A learned freshness estimator (a logistic-regression classifier over TF-IDF and temporal features of the query) predicts the probability that the model's parametric knowledge is stale, and a decision policy retrieves only when that probability — or low model confidence — crosses an admin-configurable threshold. Evidence is retrieved from a local FAISS vector store of TF-IDF embeddings and ranked by relevance and recency. When the knowledge base does not contain the answer — for example, a question about a future or very recent event — FreshGuard fetches fresh evidence live from Tavily web search and the xAI Grok model, embeds it into the knowledge base, and then retrieves and verifies it like any other document; when no live source is configured it abstains honestly instead of answering from a stale, year-mismatched document. A statement-level verification stage labels each sentence of the answer as supported, unsupported or contradicted, drops or flags the bad ones, and computes a calibrated confidence score from three signals."),
    P("On a controlled benchmark, the retrieval-decision (freshness) accuracy was 1.00, factuality 1.00, and hallucination precision, recall and F1 were each 1.00, with an expected calibration error of 0.10. The system is fully offline and deterministic by default: a suite of 45 unit, end-to-end, API and evaluation tests passed identically across ten consecutive runs. The main limitation is that live ingestion depends on an external key, and the confidence gate is an uncalibrated default; calibrating it on a larger labelled set is the most important next step."),
    P([new TextRun({ text: "Keywords: ", bold: true }),
      new TextRun("Adaptive RAG, Knowledge Freshness, Hallucination Detection, Statement Verification, Learned Retrieval Gate, Live Ingestion, Tavily, Grok, Calibration.")]),
    pageBreak(),
  ];
}

function toc() {
  return [
    center([new TextRun({ text: "TABLE OF CONTENTS", bold: true, size: 28 })], { after: 160 }),
    new TableOfContents("Contents", { hyperlink: true, headingStyleRange: "1-3" }),
    pageBreak(),
  ];
}

function listOfFigures() {
  const rows = [
    ["Fig. No.", "Title"],
    ["1.1", "High-level block diagram of FreshGuard: the two gates around generation"],
    ["4.1", "System architecture: web UI, FastAPI backend, pipeline stages and stores"],
    ["4.2", "The adaptive retrieval gate: retrieve only when stale or low-confidence"],
    ["4.3", "Live ingestion on a knowledge-base miss (Tavily + Grok)"],
    ["6.1", "Evaluation metrics on the controlled benchmark"],
    ["6.2", "Future-event query: before and after live ingestion"],
  ];
  return [
    center([new TextRun({ text: "LIST OF FIGURES", bold: true, size: 28 })], { after: 160 }),
    table([1500, 7700], rows),
    pageBreak(),
  ];
}
function listOfTables() {
  const rows = [
    ["Table No.", "Title"],
    ["2.1", "Summary of related work and how this project builds on it"],
    ["3.1", "Software requirements and development environment"],
    ["4.1", "Seed knowledge base used by the project"],
    ["4.2", "Functional requirement to component mapping (FR1–FR22)"],
    ["4.3", "Modules of the system"],
    ["5.1", "Pipeline configuration used in the experiments"],
    ["6.1", "Evaluation results on the controlled benchmark"],
    ["6.2", "Behaviour on a future-event query with and without live sources"],
  ];
  return [
    center([new TextRun({ text: "LIST OF TABLES", bold: true, size: 28 })], { after: 160 }),
    table([1600, 7600], rows),
    pageBreak(),
  ];
}
function abbreviations() {
  const rows = [["Abbreviation", "Expansion"],
    ["AI", "Artificial Intelligence"], ["API", "Application Programming Interface"],
    ["BM25", "Best Match 25 (keyword ranking function)"], ["CPU", "Central Processing Unit"],
    ["ECE", "Expected Calibration Error"], ["FAISS", "Facebook AI Similarity Search"],
    ["FR", "Functional Requirement"], ["HTTP", "HyperText Transfer Protocol"],
    ["JSON", "JavaScript Object Notation"], ["KB", "Knowledge Base"],
    ["LLM", "Large Language Model"], ["NLI", "Natural Language Inference"],
    ["RAG", "Retrieval-Augmented Generation"], ["REST", "Representational State Transfer"],
    ["SQL", "Structured Query Language"], ["TF-IDF", "Term Frequency–Inverse Document Frequency"],
    ["UI", "User Interface"], ["xAI", "x.AI (provider of the Grok model)"]];
  return [
    center([new TextRun({ text: "LIST OF ABBREVIATIONS", bold: true, size: 28 })], { after: 160 }),
    table([3200, 6000], rows),
    pageBreak(),
  ];
}

// ---------- chapters ----------
function chapter1() {
  return [
    ...chapTitle("1", "INTRODUCTION"),
    H2("1.1 Background / Overview"),
    P("People increasingly ask large language models (LLMs) factual questions — who holds an office now, what the latest price or model is, what happened in a recent event. An LLM answers from a fixed snapshot of the world taken at its training cutoff, so for anything that has changed since then it can be confidently wrong. This is the knowledge-freshness problem: the model's fluency does not tell the user whether its knowledge is still current."),
    P("Retrieval-augmented generation (RAG) is the standard remedy. Instead of answering from memory alone, the system retrieves passages from a trusted collection and asks the model to answer from them. Classic RAG, however, has two weaknesses that this project targets. First, it retrieves for every query, even for stable facts the model already knows — spending latency and cost where none is needed. Second, it grounds the prompt but does not check the answer the model finally writes: the model can still misread a passage, combine two incorrectly, or state something the evidence does not support."),
    P("FreshGuard addresses both weaknesses by placing two learned gates around generation. A freshness gate before generation decides whether retrieval is needed at all, and a verification gate after generation checks every statement of the answer against the retrieved evidence. When the knowledge base cannot answer a question — the typical case for a future or very recent event — FreshGuard reaches out to live sources (a web-search API and a hosted LLM), ingests what they return into its own knowledge base, and then retrieves and verifies that fresh evidence like any other document. Figure 1.1 shows this flow at the highest level."),
    ...figure("gate.png", "Fig. 1.1 High-level block diagram of FreshGuard: a freshness gate before generation decides whether to retrieve, and a verification gate after generation checks the answer before it is shown."),
    H2("1.2 Motivation"),
    P("The motivation was not to build another chatbot, but a system that treats both timeliness and faithfulness as first-class, measurable properties of an answer. Three observations shaped this goal."),
    bullet("First, the cost of a confidently stale answer is asymmetric. Telling a user the current interest rate, the latest flagship chipset, or the winner of a recent tournament from year-old memory is worse than saying the information is not available, because the user may act on it. A system that retrieves exactly when its knowledge is likely stale, and that says so when it cannot answer, is more trustworthy than one that always answers."),
    bullet("Second, retrieving for every query is wasteful. Most questions a user asks (the speed of light, the capital of a country, how photosynthesis works) are timeless; a well-trained model answers them directly. A learned gate that skips retrieval for such questions saves latency and cost without losing accuracy."),
    bullet("Third, grounding is not the same as verifying. A passage in the prompt constrains what the model sees but not what it writes. Checking each statement against the evidence, and showing the user which statements were supported, turns a black-box answer into one the reader can audit."),
    H2("1.3 Problem Statement"),
    P("Existing question-answering systems built on LLMs and basic RAG (1) answer from stale parametric memory without signalling that the knowledge may be out of date, (2) retrieve for every query regardless of need, (3) do not verify the generated answer against the evidence before showing it, and (4) answer anyway when the knowledge base does not contain the answer, often from a superficially similar but wrong-period document. In particular, the project addresses:"),
    num("Stale answers without warning: time-sensitive questions are answered from a fixed training snapshot with no freshness signal."),
    num("Indiscriminate retrieval: retrieval runs for every query, adding latency and cost even for timeless facts."),
    num("Unverified answers: the model's answer is shown as-is, even when a statement is not supported by, or contradicts, the evidence."),
    num("Guessing instead of abstaining: when the knowledge base has no answer, the system returns a plausible but wrong document (for example, a 2022 result for a 2026 question) rather than fetching fresh evidence or saying the answer is unavailable."),
    P("Hence there is a need for a system that (1) estimates, per query, how likely the model's knowledge is stale, (2) retrieves only when that estimate or low model confidence calls for it, (3) when local evidence is missing, fetches and ingests fresh evidence from live sources, (4) verifies every statement against the evidence and reports a calibrated confidence, and (5) abstains honestly when no reliable evidence can be found."),
    H2("1.4 Objectives"),
    bullet("Build a query-preprocessing and classification stage, and a learned freshness estimator that outputs a calibrated probability that the model's knowledge is stale (not a fixed time-decay rule)."),
    bullet("Implement an adaptive retrieval decision that retrieves only when the freshness score is high or model confidence is low, with admin-configurable thresholds."),
    bullet("Implement evidence retrieval over a local FAISS vector store with relevance-and-recency ranking, plus live ingestion from Tavily web search and the xAI Grok model on a knowledge-base miss."),
    bullet("Implement statement-level verification (supported / unsupported / contradicted), corrective action, and a confidence score combining model, evidence and freshness signals."),
    bullet("Provide a web interface with a freshness badge, confidence meter, per-statement verification, cited sources, history and feedback, and an admin panel for thresholds, the knowledge base and the live sources."),
    bullet("Evaluate retrieval decisions, factuality, hallucination detection and calibration, and demonstrate determinism; report results honestly, including limitations."),
    H2("1.5 Scope and Limitations"),
    P("FreshGuard answers natural-language questions in English over a seeded knowledge base that the admin can extend, and over evidence it fetches live when the base has none. It is a research prototype and does not give professional advice. Possible uses include a freshness-aware question-answering assistant, a teaching example of adaptive RAG and hallucination detection, and a reusable testbed for calibration research."),
    P("The scope is deliberately limited. The default LLM is an offline grounded extractive generator chosen for determinism and reproducibility; hosted models (OpenAI, Anthropic, Grok) are used only when an API key is present. Live ingestion depends on an external key (Tavily and/or xAI); without one the system runs fully offline and abstains on knowledge-base misses. The verification stage uses lightweight, deterministic linguistic rules rather than a large NLI model, and it checks what an answer says, not what it leaves out. The evaluation is small and controlled rather than large and statistical: a labelled set of twelve queries, a hallucination set of eight claims, and a demonstrated live-ingestion flow. The numbers reported are observations from logged runs, not population estimates."),
    H2("1.6 Organization of the Report"),
    P("Chapter 2 reviews the literature on retrieval-augmented generation, knowledge freshness and temporal QA, adaptive retrieval, and hallucination detection and calibration, and identifies the research gap. Chapter 3 analyses existing approaches, examines feasibility and lists the hardware and software requirements. Chapter 4 describes the knowledge base, the overall design and the methodology of every pipeline stage, including the freshness estimator, the decision policy, retrieval with live ingestion, verification, confidence and the functional-requirement mapping. Chapter 5 covers the implementation: technologies, configuration, procedures, key code and testing. Chapter 6 presents the results: the controlled benchmark, the live-ingestion demonstration, determinism, error analysis and ethical considerations. Chapter 7 concludes and lists future enhancements. The appendix contains the setup and run guide, the REST API reference and the environment configuration."),
    pageBreak(),
  ];
}

function chapter2() {
  const rows = [
    ["Area", "Representative approach", "Limitation", "In this project"],
    ["Retrieval-augmented generation", "Retrieve, then generate (Lewis et al.)", "Retrieves for every query; no check on the final answer", "Adaptive gate retrieves only when needed; every statement verified"],
    ["Knowledge freshness / temporal QA", "Time-decay or fixed-cutoff heuristics", "Fixed rules ignore the query; cannot learn what is time-sensitive", "Learned logistic-regression estimator over TF-IDF + temporal features"],
    ["Adaptive / self-RAG", "Let the model decide when to retrieve", "Decision hidden inside the LLM; not inspectable or tunable", "Explicit, thresholded decision policy; admin-configurable"],
    ["Hybrid retrieval & ranking", "Dense + keyword with recency", "Relevance alone can surface a stale, wrong-period document", "Relevance × recency ranking + a year gate that rejects period mismatch"],
    ["Live / web-augmented RAG", "Query a web API at answer time", "Results used once, not persisted or verified", "Tavily + Grok results embedded, persisted, re-retrieved and verified"],
    ["Hallucination detection", "NLI / atomic-fact / self-consistency scoring", "Used as offline evaluation, not a deployed gate", "Per-statement labels drive correction and a shown confidence"],
    ["Confidence calibration", "Temperature scaling, ECE reporting", "Reported, rarely acted on in a running system", "Three-signal confidence, ECE measured and reported (0.10)"],
  ];
  return [
    ...chapTitle("2", "LITERATURE SURVEY"),
    H2("2.1 Review of Related Work"),
    H3("2.1.1 Retrieval-Augmented Generation"),
    P("Retrieval-augmented generation grounds a generator on passages selected by a retriever, which improves factuality on knowledge-intensive tasks. Most assistants follow this design: split documents into chunks, embed each chunk, retrieve the chunks closest to the question, and let an LLM write the answer. The architecture has two well-known failure points that this project targets — the retriever can return the wrong passages, and the generator can depart from the passages it was given — and, for freshness specifically, basic RAG retrieves for every query whether or not the model already knows the answer."),
    H3("2.1.2 Knowledge Freshness and Temporal Question Answering"),
    P("A body of work studies how model knowledge decays over time and how to detect when a question is time-sensitive. Simple systems use a fixed cutoff or a time-decay weight on documents. These rules are easy to implement but do not depend on the query: they cannot tell that \"the speed of light\" is timeless while \"the current interest rate\" is not. FreshGuard instead learns a freshness estimator from a labelled set of queries, so the decision to retrieve is driven by what the query asks rather than by a global rule."),
    H3("2.1.3 Adaptive and Self-Reflective Retrieval"),
    P("Recent systems let the model decide, token by token, whether to retrieve, or train the model to emit retrieval actions. These adaptive approaches reduce needless retrieval, but the decision is hidden inside the model and is neither inspectable nor directly tunable. FreshGuard keeps the decision outside the model as an explicit, thresholded policy over a learned freshness score and a model-confidence estimate, which an administrator can tune and which the evaluation can score on its own."),
    H3("2.1.4 Hybrid Retrieval, Ranking and Recency"),
    P("Dense retrieval ranks passages by vector similarity, capturing meaning even when wording differs; keyword retrieval (BM25) is exact where dense retrieval is fuzzy. Combining them, and re-ranking by recency so that fresher documents win among equally relevant ones, is standard practice. FreshGuard uses a FAISS index over TF-IDF embeddings and a relevance × recency score, and adds a year gate: if the query names a year that none of the retrieved documents cover, the local evidence is treated as insufficient, which prevents a 2022 document from answering a 2026 question."),
    H3("2.1.5 Live and Web-Augmented Generation"),
    P("Web-augmented systems call a search API at answer time to bring in information beyond the index. Typically the results are used once, for the current answer, and discarded. FreshGuard instead treats a live source (Tavily web search and the xAI Grok model) as a way to fill a gap permanently: the fetched passages are embedded into the FAISS index, persisted into the knowledge base, re-retrieved, and passed through the same verification stage as any seed document, so a question asked once can be answered from the index thereafter."),
    H3("2.1.6 Hallucination Detection and Calibration"),
    P("Natural-language-inference entailment, atomic-fact decomposition and self-consistency sampling are established ways to detect unsupported statements, and expected calibration error (ECE) measures whether a model's confidence matches its accuracy. These are almost always used as offline evaluation tools. FreshGuard runs a lightweight, deterministic version of statement checking inside the pipeline — labelling each statement supported, unsupported or contradicted, acting on the labels, and folding the result into a confidence it shows the user — and reports ECE as a first-class metric."),
    H2("2.2 Summary Table"),
    tcap("Table 2.1 Summary of related work and how this project builds on it"),
    table([2100, 2500, 2300, 2300], rows),
    H2("2.3 Research Gap Identified"),
    P("Across this literature, freshness estimation, the retrieval decision, live augmentation and answer verification are studied separately, and the first three rarely act on the query while the last is rarely deployed inside a running system. The gaps this project addresses are:"),
    bullet("A learned, query-specific freshness estimator that outputs a calibrated staleness probability rather than a fixed time rule."),
    bullet("An explicit, tunable retrieval decision that skips retrieval for fresh questions and triggers it for stale or low-confidence ones."),
    bullet("Live ingestion that fills a knowledge-base miss from a web-search API and a hosted LLM, then embeds, persists, re-retrieves and verifies the fetched evidence."),
    bullet("A year gate and honest abstention that stop the system from answering a time-specific question from a wrong-period document."),
    bullet("Per-statement verification and a shown, calibrated confidence, with ECE measured rather than assumed."),
    bullet("A fully offline, deterministic default so the whole pipeline is reproducibly testable without any network or key."),
    pageBreak(),
  ];
}

function chapter3() {
  const sw = [
    ["Layer", "Software"],
    ["Operating system", "Windows 10/11, Linux or macOS"],
    ["Language runtime", "Python 3.11+ (backend, pipeline, experiments, tests)"],
    ["Backend framework", "FastAPI with Uvicorn; Pydantic for request/response models"],
    ["Retrieval & models", "scikit-learn (TF-IDF, LogisticRegression); FAISS (faiss-cpu); LangChain-core / LangChain-community vector store; joblib; NumPy; SciPy"],
    ["Live sources", "Tavily web-search API and xAI Grok (OpenAI-compatible endpoint) via requests; python-dotenv for keys"],
    ["Storage", "SQLite (standard library) for history, feedback, config and the knowledge base"],
    ["Frontend", "Single-page HTML + CSS + JavaScript served by FastAPI"],
    ["Testing & tools", "pytest; httpx; Git; a modern web browser"],
  ];
  return [
    ...chapTitle("3", "SYSTEM ANALYSIS AND REQUIREMENTS"),
    H2("3.1 Existing Approaches and their Drawbacks"),
    P("Existing LLM and basic-RAG question-answering systems share a common set of weaknesses when freshness matters:"),
    bullet([new TextRun({ text: "Stale answers with no signal. ", bold: true }), new TextRun("An LLM answers time-sensitive questions from a fixed snapshot and gives no indication that its knowledge may be out of date.")]),
    bullet([new TextRun({ text: "Retrieval for every query. ", bold: true }), new TextRun("Basic RAG retrieves even for timeless facts, adding latency and cost where the model already knows the answer.")]),
    bullet([new TextRun({ text: "Answers shown unchecked. ", bold: true }), new TextRun("Nothing verifies the final answer against the evidence before the user reads it.")]),
    bullet([new TextRun({ text: "Wrong-period evidence. ", bold: true }), new TextRun("Relevance-only retrieval can return a document from the wrong year, which is then used as if it answered the question.")]),
    bullet([new TextRun({ text: "No graceful fallback. ", bold: true }), new TextRun("When the base has no answer, the system guesses rather than fetching fresh evidence or abstaining.")]),
    P("FreshGuard answers each drawback directly: a learned freshness estimator and a shown badge (stale answers without a signal), an adaptive decision policy (retrieval for every query), a statement-level verification gate (answers shown unchecked), a year gate plus relevance × recency ranking (wrong-period evidence), and live ingestion with honest abstention (no graceful fallback)."),
    H2("3.2 Feasibility Study"),
    H3("3.2.1 Technical Feasibility"),
    P("Every component is available as mature open-source software or as a hosted API. The freshness estimator, TF-IDF embeddings and the FAISS index run locally on a CPU with no GPU; the learned model is a few hundred kilobytes. SQLite stores history, feedback, configuration and the knowledge base in one on-disk file, so no database server is needed. The only external dependencies are optional: a web-search key (Tavily) and an LLM key (xAI Grok or another provider), each behind an adapter so the system degrades to a fully offline grounded generator when a key is absent."),
    H3("3.2.2 Economic Feasibility"),
    P("The project was built and evaluated at essentially zero monetary cost. All local models and libraries are free, the development machines were ordinary laptops without GPUs, and the live sources were used through free tiers. No component requires a paid service to operate; a larger deployment would pay only for optional LLM and search tokens and for ordinary hosting."),
    H3("3.2.3 Operational Feasibility"),
    P("The system runs on a standard Windows, Linux or macOS machine with two commands: one to train the freshness estimator once, and one to serve the API and web interface. Users interact through a browser; every answer explains itself through its freshness badge, confidence meter, per-statement verification and cited sources, so no training in the system's internals is needed to judge an answer. The index, history and configuration are plain local files, which makes backup and inspection simple."),
    H2("3.3 Hardware Requirements"),
    bullet([new TextRun({ text: "Minimum: ", bold: true }), new TextRun("a dual-core CPU, 4 GB RAM and about 1 GB of free disk space for the Python environment, the index and the artifacts.")]),
    bullet([new TextRun({ text: "Recommended: ", bold: true }), new TextRun("a quad-core CPU and 8 GB RAM for faster test and evaluation runs.")]),
    bullet([new TextRun({ text: "GPU: ", bold: true }), new TextRun("not required; all local components run on the CPU.")]),
    bullet([new TextRun({ text: "Network: ", bold: true }), new TextRun("needed only for optional live ingestion (Tavily / Grok); the system runs fully offline without it.")]),
    H2("3.4 Software Requirements"),
    tcap("Table 3.1 Software requirements and development environment"),
    table([2600, 6600], sw),
    H2("3.5 Client (User) Requirements"),
    P("From the user's point of view, the system must:"),
    bullet("Ask questions in plain English and receive an answer with a freshness indicator, a confidence score and cited sources."),
    bullet("See clearly whether the answer was retrieved or answered directly, and whether live sources were used to fetch it."),
    bullet("Inspect the per-statement verification so a claim can be checked against its evidence."),
    bullet("Keep a history of questions and give feedback (rating or report) on an answer."),
    bullet("As an administrator, change the decision thresholds, add or remove knowledge-base documents, choose the LLM provider, enable or disable live ingestion, and run the evaluation — all without reading code."),
    pageBreak(),
  ];
}

function chapter4() {
  const kb = [
    ["Group", "Example documents", "Count", "Used for"],
    ["Time-sensitive facts", "2022 World Cup result; 2025 Fed policy rate; 2025 flagship chipset; 2024 Olympics host; 2025 EV battery price; 2025 LLM context windows", "6", "Stale queries that must trigger retrieval and verification"],
    ["Timeless facts", "Speed of light; photosynthesis; capital of Australia; Pythagorean theorem; boiling point of water; human heart", "6", "Fresh queries answered directly without retrieval"],
    ["Live-ingested (runtime)", "Documents fetched from Tavily / Grok on a KB miss, each dated and cited", "variable", "Future / very recent questions not in the seed base"],
  ];
  const fr = [
    ["FR", "Requirement", "Component"],
    ["FR1", "Natural-language query input via web UI", "frontend, api /query"],
    ["FR2", "Clean / normalise / classify query", "preprocess.py"],
    ["FR3", "Estimate knowledge-staleness likelihood", "freshness.py"],
    ["FR4", "Numeric freshness score (learned, not fixed decay)", "freshness.py (LogisticRegression)"],
    ["FR5", "Adaptive retrieval decision", "decision.py"],
    ["FR6", "Skip retrieval when fresh", "decision.py / pipeline.py"],
    ["FR7", "Evidence retrieval (vector store / live sources)", "retrieval.py, live_sources.py"],
    ["FR8", "Rank + filter by relevance and recency", "retrieval.py"],
    ["FR9", "Grounded answer generation", "llm.py"],
    ["FR10", "Statement-level verification", "verify.py"],
    ["FR11", "Hallucination labelling (supported / unsupported / contradicted)", "verify.py"],
    ["FR12", "Corrective action / warn", "verify.py, pipeline.py"],
    ["FR13", "Confidence score (model + evidence + freshness)", "confidence.py"],
    ["FR14", "Evidence highlighting per claim", "pipeline output + frontend"],
    ["FR15", "Source citation (link, title, date)", "retrieval metadata + frontend"],
    ["FR16", "Final answer + confidence + freshness + links", "pipeline.py, frontend"],
    ["FR17", "Query-history storage", "storage.py"],
    ["FR18", "User feedback (rate / report)", "storage.py, api /feedback"],
    ["FR19", "Admin config (thresholds, sources, LLM)", "config.py, storage.py, api /config"],
    ["FR20", "Knowledge-base management", "storage.py, retrieval.py, api /kb"],
    ["FR21", "Performance evaluation module", "evaluation.py, eval/run_eval.py"],
    ["FR22", "Error handling + graceful fallback", "pipeline.py (try/except, abstention)"],
  ];
  const mods = [
    ["Module", "Responsibility", "Input → Output"],
    ["preprocess.py", "Clean, normalise and classify the query", "Query → cleaned query + category"],
    ["freshness.py", "Learned staleness probability + model self-confidence", "Query → freshness score, confidence"],
    ["decision.py", "Retrieve-or-skip policy over thresholds", "Scores → decision + reason"],
    ["retrieval.py", "FAISS search, relevance × recency ranking, KB management", "Query → ranked documents"],
    ["live_sources.py", "Tavily web search and Grok fetch on a KB miss", "Query → new documents + sources used"],
    ["llm.py", "Pluggable generation (grounded / OpenAI / Anthropic / Grok)", "Query + evidence → answer text"],
    ["verify.py", "Per-statement support labels and correction", "Answer + evidence → labelled statements"],
    ["confidence.py", "Confidence from support, relevance and freshness", "Signals → 0–100 confidence"],
    ["pipeline.py", "Orchestration, highlighting, citations, error handling", "Query → full response"],
    ["storage.py", "SQLite history / feedback / config / KB", "Response / updates → persisted rows"],
    ["evaluation.py", "Freshness, factuality, hallucination, calibration", "Runs → metrics"],
    ["api/main.py", "FastAPI endpoints; serves the frontend", "HTTP request → JSON response"],
  ];
  return [
    ...chapTitle("4", "DATASET, DESIGN AND METHODOLOGY"),
    H2("4.1 Overall Pipeline and System Architecture"),
    P("FreshGuard is organised as a web application on top of a question-answering pipeline. A single-page web interface talks to a FastAPI backend over HTTP; the backend runs the pipeline stages, each implemented as a separate module, and stores documents in a FAISS index and records in SQLite. The default language model is a deterministic offline grounded generator; hosted models and the live sources are called through adapters only when a key is present. Figure 4.1 shows the components and how they connect."),
    ...figure("architecture.png", "Fig. 4.1 System architecture: the web UI and FastAPI backend, the pipeline stages (preprocess, freshness, decision, retrieval with live ingestion, generation, verification, confidence), and the FAISS, SQLite and model-artifact stores.", 6.3),
    P("The pipeline runs for every query: the query is cleaned and classified; the freshness estimator produces a staleness probability and a model-confidence estimate; the decision policy chooses whether to retrieve; if it retrieves, the FAISS store is searched and, on a knowledge-base miss, live sources are consulted and their results embedded and persisted; the chosen LLM generates a grounded answer; the verification stage labels and corrects the statements; and the confidence stage combines the signals. The final response carries the answer, a freshness badge, a confidence score, the cited sources and the per-statement verdicts."),
    H2("4.2 Dataset"),
    P("Because FreshGuard is a pipeline rather than a trained end-to-end model, its data are of three kinds: a seed knowledge base, a synthesised training set for the freshness estimator, and labelled evaluation sets. The seed knowledge base (Table 4.1) mixes time-sensitive documents, which a stale model would get wrong, with stable background documents; each document carries a fixed date so retrieval and recency ranking are deterministic for tests."),
    tcap("Table 4.1 Seed knowledge base used by the project"),
    table([2100, 4100, 1000, 2000], kb),
    P("The freshness estimator is trained on a synthesised, labelled set of queries (375 examples) in which time-sensitive phrasings (explicit years, \"latest\", \"current\", \"now\", \"this year\") are labelled stale and timeless phrasings are labelled fresh. The evaluation uses twelve labelled queries (six fresh, six stale) with gold answer keywords and gold retrieval decisions, and a hallucination set of eight sentence–evidence pairs, half of them deliberate hallucinations. No document, query or label is shared between training the estimator and evaluation; the labels were fixed before the runs, and every reported number is read from a logged results file."),
    H2("4.3 Query Preprocessing and Classification"),
    P("An incoming query is trimmed, its whitespace normalised, and its category inferred from simple lexical cues (sports, science, finance, technology, general). An empty query after cleaning triggers a graceful low-confidence fallback. The cleaned query and its category flow to the freshness estimator and, later, to generation and display."),
    H2("4.4 Learned Freshness Estimation (FR3, FR4)"),
    P("The freshness estimator is a scikit-learn pipeline: a TF-IDF vectoriser over word n-grams of the query, feeding a logistic-regression classifier that outputs the probability that the model's parametric knowledge is outdated for this query. Temporal cues — explicit years, words such as \"latest\", \"current\", \"now\" and \"this year\" — carry most of the signal, but because the model is learned rather than rule-based it generalises to phrasings not seen verbatim in training. The output is a calibrated probability in [0, 1]; it is explicitly not a fixed time-decay rule. A monotone model-confidence estimate is derived from the freshness score: static-looking queries yield high confidence, time-sensitive ones low confidence."),
    H2("4.5 Adaptive Retrieval Decision (FR5, FR6)"),
    P("The decision policy retrieves when the knowledge looks stale or the model is not confident enough to answer directly, and otherwise skips retrieval to save latency and cost. Formally, it retrieves if the freshness score is at least the freshness threshold (default 0.55) or the model confidence is below the confidence threshold (default 0.60); both thresholds are admin-configurable. Figure 4.2 shows the gate."),
    ...figure("gate.png", "Fig. 4.2 The adaptive retrieval gate: retrieve and verify when the query looks stale or model confidence is low; answer directly when the query is fresh and the model is confident."),
    H2("4.6 Retrieval, Ranking and Live Ingestion (FR7, FR8, FR20)"),
    P("Evidence is retrieved from a FAISS index built over TF-IDF embeddings of the knowledge base, through the LangChain vector-store interface. Using TF-IDF rather than a downloaded transformer keeps the whole system offline and reproducible. Each document carries a date, and the final ranking score combines semantic relevance with a recency weight (a half-life decay on the document's age), so that among equally relevant documents a fresher one ranks higher. A minimum-relevance floor filters weak matches, and ties are broken deterministically."),
    P("A year gate protects against wrong-period evidence. If the query names an explicit year that none of the retrieved documents cover, the local evidence is treated as insufficient. This is what stops a question about the 2026 World Cup from being answered from the 2022 World Cup document."),
    P("When the local evidence is insufficient — no documents, or a year mismatch — and live ingestion is enabled, FreshGuard consults its live sources. Tavily web search returns passages for the query, and the xAI Grok model returns a concise factual passage; both are keyed by environment variables and degrade to nothing when a key is absent. Each returned passage is turned into a knowledge-base document (with a title, URL, today's date and text), embedded into the FAISS index, and persisted into SQLite. The store is then re-searched, and the freshly ingested documents are ranked, generated from and verified exactly like any seed document. Figure 4.3 shows this flow. If no live source is configured and the local evidence is only a wrong-period document, FreshGuard does not answer from it; it abstains and tells the user the information is not available and which key would enable a live fetch."),
    ...figure("live_ingest.png", "Fig. 4.3 Live ingestion on a knowledge-base miss: Tavily and Grok supply fresh passages, which are embedded, persisted, re-retrieved and verified like any other document."),
    H2("4.7 Grounded Answer Generation (FR9)"),
    P("Generation is pluggable. The default GroundedExtractiveLLM composes the answer from the evidence sentences most relevant to the query, which keeps it deterministic and offline; its parametric memory holds only timeless documents, modelling a knowledge cutoff, so time-sensitive questions cannot be answered well without retrieval — exactly the case the freshness gate is meant to catch. Hosted adapters for OpenAI, Anthropic and xAI Grok are selected through configuration and used only when their API key is present; otherwise the system falls back to the grounded generator. When Grok is used both as a live source and as the generator, the user sees Grok's content, verified against the evidence it was ingested from."),
    H2("4.8 Statement-Level Verification and Correction (FR10–FR12)"),
    P("The answer is split into sentences, and each sentence is checked against the retrieved evidence and labelled: supported when its content is covered by the evidence and no conflict is found; unsupported when its content is largely absent or it names an entity the evidence never mentions; contradicted when it conflicts with the evidence through a negation flip, a numeric mismatch on an otherwise-matching claim, or an entity swap; and review for partial support. The design uses deterministic bag-of-words coverage and light linguistic cues rather than a large model, so the whole system stays offline and reproducible. Corrective action keeps supported and review statements, drops unsupported ones with a warning, and flags contradicted ones; if nothing is supported, the system returns a low-confidence response rather than an unsupported answer."),
    H2("4.9 Confidence Scoring (FR13)"),
    P("The confidence shown to the user combines three signals: the fraction of statements that were supported, the mean relevance of the cited evidence, and the alignment between the freshness estimate and whether retrieval occurred. These are weighted (0.5, 0.3, 0.2) and scaled to a 0–100 percentage. When retrieval was skipped the confidence is model-driven, reflecting that a fresh, confidently-answered question needs no evidence. Calibration is measured in Chapter 6 through the expected calibration error."),
    H2("4.10 Requirement-to-Component Mapping"),
    P("Table 4.2 maps the twenty-two functional requirements to the components that satisfy them, and Table 4.3 lists the modules of the system with their inputs and outputs. Each module has a single responsibility and communicates through plain data, which is what made it possible to test every module on its own."),
    tcap("Table 4.2 Functional requirement to component mapping (FR1–FR22)"),
    table([900, 5100, 3200], fr),
    tcap("Table 4.3 Modules of the system"),
    table([2000, 4200, 3000], mods),
    pageBreak(),
  ];
}

function chapter5() {
  const cfg = [
    ["Parameter", "Value"],
    ["Freshness estimator", "TF-IDF (word n-grams) + LogisticRegression; 375 training examples; train accuracy 1.00"],
    ["Freshness threshold", "0.55 (retrieve if freshness ≥ this)"],
    ["Confidence threshold", "0.60 (or retrieve if model confidence < this)"],
    ["Embeddings / index", "TF-IDF embeddings in a FAISS index (inner product = cosine on unit vectors)"],
    ["Top-k documents", "4"],
    ["Minimum relevance", "0.05"],
    ["Recency half-life", "2.0 years"],
    ["Verification thresholds", "support ≥ 0.32 supported; < 0.15 unsupported; between = review"],
    ["Confidence weights", "support 0.50, evidence relevance 0.30, freshness alignment 0.20"],
    ["LLM provider (default)", "offline grounded extractive; OpenAI / Anthropic / Grok via key"],
    ["Live sources", "Tavily web search (TAVILY_API_KEY); xAI Grok (XAI_API_KEY / GROK_API_KEY)"],
  ];
  return [
    ...chapTitle("5", "IMPLEMENTATION"),
    H2("5.1 Technologies Used"),
    P([new TextRun({ text: "Language and backend. ", bold: true }), new TextRun("Python 3.11+ throughout; FastAPI with Uvicorn for the REST API; Pydantic for request and response models; SQLite (standard library) for persistence.")]),
    P([new TextRun({ text: "Retrieval and machine learning. ", bold: true }), new TextRun("scikit-learn for the TF-IDF vectoriser and the logistic-regression freshness estimator; FAISS (faiss-cpu) for the vector index, through the LangChain-core and LangChain-community vector-store interface; joblib to persist the trained estimator; NumPy and SciPy for numerical support.")]),
    P([new TextRun({ text: "Live sources. ", bold: true }), new TextRun("The Tavily web-search REST API and the xAI Grok model (OpenAI-compatible endpoint), both called with the requests library and keyed through environment variables loaded by python-dotenv.")]),
    P([new TextRun({ text: "Frontend and testing. ", bold: true }), new TextRun("A single-page HTML, CSS and JavaScript interface served by FastAPI; pytest and httpx for unit, API and end-to-end tests; Git for version control.")]),
    H2("5.2 Pipeline Configuration"),
    P("No model is trained end-to-end, so the usual training-setup table does not apply; its equivalent is the pipeline configuration — the fixed settings under which every experiment in Chapter 6 was run (Table 5.1)."),
    tcap("Table 5.1 Pipeline configuration used in the experiments"),
    table([3000, 6200], cfg),
    H2("5.3 Procedures"),
    H3("5.3.1 Build Order"),
    P("The system was built stage by stage, each stage tested before the next: configuration and schemas; the seed corpus and the freshness training data; the freshness estimator; the FAISS store and ranking; generation, verification and confidence; the pipeline, storage and error handling; the API and frontend; the evaluation module; the test suite; and finally the live-ingestion stage (Tavily and Grok) with its year gate. Building in this order meant that when something failed, the failing stage was always the newest one."),
    H3("5.3.2 Live-Ingestion Integration"),
    P("Live ingestion was added as a seam that activates only on a knowledge-base miss and only when a key is present, so the offline, deterministic behaviour of the rest of the system is unchanged. Integrating it exposed a subtle verification defect: a correct, freshly-ingested answer was being flagged as contradicting a stale seed document that still ranked into the evidence. Two fixes resolved it — a year gate that drops wrong-period documents from the evidence, and a correction to the entity-swap rule so that a sentence-initial function word is no longer mistaken for the subject of a claim. Both were found by inspecting the intermediate output before changing code, and both are covered by regression tests."),
    H2("5.4 Key Code Excerpts"),
    P("The decision policy is the heart of the freshness gate: retrieve when stale or not confident, otherwise skip."),
    code([
      "def decide(freshness, model_conf, freshness_threshold, confidence_threshold):",
      "    if freshness >= freshness_threshold:",
      "        return Decision(True, f\"freshness {freshness:.2f} >= {freshness_threshold:.2f}\")",
      "    if model_conf < confidence_threshold:",
      "        return Decision(True, f\"confidence {model_conf:.2f} < {confidence_threshold:.2f}\")",
      "    return Decision(False, \"knowledge estimated fresh; answered without retrieval\")",
    ]),
    P("On a knowledge-base miss, live sources are gathered, embedded, persisted and re-retrieved:"),
    code([
      "if settings.live_ingest_enabled and _local_evidence_insufficient(query, docs):",
      "    new_docs, sources = live_sources.gather_live_docs(query, settings)",
      "    for nd in new_docs:",
      "        doc_id = self.store.upsert(nd)          # embed into FAISS",
      "        self.storage.kb_upsert(self._stored(doc_id))   # persist to SQLite",
      "        live_ingested.append(doc_id)",
      "    if live_ingested:",
      "        docs = self.store.search(query, top_k, min_relevance)",
    ]),
    P("The year gate decides when local evidence is insufficient for a time-specific query:"),
    code([
      "def _local_evidence_insufficient(query, docs):",
      "    if not docs:",
      "        return True",
      "    q_years = set(YEAR_RE.findall(query))",
      "    if q_years:",
      "        covered = set().union(*(years(d) for d in docs))",
      "        if not (q_years & covered):   # query names a year no doc covers",
      "            return True",
      "    return False",
    ]),
    H2("5.5 Web Application"),
    P("The FastAPI backend exposes each stage as an endpoint and serves a single-page interface with three views. The Ask view sends a question and shows the answer with a freshness badge, a confidence meter, a badge indicating whether evidence was retrieved and whether it was fetched live, the per-statement verification and the cited sources. The History view lists recent questions, and the Admin view edits the decision thresholds, manages the knowledge base, chooses the LLM provider, toggles live ingestion and runs the evaluation. Every answer is persisted so history survives a reload, and each failure is returned with its real reason (for example, a 404 for an unknown history id)."),
    H2("5.6 Testing and Validation"),
    P("Testing was done at three levels: unit tests of each component (preprocessing, freshness, decision, retrieval, generation, verification, confidence, storage), API tests through the FastAPI test client, and end-to-end tests of the whole pipeline including the live-ingestion path with the network layer mocked. The suite has 45 tests. Because the default pipeline is deterministic — fixed seeds, sorted tie-breaks, a fixed reference year and inert live sources without a key — the suite was verified to pass identically across ten consecutive runs. The live-ingestion path is tested with a stubbed HTTP layer that returns canned Tavily and Grok responses, so the embed-persist-retrieve-verify flow is exercised without a real key, and a separate test confirms that the sources are inert and the system abstains honestly when no key is present."),
    pageBreak(),
  ];
}

function chapter6() {
  const res = [
    ["Metric", "Value"],
    ["Retrieval-decision (freshness) accuracy", "1.00"],
    ["Factuality (gold fact present in the answer)", "1.00"],
    ["Hallucination precision", "1.00"],
    ["Hallucination recall", "1.00"],
    ["Hallucination F1", "1.00"],
    ["Expected calibration error (ECE)", "0.10"],
    ["Automated tests passing (ten consecutive runs)", "45 / 45, identical"],
  ];
  const fut = [
    ["Condition", "Behaviour on \"Who won the 2026 FIFA World Cup?\""],
    ["Before (stale KB only)", "Wrongly answers from the 2022 World Cup document (Argentina) — a stale, year-mismatched source"],
    ["After, no live key", "Does not cite the 2022 document; abstains and states the information is not in the knowledge base, naming the key that would enable a live fetch"],
    ["After, with Tavily + Grok", "Fetches fresh passages, embeds and persists them, re-retrieves, and returns a 2026 answer with both statements verified as supported and cited"],
  ];
  return [
    ...chapTitle("6", "RESULTS AND DISCUSSION"),
    H2("6.1 Evaluation on the Controlled Benchmark"),
    P("The evaluation module reports four families of metrics: the retrieval-decision (freshness) accuracy — did the gate's retrieve-or-skip decision match the gold label; factuality — did the answer contain the gold key fact; hallucination precision, recall and F1 — how well unsupported and contradicted claims are flagged; and calibration — the expected calibration error between the confidence and correctness. Table 6.1 and Figure 6.1 report the results."),
    tcap("Table 6.1 Evaluation results on the controlled benchmark"),
    table([5600, 3600], res),
    ...figure("metrics.png", "Fig. 6.1 Evaluation metrics on the controlled benchmark: perfect freshness-decision, factuality and hallucination scores, with a calibration error of 0.10 (shown as 1 − ECE).", 5.6),
    P("On this controlled set the freshness gate chose correctly for every query — retrieving for the six time-sensitive questions and answering the six timeless ones directly — so the retrieval-decision accuracy was 1.00. Every answered question contained its gold fact (factuality 1.00), and the verification stage flagged every deliberate hallucination while passing every faithful statement (precision, recall and F1 all 1.00). The expected calibration error was 0.10, meaning the shown confidence tracked correctness closely but not perfectly. These numbers are observations on a small, controlled set, not population estimates; the set is deliberately clean so that each stage can be credited or blamed on its own."),
    H2("6.2 Freshness Gate and Determinism"),
    P("The freshness estimator separated the fresh and stale queries cleanly: timeless questions such as \"what is the speed of light\" scored well below the threshold and were answered directly from parametric memory, while time-sensitive questions such as \"what is the current federal reserve interest rate in 2025\" scored near the top of the range and triggered retrieval. Because the whole default pipeline is deterministic, the full test suite — 45 unit, API, end-to-end and evaluation tests — passed with identical results across ten consecutive runs, which is the evidence that the reported behaviour is reproducible rather than incidental."),
    H2("6.3 Live Ingestion on a Knowledge-Base Miss"),
    P("The behaviour that motivated the live-ingestion feature is a future-event question the seed base cannot answer. Table 6.2 and Figure 6.2 show the three cases. Before the feature, the system wrongly answered a 2026 question from the 2022 World Cup document, because relevance alone surfaced it. With the year gate but no live key, the system no longer cites the wrong-period document; it abstains and states that the information is not available, naming the key that would enable a live fetch. With Tavily and Grok configured, the system fetches fresh passages, embeds and persists them into the knowledge base, re-retrieves, and returns a correct 2026 answer whose statements are verified as supported and carry citations to the live sources."),
    tcap("Table 6.2 Behaviour on a future-event query with and without live sources"),
    table([2600, 6600], fut),
    ...figure("future_event.png", "Fig. 6.2 A future-event query before and after live ingestion: the stale knowledge base alone answers from the wrong year, while live ingestion fetches, verifies and cites a current answer, and abstains honestly when no key is present.", 5.8),
    H2("6.4 Error Analysis and Discussion"),
    P("Two failure directions remain. The first is an uncalibrated gate: the freshness and confidence thresholds are sensible defaults rather than values tuned on a held-out set, and the expected calibration error of 0.10 shows there is room to tighten the match between confidence and correctness. The second concerns the verification rules: they are deliberately lightweight and deterministic, so they can be conservative on paraphrased or compositional statements, and they check what an answer says, not what it leaves out — a complete-looking answer can still omit part of what was asked. Integrating live ingestion also surfaced, and fixed, a verification bug in which a correct live answer was flagged as contradicting a stale seed document; the year gate and the corrected entity-swap rule resolved it, and both are covered by regression tests."),
    H2("6.5 Ethical and Responsible AI Considerations"),
    P([new TextRun({ text: "Not professional advice. ", bold: true }), new TextRun("FreshGuard is a research prototype. Its value is that it signals when its knowledge may be stale, shows its evidence, and abstains when the evidence is weak, which leaves the final judgement with a person who can read the source.")]),
    P([new TextRun({ text: "Data flow and keys. ", bold: true }), new TextRun("The index and history are stored locally. When live ingestion is enabled, the query is sent to the Tavily and xAI services, which users should be aware of; without a key the system runs entirely offline. API keys are kept in a git-ignored environment file and never committed.")]),
    P([new TextRun({ text: "Honesty about limits. ", bold: true }), new TextRun("The evaluation is reported with its limitations — a small controlled set, an uncalibrated gate, and conservative verification rules — so the system's limits are part of its record rather than a surprise for its users.")]),
    pageBreak(),
  ];
}

function chapter7() {
  return [
    ...chapTitle("7", "CONCLUSION AND FUTURE SCOPE"),
    H2("7.1 Conclusion"),
    P("This project set out to make an LLM's factual answers both timely and checkable. FreshGuard estimates, per query, how likely the model's knowledge is stale; retrieves only when that estimate or low model confidence calls for it; fetches and ingests fresh evidence from Tavily and Grok when the knowledge base has none; verifies every statement of the answer against the evidence; and reports a freshness badge, a calibrated confidence and cited sources — abstaining honestly when no reliable evidence can be found."),
    P("The experiments support the design. On the controlled benchmark the retrieval-decision accuracy, factuality and hallucination F1 were each 1.00 and the expected calibration error was 0.10, and the 45-test suite passed identically across ten consecutive runs, so the behaviour is reproducible. The live-ingestion feature turned a question the seed base answered wrongly (a 2026 event answered from a 2022 document) into one answered correctly from freshly fetched, verified and cited evidence — or, without a key, into an honest abstention rather than a confident error."),
    P("In summary, a small, mostly-offline pipeline can make an LLM freshness-aware and auditable: the combination of a learned freshness gate, adaptive retrieval with live ingestion, and statement-level verification kept stale and unsupported answers away from the user in our tests. The most important next steps are to calibrate the gate and to evaluate on a larger, more realistic question set."),
    H2("7.2 Limitations"),
    bullet("Uncalibrated gate: the 0.55 freshness and 0.60 confidence thresholds are defaults; no held-out calibration set has been used to tune them."),
    bullet("Small evaluation: twelve labelled queries, eight hallucination pairs and a demonstrated live-ingestion flow — the numbers are observations, not rates."),
    bullet("Live ingestion depends on an external key, and the quality of a live answer depends on the quality of what Tavily and Grok return."),
    bullet("Lightweight verification: deterministic bag-of-words and linguistic rules can be conservative on paraphrase, and they check faithfulness, not completeness."),
    bullet("Offline default generator: the grounded extractive LLM composes from evidence rather than reasoning over it, which bounds answer fluency unless a hosted model is configured."),
    H2("7.3 Future Enhancements"),
    P([new TextRun({ text: "Freshness and calibration. ", bold: true }), new TextRun("Calibrate the thresholds on a held-out set that includes unanswerable questions; enrich the freshness estimator with character n-grams and an explicit entity-recency feature; and recalibrate the shown confidence (for example, with temperature scaling) to drive the ECE below its current 0.10.")]),
    P([new TextRun({ text: "Retrieval and sources. ", bold: true }), new TextRun("Add a transformer encoder and a keyword (BM25) index alongside TF-IDF, with rank fusion; cache and deduplicate live-ingested documents; and add more live sources behind the same seam.")]),
    P([new TextRun({ text: "Verification. ", bold: true }), new TextRun("Add a natural-language-inference model for entailment, atomic-fact decomposition for long answers, and a completeness check that flags when part of a multi-part question was not answered.")]),
    P([new TextRun({ text: "Product. ", bold: true }), new TextRun("A local LLM option so confidential queries need not leave the machine, authentication and per-user knowledge bases for deployment, and a larger, harder evaluation set drawn from real time-sensitive questions.")]),
    pageBreak(),
  ];
}

function references() {
  const refs = [
    "P. Lewis et al., \"Retrieval-augmented generation for knowledge-intensive NLP tasks,\" in Advances in Neural Information Processing Systems (NeurIPS), 2020.",
    "A. Asai, Z. Wu, Y. Wang, A. Sil and H. Hajishirzi, \"Self-RAG: Learning to retrieve, generate and critique through self-reflection,\" in Proc. ICLR, 2024.",
    "A. Mallen et al., \"When not to trust language models: Investigating the effectiveness of parametric and non-parametric memories,\" in Proc. ACL, 2023.",
    "S. Jin, Z. Wang et al., \"FreshLLMs: Refreshing large language models with search-engine augmentation,\" in Findings of ACL, 2024.",
    "N. Reimers and I. Gurevych, \"Sentence-BERT: Sentence embeddings using Siamese BERT-networks,\" in Proc. EMNLP-IJCNLP, 2019.",
    "J. Johnson, M. Douze and H. Jégou, \"Billion-scale similarity search with GPUs (FAISS),\" IEEE Transactions on Big Data, 2019.",
    "G. V. Cormack, C. L. A. Clarke and S. Büttcher, \"Reciprocal rank fusion outperforms Condorcet and individual rank learning methods,\" in Proc. ACM SIGIR, 2009.",
    "S. Min et al., \"FActScore: Fine-grained atomic evaluation of factual precision in long-form text generation,\" in Proc. EMNLP, 2023.",
    "P. Manakul, A. Liusie and M. J. F. Gales, \"SelfCheckGPT: Zero-resource black-box hallucination detection for generative large language models,\" in Proc. EMNLP, 2023.",
    "C. Guo, G. Pleiss, Y. Sun and K. Q. Weinberger, \"On calibration of modern neural networks,\" in Proc. ICML, 2017.",
    "A. Williams, N. Nangia and S. R. Bowman, \"A broad-coverage challenge corpus for sentence understanding through inference (MultiNLI),\" in Proc. NAACL-HLT, 2018.",
    "F. Pedregosa et al., \"Scikit-learn: Machine learning in Python,\" Journal of Machine Learning Research, vol. 12, 2011.",
    "S. Ramírez, \"FastAPI,\" 2018. [Online]. Available: https://fastapi.tiangolo.com (accessed 2026).",
    "H. Chase, \"LangChain,\" 2022. [Online]. Available: https://github.com/langchain-ai/langchain (accessed 2026).",
    "Tavily, \"Tavily Search API documentation,\" 2024. [Online]. Available: https://docs.tavily.com (accessed 2026).",
    "xAI, \"Grok API (x.ai) documentation,\" 2024. [Online]. Available: https://docs.x.ai (accessed 2026).",
  ];
  return [
    center([new TextRun({ text: "REFERENCES", bold: true, size: 28 })], { after: 160 }),
    ...refs.map((r, i) => new Paragraph({ spacing: { after: 110, line: 290 }, alignment: AlignmentType.JUSTIFIED,
      children: [new TextRun({ text: `[${i + 1}] `, bold: true }), new TextRun(r)] })),
    pageBreak(),
  ];
}

function appendix() {
  const ep = [
    ["Method", "Path", "Purpose"],
    ["POST", "/api/query", "Run the full pipeline: answer, confidence, freshness, sources, verdicts"],
    ["POST", "/api/feedback", "Rate or report a stored answer"],
    ["GET", "/api/history", "List recent questions"],
    ["GET", "/api/history/{id}", "Full stored response"],
    ["GET / POST", "/api/config", "Read / update admin configuration (thresholds, provider, live toggles)"],
    ["GET / POST / DELETE", "/api/kb, /api/kb/{id}", "List, add and remove knowledge-base documents"],
    ["GET", "/api/eval", "Run the evaluation module"],
    ["GET", "/api/health", "Liveness check and index size"],
  ];
  return [
    center([new TextRun({ text: "APPENDIX A — SETUP, RUN AND API GUIDE", bold: true, size: 26 })], { after: 160 }),
    H2("A.1 Environment Setup"),
    P("From the project root, using a Python 3.11+ environment:"),
    code([
      "pip install -r requirements.txt",
      "python run.py --train      # train the freshness estimator (one-off)",
      "python run.py              # serve the API + web UI at http://127.0.0.1:8000",
    ]),
    P("Then open http://127.0.0.1:8000, ask a question (the example chips are a quick start), and use the Admin tab to change thresholds, manage the knowledge base and run the evaluation."),
    H2("A.2 Environment File (keys)"),
    P("A git-ignored .env file holds the optional keys; with all of them blank the system runs fully offline."),
    code([
      "# Live knowledge ingestion (used when a query is not in the knowledge base)",
      "TAVILY_API_KEY=            # Tavily web search",
      "XAI_API_KEY=              # xAI Grok (or GROK_API_KEY)",
      "# Optional hosted generation",
      "OPENAI_API_KEY=",
      "ANTHROPIC_API_KEY=",
    ]),
    H2("A.3 Testing"),
    code([
      "python -m pytest -q        # 45 unit / end-to-end / API / evaluation tests",
      "python eval/run_eval.py    # writes eval/metrics.json",
    ]),
    H2("A.4 REST API Endpoints"),
    tcap("Table A.1 REST endpoints of the backend"),
    table([2000, 2600, 4600], ep),
    H2("A.5 Repository Layout"),
    code([
      "freshguard/   core package (preprocess, freshness, decision, retrieval,",
      "              live_sources, llm, verify, confidence, pipeline, storage,",
      "              evaluation, api, data)",
      "frontend/     single-page web UI (ask, history, admin)",
      "tests/        pytest suite (45 tests)",
      "eval/         evaluation runner + metrics.json",
      "artifacts/    trained freshness estimator + SQLite database",
      "run.py        launcher: --train, or serve the API + UI",
    ]),
    P([new TextRun({ text: "Research prototype — not for professional decisions. Verify important answers independently.", italics: true })]),
  ];
}

// ---------- assemble ----------
const doc = new Document({
  creator: "FreshGuard Project Team",
  title: TITLE,
  styles: {
    default: { document: { run: { font: "Times New Roman", size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: "Times New Roman", size: 26, bold: true, color: NAVY } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: "Times New Roman", size: 24, bold: true, color: "1A202C" } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: "Times New Roman", size: 23, bold: true, italics: true, color: "2D3748" } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 480, hanging: 240 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "◦", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 960, hanging: 240 } } } },
      ] },
      { reference: "nums", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 480, hanging: 240 } } } },
      ] },
    ],
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1440, left: 1620, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], size: 18 })] })] }) },
    children: [
      ...titlePage(),
      ...certificate(),
      ...declaration(),
      ...acknowledgement(),
      ...abstract(),
      ...toc(),
      ...listOfFigures(),
      ...listOfTables(),
      ...abbreviations(),
      ...chapter1(),
      ...chapter2(),
      ...chapter3(),
      ...chapter4(),
      ...chapter5(),
      ...chapter6(),
      ...chapter7(),
      ...references(),
      ...appendix(),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = path.join(__dirname, "FreshGuard_Project_Report.docx");
  fs.writeFileSync(out, buf);
  console.log("wrote", out, buf.length, "bytes");
});
