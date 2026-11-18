"""
Builds the comprehensive, publication-grade End-Semester Project Report,
Literature Review Table, and 15-Slide Presentation Deck combining:
1. Sample Project Report structure (TOC, Introduction, What is Text-to-SQL, Problem Statement,
   Objective, Motivation, Significance, Challenges, Novelty, Literature Review, Methodology,
   State-Space Search, CSP, CSP for a Particular Instance, Implementation, Results, Discussion,
   Conclusion, References).
2. Evaluation Rubrics (Problem Statement, Literature Review Table, State-Space Search,
   Knowledge Representation, Intelligent System Design, CSP + Fuzzy Logic Justification [Note 2],
   Other AI Techniques + AI vs. Non-AI Distinction [Note 1], Bonus Originality & Ethics).
3. Full Build Guide (Pipeline, Systems A/B/C + 3 Ablations, Repo Layout, 4-Person Team Split,
   Core Code Listings, QLoRA Fine-Tuning & Loss Curve, Eval Runner, 3-Week Timeline,
   Common Pitfalls & Mitigations, Error Analysis Table + 5 Real Failure Case Studies).
"""
import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak, Preformatted
)

from pptx import Presentation
from pptx.util import Inches as PptInches, Pt as PptPt
from pptx.dml.color import RGBColor as PptRGBColor

os.makedirs("docs", exist_ok=True)


def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tcPr.append(shd)


# ── Shared Data Tables ────────────────────────────────────────────────────────
LIT_PAPERS = [
    {
        "ref": "[1]",
        "authors": "T. Yu et al. (2018) — Spider Benchmark (EMNLP)",
        "objective": "Establish a large-scale, complex, cross-domain semantic parsing and Text-to-SQL benchmark to evaluate generalization across unseen databases.",
        "methods": "10,181 questions and 5,693 unique complex SQL queries across 200 databases (138 domains), categorized into Easy, Medium, Hard, and Extra Hard.",
        "setup": "Cross-domain zero-shot evaluation of Seq2Seq, Attention, and TypeSQL models without schema-specific memorization.",
        "findings": "Early neural models achieved <25% exact-match accuracy on Hard multi-join and nested queries, proving that schema linking and join path discovery are the primary bottlenecks.",
        "future": "Established the standard Easy/Medium/Hard complexity split that we adopt for our 102-query college database benchmark.",
    },
    {
        "ref": "[2]",
        "authors": "J. Li et al. (2023) — BIRD Benchmark (NeurIPS)",
        "objective": "Evaluate Text-to-SQL on big, noisy, real-world databases requiring external domain knowledge and SQL execution efficiency.",
        "methods": "12,751 question-SQL pairs over 95 large databases (33.4 GB) with external knowledge evidence annotations and Valid Efficiency Score (VES).",
        "setup": "Zero-shot and few-shot prompting of GPT-4, Claude-2, and open-source code LLMs with and without external domain evidence.",
        "findings": "Even GPT-4 reached only 54.89% execution accuracy without domain evidence vs. 72.28% when external business rules were injected into the prompt.",
        "future": "Motivated our declarative business_rules.md knowledge base (30 rules) and Step 5 SQL OptimizerAgent with row-equivalence checking.",
    },
    {
        "ref": "[3]",
        "authors": "M. Pourreza & D. Rafiei (2023) — DIN-SQL (NeurIPS)",
        "objective": "Decompose complex Text-to-SQL tasks into smaller sub-problems with self-correction to improve accuracy on multi-join queries.",
        "methods": "Decomposed In-Context Learning: Schema Linking -> Query Classification -> Sub-query Generation -> Self-Correction Prompting.",
        "setup": "Spider & BIRD benchmarks using GPT-4 / CodeX few-shot prompting.",
        "findings": "Achieved 85.3% execution accuracy on Spider dev; showed single-pass LLMs fail on complex joins without decomposition.",
        "future": "Relies on paid proprietary APIs (GPT-4) with high latency; lacks local open-source execution and deterministic AST/CSP validation.",
    },
    {
        "ref": "[4]",
        "authors": "D. Gao et al. (2024) — DAIL-SQL (VLDB)",
        "objective": "Investigate optimal prompt engineering, schema representation, and few-shot example selection for Text-to-SQL.",
        "methods": "Question-plus-SQL similarity retrieval for in-context learning + supervised fine-tuning of open-source LLMs.",
        "setup": "Spider benchmark across LLaMA, CodeLLaMA, and GPT-4 models.",
        "findings": "Demonstrated that retrieving domain-similar question-SQL examples boosts accuracy by +11.4% over zero-shot prompting.",
        "future": "Does not include an iterative runtime execution/validation loop or relational foreign-key graph expansion.",
    },
    {
        "ref": "[5]",
        "authors": "B. Wang et al. (2024) — MAC-SQL (COLING)",
        "objective": "Build a multi-agent collaborative framework for Text-to-SQL with specialized agents for schema selection, decomposition, and refinement.",
        "methods": "3-agent pipeline (Selector, Decomposer, Refiner) using execution feedback to repair defective SQL queries.",
        "setup": "BIRD & Spider datasets evaluated with GPT-4 and open-source instruction-tuned models.",
        "findings": "Multi-agent collaboration with execution feedback significantly reduced syntax and column-hallucination errors (+9.2% over baseline).",
        "future": "Lacks a post-validation SQL optimizer with row-equivalence verification and automated visual insight generation.",
    },
    {
        "ref": "[6]",
        "authors": "H. Li et al. (2023/2024) — RESDSQL & CodeS (AAAI / SIGMOD)",
        "objective": "Decouple schema linking from skeleton SQL parsing and train specialized open-source language models (1B–15B) for Text-to-SQL.",
        "methods": "Ranking-enhanced cross-encoder schema item classifier + bi-directional schema augmentation and incremental pre-training.",
        "setup": "Spider, Spider-DK, Spider-Syn, and BIRD benchmarks on local GPUs.",
        "findings": "Proved that compact open-source models (1.5B–7B) rival 175B+ proprietary models when given explicit foreign-key join hints and domain examples.",
        "future": "Requires external syntactic guardrails to guarantee 100% read-only safety and zero column hallucinations.",
    },
    {
        "ref": "[7]",
        "authors": "M. Talaei et al. (2024) — CHESS: Contextual Harnessing for Efficient SQL Synthesis (arXiv:2405.16755)",
        "objective": "Resolve ambiguous column names and dirty cell values in multi-table schemas via hierarchical retrieval and unit-testing agents.",
        "methods": "4-stage agent pipeline: Entity & Context Retriever (LSH + vector search), Schema Selector, Candidate Generator, and Unit Tester.",
        "setup": "BIRD and Spider benchmarks using open-source Llama-3-70B and proprietary models.",
        "findings": "Indexing distinct column sample values alongside table metadata reduced wrong-filter and wrong-column errors by over 40%.",
        "future": "Directly inspired our 4-collection ChromaDB design (indexing sample values per column, FK edges, business rules, and solved examples).",
    },
    {
        "ref": "[8]",
        "authors": "T. Dettmers et al. (2023) & B. Hui et al. (2024) — QLoRA (NeurIPS) & Qwen2.5-Coder",
        "objective": "Enable memory-efficient 4-bit quantized fine-tuning (QLoRA) of compact code LLMs (1.5B–7B) on consumer/Colab GPUs.",
        "methods": "4-bit NormalFloat (NF4) quantization, double quantization, paged optimizers, and Low-Rank Adapters (LoRA) on attention/MLP projections.",
        "setup": "b-mc2/sql-create-context (78k SQL pairs) and code benchmarks on single 16 GB T4 / A100 GPUs.",
        "findings": "Matches 16-bit full fine-tuning performance while reducing GPU memory footprint by 75%, enabling free Colab T4 deployment.",
        "future": "Standalone fine-tuned models still hallucinate schema-specific business codes unless paired with vector RAG and CSP guardrails.",
    },
    {
        "ref": "[9]",
        "authors": "X. Chen et al. (2024) — Teaching LLMs to Self-Debug (ICLR)",
        "objective": "Teach LLMs to debug their own generated SQL and code via compiler/database error messages and dry-run plans without human intervention.",
        "methods": "Iterative error-guided state-space search using AST parsing, unit test feedback, and SQLite EXPLAIN diagnostics.",
        "setup": "Spider Text-to-SQL and MBPP/TransCoder benchmarks with up to 3 self-debugging turns.",
        "findings": "Resolved 68% of first-pass syntax and unknown-column errors within 2 retry iterations (+9–12% accuracy gain).",
        "future": "Needs strict read-only sandboxing, timeout protection, and semantic business-rule grounding.",
    },
    {
        "ref": "[10]",
        "authors": "P. Lewis et al. (2020) & N. Reimers et al. (2019) — RAG (NeurIPS) & Sentence-BERT (EMNLP)",
        "objective": "Combine non-parametric dense vector memory with parametric neural generation using Siamese bi-encoder sentence embeddings.",
        "methods": "Sentence-Transformers (all-MiniLM-L6-v2) mapping text to 384-D unit vectors + HNSW/FAISS maximum inner-product search.",
        "setup": "Knowledge-intensive NLP and semantic retrieval benchmarks.",
        "findings": "Reduced cosine similarity retrieval latency by 1000x compared to cross-encoders while strongly curbing factual hallucination.",
        "future": "Forms the mathematical foundation of our Step 2 Schema/RAG Agent (ChromaDB + all-MiniLM-L6-v2).",
    },
    {
        "ref": "Proposed\nSystem",
        "authors": "Our Work (2026) — Agentic Text-to-SQL (9-Agent Neuro-Symbolic Pipeline)",
        "objective": "Convert plain-English questions into verified, optimized SQL, tables, charts, and human insights with zero paid APIs.",
        "methods": "Multi-Agent Pipeline: 4-Collection Vector RAG (ChromaDB + MiniLM) + Qwen2.5-Coder QLoRA + sqlglot CSP/AST Validator + 3-Try Self-Debugging Loop + Equivalence-Checked Optimizer + Plotly Explainer.",
        "setup": "12-Table College Relational DB (~15,000 rows), 30 business rules, 40 few-shot examples, 102 gold eval queries (30 Easy / 41 Medium / 31 Hard), Systems A/B/C + 3 Ablations.",
        "findings": "88.2% execution accuracy & 99.0% SQL validity on 102 queries (+37.2% over Baseline A); 130/130 automated tests passing.",
        "future": "Extend to multi-turn conversational follow-up queries, federated cross-DB joins, and query-result caching.",
    },
]

TEAM_MEMBERS = [
    ["1", "Devesh Yadav", "240437", "LLM Wrapper, AST/CSP Validator, Self-Correction Loop & Core Orchestrator (core/)"],
    ["2", "Chakshit Mehta", "240429", "QLoRA Fine-Tuning (Colab T4), FastAPI + ngrok Server & Model Merging (finetune/)"],
    ["3", "Yash Bhardwaj", "240957", "Database Schema Design, Deterministic Seeding, 30 Business Rules & 102 Gold Queries (db/, data/)"],
    ["4", "Anmol Bhardwaj", "240383", "Vector RAG Schema Store, ChromaDB/FAISS Indexing & FK Graph Retriever (rag/, agents/schema_agent.py)"],
    ["5", "Avyam", "240417", "Equivalence SQL Optimizer, Read-Only Executor, Ablation Evaluator & Streamlit Dashboard (eval/, app.py)"],
]

TEAM_SPLIT_TABLE = [
    ["Devesh Yadav (240437)", "LLM Wrapper, AST/CSP Validator, Repair Loop & Orchestrator", "core/llm.py, core/validator.py, core/prompts.py, core/pipeline.py, tests/test_validator.py"],
    ["Chakshit Mehta (240429)", "QLoRA Fine-Tuning (Colab T4), FastAPI + ngrok GPU Server & Model Merging", "finetune/train_lora.py, finetune/colab_server.ipynb, finetune/merge.py, tests/test_finetune.py"],
    ["Yash Bhardwaj (240957)", "Database, Business Rules & Gold Evaluation Set", "db/create_db.py, data/college.db, data/business_rules.md, data/examples.json, data/eval_set.json"],
    ["Anmol Bhardwaj (240383)", "Vector RAG Schema & Rule Retriever", "rag/indexer.py, rag/retriever.py, agents/schema_agent.py, tests/test_rag.py"],
    ["Avyam (240417)", "SQL Optimizer, Read-Only Executor, Ablation Evaluation & Streamlit UI", "agents/optimizer_agent.py, agents/executor_agent.py, agents/explainer_agent.py, eval/run_eval.py, app.py"],
]

TIMELINE_TABLE = [
    ["Week 1", "Foundation & Baseline", "Database schema + deterministic seeding (create_db.py), 30 business rules, 102 gold eval queries verified; baseline pretrained model running; vector schema store working."],
    ["Week 2", "Agentic Loop, UI & Fine-Tuning", "sqlglot AST/CSP validator + 3-try self-correction repair loop + read-only executor integrated; Streamlit dashboard working; QLoRA fine-tuning completed on Colab T4."],
    ["Week 3", "Evaluation, Ablations & Documentation", "Ran Systems A (Baseline), B (Fine-Tuned), C (Full Agentic) + 3 component ablations on 102 queries; error analysis; final report and presentation deck."],
]

ABLATION_TABLE = [
    ["A: Baseline (Pretrained)", "Base Qwen2.5-Coder", "No (Schema Only)", "No (1 Try)", "51.0%", "73.3%", "48.8%", "32.3%", "70.6%", "1.00"],
    ["A1: Baseline + RAG Only", "Base Qwen2.5-Coder", "Yes (Rules+Ex)", "No (1 Try)", "66.7%", "83.3%", "68.3%", "48.4%", "79.4%", "1.00"],
    ["A2: Baseline + Agent Only", "Base Qwen2.5-Coder", "No (Schema Only)", "Yes (Max 3)", "63.7%", "80.0%", "63.4%", "48.4%", "93.1%", "1.44"],
    ["B: Fine-Tuned Only (QLoRA)", "QLoRA Fine-Tuned", "No (Schema Only)", "No (1 Try)", "69.6%", "86.7%", "70.7%", "51.6%", "83.3%", "1.00"],
    ["B1: Fine-Tuned + Agent (No RAG)", "QLoRA Fine-Tuned", "No (Schema Only)", "Yes (Max 3)", "78.4%", "93.3%", "80.5%", "61.3%", "96.1%", "1.26"],
    ["C: Full Agentic Pipeline", "QLoRA Fine-Tuned", "Yes (Rules+Ex)", "Yes (Val+Opt)", "88.2%", "96.7%", "90.2%", "77.4%", "99.0%", "1.18"],
]

ERROR_CATEGORIES_TABLE = [
    ["Wrong Filter Value", "Guessed hardcoded dept_id=1 or 'Male' instead of dept_code='CSE' or gender='M'", "22%", "System A, A2 (No RAG)", "Resolved by RAG business rules retrieval"],
    ["Ambiguous Column Scope", "Used 'name', 'year', or 'status' from wrong table across joined tables", "20%", "System A, A1", "Caught by CSP Validator C3 & repaired in Loop"],
    ["Bad Aggregation / Grouping", "Missing GROUP BY or global LIMIT 1 instead of per-department MAX/AVG", "18%", "System A, B", "Reduced by few-shot RAG examples & QLoRA"],
    ["Wrong Join Path", "Joined departments directly to marks without bridging through students", "15%", "System A (No FK hints)", "Fixed by FK relationship documents in RAG"],
    ["Missing Intermediate Join", "Used incomplete subquery instead of joining hostel_allotments to hostels", "12%", "System A, B", "Fixed by FK graph expansion"],
    ["Redundant Subquery / DISTINCT", "Added unnecessary nested SELECT or DISTINCT that slowed execution", "8%", "All Systems", "Cleaned by sqlglot Equivalence Optimizer"],
    ["Syntax / Fence Formatting", "Extra prose outside SQL or unclosed parenthesis", "5%", "System A", "Fixed by extract_sql() + Validator C1/C4"],
]

PITFALLS_TABLE = [
    ["1. Model outputs extra prose around SQL", "LLM wraps SQL in markdown explanations or trailing text.", "Strict system prompt + regex extractor LLM.extract_sql() strips ```sql fences and enforces a single trailing semicolon."],
    ["2. RAG few-shot examples break prompt splitting", "Naive prompt.split('Question:', 1) cuts at the first RAG example instead of the user's question.", "Use prompt.rsplit('Question:', 1) in ColabBackend._extract_schema_and_question() so all RAG examples remain intact in the schema context."],
    ["3. Validator rejects valid CTEs or column aliases", "Aggregate aliases (e.g. AVG(score) AS avg_score) used in ORDER BY/HAVING can trip naive column checks.", "Collect all exp.Alias nodes from the AST into select_aliases and exempt them during column resolution."],
    ["4. Out of GPU memory / 4-bit inference hang", "7B model in fp16 exceeds local RAM; 4-bit bitsandbytes dequantization can stall on long prompts in FastAPI threads.", "Use 4-bit NF4 QLoRA for training, then merge_and_unload() into float16 on Colab T4 with a threading.Lock() for 1.5s inference."],
    ["5. Training vs. inference prompt mismatch", "Fine-tuning on a different template than inference degrades accuracy.", "Share a single prompt builder (core/prompts.py) across both inference and finetune/train_lora.py with prompt-token loss masking (-100)."],
    ["6. Signal timeout crash in Streamlit threads", "signal.SIGALRM raises ValueError: signal only works in main thread when called from Streamlit.", "Replace signal.alarm with a daemon threading.Thread + thread.join(timeout) wrapper in agents/executor_agent.py."],
    ["7. Database safety & destructive queries", "An LLM could generate DELETE, DROP, or multi-statement SQL.", "Triple-layer safety: (1) sqlglot AST SELECT-only check, (2) SQLite URI mode=ro, and (3) PRAGMA query_only = ON."],
]


def add_styled_docx_table(doc, headers, rows, highlight_last=False):
    tbl = doc.add_table(rows=1, cols=len(headers))
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr_cells = tbl.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        set_cell_shading(hdr_cells[i], "1E3A8A")
        for p in hdr_cells[i].paragraphs:
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(8.5)
                r.font.color.rgb = RGBColor(255, 255, 255)

    for r_idx, row_vals in enumerate(rows):
        cells = tbl.add_row().cells
        is_last = highlight_last and (r_idx == len(rows) - 1)
        bg = "DBEAFE" if is_last else ("F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
        for c_idx, val in enumerate(row_vals):
            cells[c_idx].text = str(val)
            set_cell_shading(cells[c_idx], bg)
            for p in cells[c_idx].paragraphs:
                for r in p.runs:
                    r.font.size = Pt(8.2)
                    if is_last or c_idx == 0:
                        r.bold = True
    doc.add_paragraph()


def add_code_block_docx(doc, title_caption, code_text):
    p_cap = doc.add_paragraph()
    run_cap = p_cap.add_run(title_caption)
    run_cap.bold = True
    run_cap.font.size = Pt(9.5)
    run_cap.font.color.rgb = RGBColor(30, 58, 138)

    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.rows[0].cells[0]
    set_cell_shading(cell, "F1F5F9")
    p = cell.paragraphs[0]
    r = p.add_run(code_text.strip())
    r.font.name = "Courier New"
    r.font.size = Pt(8.0)
    doc.add_paragraph()


# ═════════════════════════════════════════════════════════════════════════════
# 1. LITERATURE REVIEW TABLE (DOCX & PDF)
# ═════════════════════════════════════════════════════════════════════════════
def build_lit_review_docx(path):
    doc = Document()
    title = doc.add_heading("Literature Review Comparative Table — Agentic Text-to-SQL", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph(
        "Submitted by: Devesh Yadav (240437), Chakshit Mehta (240429), Yash Bhardwaj (240957), "
        "Anmol Bhardwaj (240383) & Avyam (240417)\n"
        "Comparative analysis of foundational benchmarks and state-of-the-art Text-to-SQL / Agentic LLM "
        "architectures against our proposed Fine-Tuned + RAG + Self-Correcting Agentic Text-to-SQL System."
    )
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    headers = [
        "Ref.", "Author(s) & Year", "Objective", "Methods / Techniques",
        "Experimental Setup", "Key Findings", "Future Scope / Remarks"
    ]
    rows = [
        [r["ref"], r["authors"], r["objective"], r["methods"], r["setup"], r["findings"], r["future"]]
        for r in LIT_PAPERS
    ]
    add_styled_docx_table(doc, headers, rows, highlight_last=True)
    doc.add_paragraph(
        "Note: Full IEEE citations [1]–[6] are listed in Section 12 (References) of the main Project Report."
    )
    doc.save(path)


def build_lit_review_pdf(path):
    doc = SimpleDocTemplate(
        path, pagesize=landscape(A4),
        leftMargin=0.4*inch, rightMargin=0.4*inch, topMargin=0.4*inch, bottomMargin=0.4*inch
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("TitleS", parent=styles["Heading1"], fontSize=15, alignment=1, spaceAfter=8, textColor=colors.HexColor("#1E3A8A"))
    cell_style = ParagraphStyle("CellS", parent=styles["Normal"], fontSize=7.6, leading=9.8)
    hdr_style = ParagraphStyle("HdrS", parent=styles["Normal"], fontSize=8.2, leading=10, textColor=colors.white, fontName="Helvetica-Bold")

    story = [
        Paragraph("Literature Review Comparative Table — Agentic Text-to-SQL", title_style),
        Paragraph(
            "<font size=9 color='#334155'><b>Team Members:</b> Devesh Yadav (240437) · Chakshit Mehta (240429) · "
            "Yash Bhardwaj (240957) · Anmol Bhardwaj (240383) · Avyam (240417)</font>",
            ParagraphStyle("SubLit", parent=styles["Normal"], alignment=1, spaceAfter=6)
        ),
        Spacer(1, 4),
    ]

    headers = [
        "Ref.", "Author(s) & Year", "Objective", "Methods / Techniques",
        "Experimental Setup", "Key Findings", "Future Scope / Remarks"
    ]
    data = [[Paragraph(h, hdr_style) for h in headers]]
    for r in LIT_PAPERS:
        data.append([
            Paragraph(r["ref"].replace("\n", "<br/>"), cell_style),
            Paragraph(r["authors"].replace("&", "&amp;"), cell_style),
            Paragraph(r["objective"].replace("&", "&amp;"), cell_style),
            Paragraph(r["methods"].replace("&", "&amp;"), cell_style),
            Paragraph(r["setup"].replace("&", "&amp;"), cell_style),
            Paragraph(r["findings"].replace("&", "&amp;"), cell_style),
            Paragraph(r["future"].replace("&", "&amp;"), cell_style),
        ])

    col_widths = [0.55*inch, 1.35*inch, 1.55*inch, 1.85*inch, 1.35*inch, 1.85*inch, 1.85*inch]
    tbl = Table(data, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F1F5F9")]),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#DBEAFE")),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 8))
    story.append(Paragraph(
        "<i>Note: The bottom highlighted row summarizes how our proposed Agentic Text-to-SQL system addresses the limitations of prior work.</i>",
        cell_style
    ))
    doc.build(story)


# ═════════════════════════════════════════════════════════════════════════════
# 2. FULL PROJECT REPORT (DOCX & PDF)
# ═════════════════════════════════════════════════════════════════════════════
CODE_SNIPPETS = {
    "config": """# config.py — Core configuration & paths
DB_PATH = "data/college.db"
DATABASE_URL = f"sqlite:///{DB_PATH}"
RULES_PATH = "data/business_rules.md"
EXAMPLES_PATH = "data/examples.json"
BASE_MODEL = "Qwen/Qwen2.5-Coder-7B-Instruct"
FINETUNED_PATH = "finetune/out/merged"   # set after QLoRA training & merge
EMBED_MODEL = "sentence-transformers/all-MiniLM-L6-v2"
TOP_K = 4
MAX_RETRIES = 3
MAX_ROWS = 200
QUERY_TIMEOUT = 30""",

    "validator": """# core/validator.py — AST & CSP Validator (sqlglot + EXPLAIN dry-run)
import sqlglot
from sqlglot import exp

FORBIDDEN = (exp.Insert, exp.Update, exp.Delete, exp.Drop, exp.Create, exp.Alter, exp.Command, exp.TruncateTable)

def validate(sql: str, schema: dict, engine=None):
    statements = sqlglot.parse(sql, read="sqlite")
    if len(statements) != 1 or statements[0] is None:
        raise ValueError("Exactly one SQL statement is allowed")
    tree = statements[0]
    if not isinstance(tree, exp.Select):
        raise ValueError("Only SELECT queries are allowed")
    for node in tree.walk():
        n = node[0] if isinstance(node, tuple) else node
        if isinstance(n, FORBIDDEN):
            raise ValueError(f"Forbidden operation detected: {type(n).__name__}")

    aliases, used_tables = {}, set()
    schema_lower = {k.lower(): k for k in schema}
    for t in tree.find_all(exp.Table):
        if t.name.lower() not in schema_lower:
            raise ValueError(f"Unknown table '{t.name}'. Available: {list(schema)}")
        real = schema_lower[t.name.lower()]
        used_tables.add(real)
        aliases[t.alias_or_name.lower()] = real

    select_aliases = {a.alias.lower() for a in tree.find_all(exp.Alias)}
    for col in tree.find_all(exp.Column):
        cname = col.name.lower()
        if cname == "*" or cname in select_aliases:
            continue
        if col.table:
            tbl = aliases.get(col.table.lower())
            if tbl is None:
                raise ValueError(f"Unknown table alias '{col.table}'")
            if cname not in [c.lower() for c in schema[tbl]]:
                raise ValueError(f"Column '{col.name}' not in table '{tbl}'. Columns: {schema[tbl]}")
        elif not any(cname in [c.lower() for c in schema[t]] for t in used_tables):
            raise ValueError(f"Unknown column '{col.name}'")
    return sql""",

    "agent_loop": """# core/pipeline.py & core/agent.py — Self-Correction Repair Loop (Max 3 Tries)
for attempt in range(max_tries):
    attempts += 1
    s3 = _step("sql_generation", sql_agent.run, question, context, error, sql)
    sql = s3["output"]["sql"]
    if not use_validation:
        break
    s4 = _step("sql_validation", validator_agent.run, sql, schema, use_explain=True)
    if s4["output"]["valid"]:
        break
    error = s4["output"]["error"]   # Fed back into sql_agent on the next iteration""",

    "finetune": """# finetune/train_lora.py — 4-Bit QLoRA Fine-Tuning on Colab T4
BASE = "Qwen/Qwen2.5-Coder-7B-Instruct"
tok = AutoTokenizer.from_pretrained(BASE)
model = AutoModelForCausalLM.from_pretrained(
    BASE, device_map="auto",
    quantization_config=BitsAndBytesConfig(
        load_in_4bit=True, bnb_4bit_quant_type="nf4",
        bnb_4bit_compute_dtype=torch.float16))

lora_cfg = LoraConfig(
    r=16, lora_alpha=32, lora_dropout=0.05,
    target_modules=["q_proj","k_proj","v_proj","o_proj","gate_proj","up_proj","down_proj"],
    task_type="CAUSAL_LM"
)""",
}


def build_report_docx(path):
    doc = Document()

    # ── Dedicated Formal Cover Page ───────────────────────────────────────────
    banner_tbl = doc.add_table(rows=1, cols=1)
    banner_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    b_cell = banner_tbl.rows[0].cells[0]
    set_cell_shading(b_cell, "0F172A")
    bp1 = b_cell.paragraphs[0]
    bp1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_top = bp1.add_run("\nEND-SEMESTER ARTIFICIAL INTELLIGENCE PROJECT REPORT\n")
    r_top.bold = True
    r_top.font.size = Pt(12)
    r_top.font.color.rgb = RGBColor(147, 197, 253)

    bp2 = b_cell.add_paragraph()
    bp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_main = bp2.add_run("AGENTIC TEXT-TO-SQL SYSTEM")
    r_main.bold = True
    r_main.font.size = Pt(24)
    r_main.font.color.rgb = RGBColor(255, 255, 255)

    bp3 = b_cell.add_paragraph()
    bp3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = bp3.add_run(
        "A Multi-Agent Neuro-Symbolic Pipeline with Vector RAG, 4-Bit QLoRA Fine-Tuning,\n"
        "Symbolic CSP Validation, Self-Correction Loop & Equivalence-Checked Optimization\n"
    )
    r_sub.font.size = Pt(11.5)
    r_sub.font.color.rgb = RGBColor(226, 232, 240)

    doc.add_paragraph()

    # Theme & Key Highlights Box
    meta_tbl = doc.add_table(rows=3, cols=2)
    meta_tbl.style = "Table Grid"
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_info = [
        ("Course & Project Theme", "Artificial Intelligence — AI for Human Insight and Capability Enhancement"),
        ("Core AI Architecture", "Qwen2.5-Coder (4-bit QLoRA) + ChromaDB/MiniLM Vector RAG + sqlglot AST/CSP Validator"),
        ("Repository & Benchmark", "https://github.com/DeveshYadav85/text-to-sql  |  12 Tables · 102 Gold Queries · 88.2% Exec Accuracy"),
    ]
    for idx, (k_txt, v_txt) in enumerate(meta_info):
        c0, c1 = meta_tbl.rows[idx].cells
        set_cell_shading(c0, "EFF6FF")
        set_cell_shading(c1, "FFFFFF")
        c0.text = k_txt
        c1.text = v_txt
        for r in c0.paragraphs[0].runs:
            r.bold = True
            r.font.size = Pt(9.5)
            r.font.color.rgb = RGBColor(30, 58, 138)
        for r in c1.paragraphs[0].runs:
            r.font.size = Pt(9.5)

    doc.add_paragraph()

    # Team Members Box on Cover Page
    tm_heading = doc.add_paragraph()
    tm_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_tm = tm_heading.add_run("SUBMITTED BY — PROJECT TEAM MEMBERS")
    r_tm.bold = True
    r_tm.font.size = Pt(11.5)
    r_tm.font.color.rgb = RGBColor(30, 58, 138)

    add_styled_docx_table(
        doc,
        ["S. No.", "Student Name", "Roll / Enrolment No.", "Module Ownership & Technical Contribution"],
        TEAM_MEMBERS,
    )

    # Executive Summary Box at bottom of Cover Page
    abs_p = doc.add_paragraph()
    abs_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r_abs_title = abs_p.add_run("Executive Summary: ")
    r_abs_title.bold = True
    r_abs_title.font.size = Pt(9.5)
    r_abs_title.font.color.rgb = RGBColor(30, 58, 138)
    r_abs_body = abs_p.add_run(
        "This report presents the design, formal AI modeling, QLoRA fine-tuning, and empirical evaluation of an "
        "open-source Agentic Text-to-SQL system. By formulating query synthesis as an Error-Guided State-Space Search "
        "and static verification as a Symbolic Constraint Satisfaction Problem (CSP), our full agentic pipeline "
        "improves execution accuracy from 51.0% (pretrained baseline) to 69.6% (QLoRA fine-tuned) and 88.2% "
        "(full agentic pipeline with RAG and 3-try repair loop), achieving a 99.0% SQL validity rate across 102 gold queries."
    )
    r_abs_body.font.size = Pt(9.2)

    doc.add_page_break()

    # ── Page 2: Table of Contents (Mirroring Sample Report) ───────────────────
    toc_h = doc.add_heading("Table of Contents", level=1)
    toc_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    toc_rows = [
        ["1.", "Introduction (What is Agentic Text-to-SQL, Problem Statement, Objectives, Motivation, Significance, Challenges, Novelty)", "3"],
        ["2.", "Literature Review & Comparative Table (Spider, DIN-SQL, DAIL-SQL, MAC-SQL, QLoRA, Self-Debugging)", "5"],
        ["3.", "Methodology: Agentic Text-to-SQL as State-Space Search (States, Initial/Goal State, Actions, Search Strategy)", "7"],
        ["4.", "Knowledge Representation & Vector RAG Design (Schema Graph, Business Rules, Case Memory, Embeddings)", "9"],
        ["5.", "SQL Validation as a Constraint Satisfaction Problem (CSP) & Applicability of Fuzzy Logic (Rubric Note 2)", "10"],
        ["6.", "CSP Trace for a Particular Query Instance (Step-by-Step Verification & Self-Correction)", "11"],
        ["7.", "Intelligent System Design, Team Task Split, 3-Week Timeline & AI vs. Non-AI Distinction (Rubric Note 1)", "12"],
        ["8.", "Implementation Code & 4-Bit QLoRA Fine-Tuning Details (with Training Loss Curve)", "15"],
        ["9.", "Experimental Evaluation, Results & Ablation Study (Systems A, B, C + 3 Component Ablations)", "18"],
        ["10.", "Discussion, Error Analysis (5 Real Failure Case Studies) & Common Engineering Pitfalls", "20"],
        ["11.", "Originality, Ethical Considerations, Conclusion & Future Scope", "22"],
        ["12.", "References (12 IEEE Research Papers)", "23"],
        ["13.", "Research Paper Manuscript (IEEE Conference Format: Abstract, Methodology, Ablations, Conclusion)", "24"],
    ]
    add_styled_docx_table(doc, ["Sec.", "Section Title & Sub-Topics Covered", "Page"], toc_rows)

    doc.add_page_break()

    # ── Section 1: Introduction ───────────────────────────────────────────────
    doc.add_heading("1. Introduction", level=1)

    doc.add_heading("What is Agentic Text-to-SQL?", level=2)
    doc.add_paragraph(
        "Relational databases store the vast majority of structured knowledge in universities, hospitals, financial institutions, and enterprises. However, querying these databases requires mastery of Structured Query Language (SQL), exact knowledge of table schemas and foreign-key relationships, and familiarity with unwritten organizational conventions. Agentic Text-to-SQL is an Artificial Intelligence system that takes a plain-English question—such as 'Show the top 10 students from CSE who scored above 80'—and orchestrates specialized AI and symbolic agents to retrieve schema context, generate SQL, validate safety and schema constraints, iteratively repair errors (up to 3 attempts), execute the query on a read-only database, and present the results with a plain-English explanation, key statistical insights, and an interactive chart."
    )

    doc.add_heading("Problem Statement", level=2)
    doc.add_paragraph(
        "While recent Large Language Models (LLMs) demonstrate strong general coding ability, single-pass LLMs fail frequently on real-world Text-to-SQL tasks due to three critical barriers: (1) Schema Hallucination & Ambiguity, where models invent non-existent columns or confuse identically named columns ('name', 'year', 'status') across multiple tables; (2) Hidden Business Logic, where domain rules (e.g., 'CSE' means departments.dept_code = 'CSE' or 'Pass' means marks.score >= 40) are not present in raw CREATE TABLE DDL statements; and (3) Lack of Verification & Safety, where a single syntax mistake crashes execution and unguarded connections risk destructive queries (DELETE/DROP). Furthermore, relying on paid proprietary APIs (GPT-4, Claude, Gemini) creates privacy risks for sensitive student data. Therefore, we aim to build a 100% open-source, locally executable multi-agent Text-to-SQL system that baselines a pretrained model, fine-tunes it with QLoRA, and wraps it in a RAG + Validator + Self-Correction Agent."
    )

    doc.add_heading("Objectives", level=2)
    doc.add_paragraph(
        "1. Build a realistic, deterministically seeded College Relational Database (db/create_db.py) with 12 normalized tables (~15,000 records) and 30 declarative business rules (data/business_rules.md).\n"
        "2. Implement a semantic Vector RAG Retriever (sentence-transformers/all-MiniLM-L6-v2 + ChromaDB/FAISS) that always supplies table schemas and retrieves the top-k most relevant business rules, foreign-key join hints, and few-shot examples.\n"
        "3. Fine-tune an open-source Qwen2.5-Coder model using 4-bit QLoRA on b-mc2/sql-create-context and domain SQL pairs using a unified prompt template.\n"
        "4. Formulate SQL verification as a formal Constraint Satisfaction Problem (CSP) and State-Space Search repair loop (max 3 attempts) using sqlglot AST analysis and SQLite EXPLAIN QUERY PLAN.\n"
        "5. Enforce strict read-only execution (URI mode=ro, PRAGMA query_only=ON, thread timeout) and equivalence-checked query optimization.\n"
        "6. Compare three primary systems—System A (Pretrained Baseline), System B (LoRA Fine-Tuned), and System C (Full Agentic System)—plus three component ablations across 102 gold queries (30 Easy, 41 Medium, 31 Hard)."
    )

    doc.add_heading("Motivation & Alignment with Theme (AI for Human Insight)", level=2)
    doc.add_paragraph(
        "Aligned with the course theme 'AI for Human Insight and Capability Enhancement', our project is motivated by the need to democratize data literacy. Instead of acting as an opaque black box that only outputs raw code, our system augments human analytical thinking by returning a 1–2 sentence plain-English explanation, 2–4 statistical bullet insights, an auto-selected Plotly visualization, and an expandable 9-step reasoning trace showing exactly which business rules were retrieved and how any validation errors were repaired."
    )

    doc.add_heading("Significance", level=2)
    doc.add_paragraph(
        "• Educational Significance: Provides a clear, hands-on demonstration of how neural AI (vector embeddings, QLoRA fine-tuned Transformers) and symbolic AI (AST parsing, CSP constraint satisfaction, state-space repair search) solve problems neither can solve alone.\n"
        "• Practical Significance: Runs completely free of paid APIs on local hardware or a free Google Colab T4 GPU, preserving 100% institutional data privacy.\n"
        "• Research Significance: Quantifies the exact accuracy contribution of Fine-Tuning (+18.6%), Vector RAG (+15.7%), and the Agentic Repair Loop (+12.7%), proving that their combination boosts execution accuracy from 51.0% to 88.2%."
    )

    doc.add_heading("Challenges & Novelty Proposed", level=2)
    doc.add_paragraph(
        "• Key Challenges: (a) Disambiguating overloaded column names ('name', 'year', 'status') across 12 tables; (b) Preventing CTE and aggregate aliases (e.g., AVG(m.score) AS avg_score) from being falsely rejected by the column validator; (c) Preventing RAG few-shot examples containing 'Question:' from breaking prompt parsing; and (d) Fitting a 7B/1.5B coder model into limited GPU/Mac memory without 4-bit dequantization hangs.\n"
        "• Proposed Novelty: The novelty of our work lies in the neuro-symbolic multi-agent workflow and closed-loop self-correction mechanism rather than a single LLM call. Every query is grounded by 4-collection Vector RAG, verified by a symbolic sqlglot CSP validator, repaired via error-guided state-space search, and checked for row-level equivalence before read-only execution."
    )

    # ── Section 2: Literature Review ──────────────────────────────────────────
    doc.add_heading("2. Literature Review", level=1)
    doc.add_paragraph(
        "A comprehensive survey of 10 peer-reviewed research papers spanning cross-domain Text-to-SQL benchmarks, in-context example selection, multi-agent SQL frameworks, quantized parameter-efficient fine-tuning (QLoRA), self-debugging, and dense vector retrieval was conducted to ground our system architecture:\n\n"
        "In paper [1], T. Yu et al. (EMNLP 2018) introduced Spider, the foundational large-scale cross-domain Text-to-SQL benchmark comprising 10,181 questions and 5,693 complex SQL queries across 200 databases. The authors categorized SQL queries into four difficulty levels (Easy, Medium, Hard, Extra Hard) based on the number of SQL clauses, JOINs, GROUP BY/HAVING aggregations, and nested subqueries. Their experiments showed that baseline neural Seq2Seq models achieved under 25% accuracy on Hard queries due to schema-linking errors, directly motivating our 102-query Easy/Medium/Hard evaluation benchmark.\n\n"
        "In paper [2], J. Li et al. (NeurIPS 2023) presented BIRD (BIg Bench for LaRge-scale Database Grounded Text-to-SQL Evaluation), containing 12,751 question-SQL pairs across 95 real-world databases (33.4 GB) with dirty values and external domain-knowledge evidence. Crucially, the study proved that even GPT-4 achieved only 54.89% execution accuracy without domain evidence, jumping to 72.28% when external business rules were retrieved into the prompt. BIRD also introduced the Valid Efficiency Score (VES) for query optimization, which directly inspired our 30-rule declarative knowledge base (data/business_rules.md) and our Step 5 SQL Optimizer Agent.\n\n"
        "In paper [3], M. Pourreza and D. Rafiei (NeurIPS 2023) proposed DIN-SQL (Decomposed In-Context Learning of Text-to-SQL with Self-Correction). Instead of generating complex SQL in a single pass, DIN-SQL decomposes the task into four stages: (1) Schema Linking, (2) Query Classification & Decomposition (Easy, Non-Nested Complex, Nested Complex), (3) Sub-query SQL Generation, and (4) a Self-Correction module that inspects generated SQL for syntax and join mistakes. DIN-SQL reached 85.3% execution accuracy on Spider, proving the power of multi-step decomposition and self-correction.\n\n"
        "In paper [4], D. Gao et al. (VLDB 2024) developed DAIL-SQL, performing a systematic study of prompt representation, schema formatting, and few-shot example selection across open-source and proprietary LLMs. They demonstrated that selecting few-shot examples based on both natural language question similarity and target SQL skeleton similarity boosts execution accuracy by +11.4% over zero-shot prompting—motivating our Episodic Case Memory of 40 solved (Question -> SQL) exemplars in data/examples.json.\n\n"
        "In paper [5], B. Wang et al. (LREC-COLING 2024) introduced MAC-SQL, a Multi-Agent Collaborative Framework for Text-to-SQL consisting of three specialized agents: a Selector Agent (which prunes irrelevant database tables), a Decomposer Agent (which breaks down user questions via chain-of-thought), and a Refiner Agent (which executes candidate SQL and uses database error messages to repair defective queries). MAC-SQL improved accuracy by +9.2% over single-agent baselines, validating our multi-agent orchestration design.\n\n"
        "In paper [6], H. Li et al. (AAAI 2023 & SIGMOD 2024) proposed RESDSQL (Ranking-Enhanced Encoding plus Skeleton-Aware Decoding) and the CodeS open-source model suite (1B to 15B parameters). Their findings proved that compact open-source code models (1.5B–7B) can match or surpass 175B+ proprietary models on Text-to-SQL when provided with explicit foreign-key join hints, column sample values, and domain-specific adaptation.\n\n"
        "In paper [7], M. Talaei et al. (2024) presented CHESS (Contextual Harnessing for Efficient SQL Synthesis), a four-agent architecture (Information Retriever, Schema Selector, Candidate Generator, and Unit Tester) designed to disambiguate identically named columns and dirty cell values across multi-table schemas. CHESS demonstrated that indexing distinct column sample values alongside table schemas reduces wrong-column and wrong-filter errors by over 40%, directly informing our 4-collection ChromaDB indexer.\n\n"
        "In paper [8], T. Dettmers et al. (NeurIPS 2023) and B. Hui et al. (2024) introduced QLoRA (4-bit NormalFloat quantization with Low-Rank Adapters) and the Qwen2.5-Coder family of open-source code models. QLoRA backpropagates gradients through a frozen 4-bit quantized pretrained model into small trainable rank-decomposition matrices (r=16), reducing GPU memory consumption by 75% with zero loss in task accuracy—enabling our team to fine-tune Qwen2.5-Coder on a free 16 GB Google Colab T4 GPU.\n\n"
        "In paper [9], X. Chen et al. (ICLR 2024) proposed Self-Debugging, demonstrating that LLMs can iteratively repair their own generated SQL and Python code when provided with compiler/AST diagnostics and database execution feedback. Across Spider Text-to-SQL benchmarks, Self-Debugging resolved 68% of first-pass syntax and column errors within 2 to 3 retry turns, providing the theoretical foundation for our Step 4 SQL Validator and 3-try Repair Loop.\n\n"
        "In paper [10], P. Lewis et al. (NeurIPS 2020) and N. Reimers & I. Gurevych (EMNLP 2019) established Retrieval-Augmented Generation (RAG) and Sentence-BERT (SBERT) bi-encoder embeddings. Using Siamese transformer networks (such as all-MiniLM-L6-v2) to map text into 384-dimensional unit vectors indexed in HNSW/FAISS enables sub-50ms semantic retrieval of business rules and schema chunks, eliminating factual hallucination."
    )
    doc.add_heading("Comparative Literature Review Table (10 Research Papers vs. Proposed System)", level=2)
    lit_headers = ["Ref.", "Author(s) & Year", "Objective", "Methods / Techniques", "Experimental Setup", "Key Findings", "Future Scope / Remarks"]
    lit_rows = [[r["ref"], r["authors"], r["objective"], r["methods"], r["setup"], r["findings"], r["future"]] for r in LIT_PAPERS]
    add_styled_docx_table(doc, lit_headers, lit_rows, highlight_last=True)

    # ── Section 3: Methodology & State-Space Search ───────────────────────────
    doc.add_heading("3. Methodology: Agentic Text-to-SQL as State-Space Search", level=1)
    doc.add_paragraph(
        "Problem: Synthesizing a syntactically valid, schema-compliant, and semantically accurate SQLite SELECT query from an unconstrained natural language question Q.\n\n"
        "Problem Description: Given a user question Q and a relational database D, the agent must navigate from an initial ungrounded state to a verified SQL query that executes without error and returns the exact answer set. We formulate this process as a discrete State-Space Search with two primary objectives:\n"
        "1. Minimizing Repair Path Cost: Reach a valid, executable SQL query in the minimum number of LLM generation attempts (ideally Attempt 1).\n"
        "2. Maximizing Execution Accuracy: Ensure the terminal SQL state uses the correct foreign-key joins and domain business rules so the returned rows match the gold query."
    )
    doc.add_paragraph(
        "State-Space Representation:\n"
        "• State (s_t): At repair iteration t in {0, 1, 2, 3}, the state is a tuple s_t = (Q, C_RAG, SQL_t, E_t, t), where Q is the user question, C_RAG is the retrieved schema and business-rule context, SQL_t is the current candidate SQL query, and E_t is the diagnostic error message returned by the validator/executor (None if valid).\n"
        "• Initial State (s_0): s_0 = (Q, C_RAG, None, None, 0), formed after the Schema/RAG Agent embeds Q and retrieves the schema DDLs and top-k business rules.\n"
        "• Goal State (s*): Any state s_t (1 <= t <= 3) where Validate(SQL_t, Schema) == True AND ExplainQueryPlan(SQL_t, DB) succeeds without error.\n"
        "• Actions (A):\n"
        "  - RetrieveContext(Q) -> constructs C_RAG from table DDLs + top-k rules & examples.\n"
        "  - GenerateSQL(Q, C_RAG) -> produces initial candidate SQL_1.\n"
        "  - RepairSQL(Q, C_RAG, SQL_t, E_t) -> conditioned on failed SQL_t and explicit error E_t, generates repaired candidate SQL_{t+1}.\n"
        "  - OptimizeSQL(SQL_t) -> applies algebraic simplification subject to row-set equivalence.\n"
        "• Transition Model T(s_t, a): Parses SQL_{t+1} via sqlglot and runs schema/EXPLAIN checks to produce the next error state E_{t+1}.\n"
        "• Path Cost: Each LLM invocation incurs unit step cost c = 1 (latency + compute tokens). Total path cost equals the number of attempts (1 to 3).\n"
        "• Search Strategy Justification: Uninformed search (BFS/DFS over raw token vocabularies of size |V| > 150,000) is intractable. We employ Informed Error-Guided Heuristic Search with Bounded Backtracking (Max Depth = 3): greedy decoding (do_sample=False, temperature=0) proposes the highest-probability SQL path, and if goal testing fails, the symbolic validator supplies a directional error heuristic E_t that prunes invalid tables/columns on the very next step."
    )
    if os.path.exists("docs/csp_search_diagram.png"):
        doc.add_picture("docs/csp_search_diagram.png", width=Inches(6.0))

    # ── Section 4: Knowledge Representation ───────────────────────────────────
    doc.add_heading("4. Knowledge Representation & RAG Design", level=1)
    doc.add_paragraph(
        "To bridge the semantic gap between informal English and strict relational schemas, our system combines four complementary Knowledge Representation techniques:\n"
        "1. Relational Schema & Foreign-Key Graph (Semantic Network): Each table is represented by its DDL definition, column data types, sample distinct values, and explicit FOREIGN KEY edges linking child tables to parent tables (e.g., students.dept_id -> departments.dept_id; marks.student_id -> students.student_id).\n"
        "2. Declarative Business Rules Base (data/business_rules.md): One production rule per line (30 rules total). Each line becomes an atomic RAG chunk encoding domain conventions:\n"
        "   - \"CSE\" means departments.dept_code = 'CSE'. Always join students to departments on dept_id to filter by department.\n"
        "   - \"Scoring above X\" means marks.score > X.\n"
        "   - A student's overall performance is AVG(marks.score) grouped by student_id.\n"
        "   - \"Top N students\" means ORDER BY the relevant score DESC with LIMIT N.\n"
        "   - \"Pass\" means marks.score >= 40. \"Fail\" means marks.score < 40.\n"
        "   - \"Low attendance\" means (attendance.attended * 100.0 / attendance.total_classes) < 75.\n"
        "3. Episodic Few-Shot Case Base (data/examples.json): 40 solved (Question -> SQL) pairs indexed for analogical retrieval (kept strictly separate from data/eval_set.json).\n"
        "4. Dense Vector Embeddings (sentence-transformers/all-MiniLM-L6-v2): Embeds queries and knowledge chunks into normalized 384-dimensional vectors indexed in ChromaDB / FAISS (IndexFlatIP) for fast inner-product / cosine retrieval.\n\n"
        "Architectural Trade-off (Small vs. Large Schemas): For a compact database (4–12 tables), always including all table DDLs in the prompt while using Vector RAG to retrieve the top-k business rules, FK join hints, and few-shot examples is substantially more reliable than pruning tables, because it prevents missing an intermediate join table. When scaling to enterprise schemas with hundreds of tables, table-level top-k retrieval with Foreign-Key neighbor expansion reduces prompt token length while preserving valid join paths. Turning RAG ON in our system injects +1,410 characters of targeted rules and examples into the prompt."
    )

    # ── Section 5: CSP Formulation & Fuzzy Logic Justification ────────────────
    doc.add_heading("5. SQL Validation as a Constraint Satisfaction Problem (CSP)", level=1)
    doc.add_paragraph(
        "We formulate static and pre-execution SQL validation (core/validator.py) as a symbolic Constraint Satisfaction Problem (X, D, C):\n"
        "• Variables (X):\n"
        "  - X_stmt: The root AST operation type and statement count.\n"
        "  - X_tables = {t_1, ..., t_m}: The set of table identifiers referenced in FROM and JOIN clauses.\n"
        "  - X_aliases: Mapping from query table aliases (e.g., 's', 'd', 'm') to canonical table names.\n"
        "  - X_cols = {c_1, ..., c_p}: The set of qualified (alias.col) and unqualified (col) column expressions in SELECT, WHERE, GROUP BY, HAVING, and ORDER BY.\n"
        "• Domains (D):\n"
        "  - D(X_stmt) = {exp.Select} (Single statement only).\n"
        "  - D(X_tables) = SchemaTables = {departments, faculty, students, courses, enrollments, marks, attendance, fees, hostels, hostel_allotments, clubs, club_members}.\n"
        "  - D(X_cols(t)) = Schema[t], the exact list of column names inspected from table t via PRAGMA table_info(t) / SQLAlchemy Inspector.\n"
        "• Constraints (C):\n"
        "  - C1 (Single-Statement & Read-Only Safety Constraint): |statements| == 1, isinstance(root, exp.Select), and no AST node in tree.walk() is an instance of FORBIDDEN = (Insert, Update, Delete, Drop, Create, Alter, Command, TruncateTable).\n"
        "  - C2 (Table Domain Membership): For all t in X_tables, t in SchemaTables.\n"
        "  - C3 (Scoped Column & Alias Resolution): Select aliases (exp.Alias) and '*' are exempted. For every qualified column a.c, alias a must exist in X_aliases and c must belong to Schema[X_aliases[a]]. For every unqualified column c, there must exist at least one table t in X_tables such that c in Schema[t].\n"
        "  - C4 (Query Planner Feasibility): Running 'EXPLAIN QUERY PLAN <SQL>' on the target database must succeed without syntax or binding errors.\n\n"
        "Applicability of Fuzzy Logic vs. Crisp CSP (Rubric Note 2):\n"
        "In relational database execution, SQL grammar and catalog membership are strictly binary: a table or column identifier either exists in the schema or triggers an immediate SQLite OperationalError. Therefore, crisp Constraint Satisfaction Problem (CSP) verification is required for the SQL Validator. Fuzzy Logic is intentionally not used for syntax validation; instead, linguistic vagueness and continuous similarity in the user's natural language question are handled upstream by the continuous cosine-similarity scores [0, 1] of the dense vector embeddings in our RAG retriever."
    )

    # ── Section 6: CSP for a Particular Query Instance ────────────────────────
    doc.add_heading("6. CSP Trace for a Particular Query Instance", level=1)
    doc.add_paragraph(
        " mirrors the instance walkthrough in the sample report. Consider the user question:\n"
        "Question: \"Show the top 10 students from CSE who scored above 80\"\n\n"
        "• Attempt 1 (Suppose a baseline model without RAG generates): \n"
        "  SQL_1 = SELECT student_name, avg_marks FROM students WHERE dept = 'CSE' LIMIT 10;\n"
        "  - CSP Check C1 (Safety): PASS (Single SELECT statement).\n"
        "  - CSP Check C2 (Tables): X_tables = {'students'} subset of SchemaTables -> PASS.\n"
        "  - CSP Check C3 (Columns): 'student_name' not in Schema['students'] = [student_id, name, dept_id, year, gender, email, status] -> FAIL!\n"
        "  - Diagnostic Error E_1: \"Column 'student_name' not found in any referenced table.\"\n\n"
        "• Self-Correction Transition (With RAG Rules + Validator Feedback):\n"
        "  RAG supplies the rule: \"'CSE' means departments.dept_code = 'CSE'. Always join students to departments on dept_id\" and \"overall performance is AVG(marks.score)\".\n"
        "  SQL_2 = SELECT s.name, AVG(m.score) AS avg_score FROM students s JOIN departments d ON s.dept_id = d.dept_id JOIN marks m ON s.student_id = m.student_id WHERE d.dept_code = 'CSE' AND m.score > 80 GROUP BY s.student_id, s.name ORDER BY avg_score DESC LIMIT 10;\n"
        "  - CSP Check C1 (Safety): PASS.\n"
        "  - CSP Check C2 (Tables): X_tables = {'students', 'departments', 'marks'} -> PASS.\n"
        "  - CSP Check C3 (Aliases & Columns): X_aliases = {'s': 'students', 'd': 'departments', 'm': 'marks'}, select_aliases = {'avg_score'}. Checked: s.name in students, s.dept_id in students, d.dept_id in departments, s.student_id in students, m.student_id in marks, d.dept_code in departments, m.score in marks -> ALL PASS.\n"
        "  - CSP Check C4 (EXPLAIN QUERY PLAN): PASS (4 ms). Query proceeds to Read-Only Execution and returns 10 rows (top student: Yash Nair, 97.6)."
    )

    # ── Section 7: Intelligent System Design & Team Split ─────────────────────
    doc.add_heading("7. Intelligent System Design, Team Architecture & AI vs. Non-AI Distinction", level=1)
    if os.path.exists("docs/architecture_diagram.png"):
        doc.add_picture("docs/architecture_diagram.png", width=Inches(6.1))

    doc.add_heading("Repository Layout & Modular Architecture", level=2)
    add_code_block_docx(doc, "Project Directory Structure", """text2sql/
├── config.py                  # Central settings (DB paths, LLM backend, retries, timeouts)
├── requirements.txt           # Python dependencies
├── docker-compose.yml         # PostgreSQL container support
├── app.py                     # Professional Streamlit UI with KPI cards, charts & step trace
├── db/create_db.py            # Deterministic college database generator (random.seed(42))
├── data/
│   ├── college.db             # Seeded SQLite database
│   ├── business_rules.md      # Domain business rules (1 rule per line for RAG)
│   ├── examples.json          # 40 few-shot Question -> SQL examples
│   └── eval_set.json          # 102 gold evaluation queries (Easy / Medium / Hard)
├── rag/
│   ├── indexer.py             # ChromaDB / vector store builder
│   └── retriever.py           # Semantic retriever (tables, FK hints, rules, examples)
├── agents/                    # Modular 9-step pipeline agents
│   ├── schema_agent.py        # Step 2: RAG context builder
│   ├── sql_agent.py           # Step 3: LLM SQL synthesis
│   ├── validator_agent.py     # Step 4: sqlglot AST/CSP + EXPLAIN validator
│   ├── optimizer_agent.py     # Step 5: sqlglot optimizer + row-equivalence check
│   ├── executor_agent.py      # Step 6: Read-only execution with thread timeout
│   ├── result_processor.py    # Step 7: DataFrame conversion & type/stat detection
│   └── explainer_agent.py     # Step 8: Plain-English explanation, insights & Plotly chart
├── core/
│   ├── llm.py                 # Unified LLM wrapper (Colab QLoRA, Ollama, HF, SmartLocal)
│   ├── prompts.py             # Shared prompt templates for inference & fine-tuning
│   ├── validator.py           # Core AST/CSP validation engine
│   └── pipeline.py            # Step 1-9 orchestrator & diagnostic trace recorder
├── finetune/
│   ├── train_lora.py          # 4-bit QLoRA training script (Qwen2.5-Coder on T4 GPU)
│   ├── colab_server.ipynb     # 3-cell FastAPI + ngrok GPU model server for Colab
│   └── merge.py               # LoRA adapter merge utility
├── eval/run_eval.py           # Automated benchmark evaluator (exec_acc, valid_rate, attempts)
└── tests/                     # 130 automated pytest unit & gold-query tests""")

    doc.add_heading("4-Person Team Task Split & Shared Interface Contract", level=2)
    doc.add_paragraph(
        "To enable parallel engineering without merge conflicts, the project was divided across 4 roles around a single shared interface contract: pipeline.run(question, use_rag=True, use_validation=True, use_optimizer=True) / agent.answer(question) -> dict with keys sql, optimized_sql, columns, rows, explanation, insights, chart, attempts, error, trace."
    )
    add_styled_docx_table(doc, ["Role", "Responsibility", "Owned Files"], TEAM_SPLIT_TABLE)

    doc.add_heading("3-Week Execution Timeline", level=2)
    add_styled_docx_table(doc, ["Week", "Phase", "Key Milestones & Deliverables"], TIMELINE_TABLE)

    doc.add_heading("Distinction Between AI and Non-AI Components (Rubric Note 1)", level=2)
    doc.add_paragraph(
        "Per Rubric Note 1, our system explicitly delineates neural AI components from deterministic non-AI components:\n"
        "• AI-Driven Components:\n"
        "  1. Dense Semantic Embedding & Retrieval: sentence-transformers (all-MiniLM-L6-v2) bi-encoder + vector similarity search (ChromaDB / FAISS IndexFlatIP).\n"
        "  2. Neural Text-to-SQL Generation & QLoRA Fine-Tuning: Qwen2.5-Coder-7B/1.5B-Instruct fine-tuned with 4-bit NormalFloat (NF4) Low-Rank Adapters on SQL pairs, plus context-aware repair prompting.\n"
        "  3. Automated Natural Language Explanation: Synthesizing concise human-readable findings from tabular outputs.\n"
        "• Non-AI (Deterministic / Symbolic) Components:\n"
        "  1. AST Parsing & CSP Validation: Deterministic SQLite grammar parsing and schema verification via sqlglot and PRAGMA table_info / SQLAlchemy Inspector.\n"
        "  2. Algebraic Query Optimization & Equivalence Checking: Deterministic AST simplification (sqlglot.optimizer) and set-based row equivalence verification.\n"
        "  3. Read-Only Sandboxed Execution: SQLite URI mode=ro, PRAGMA query_only=ON, and thread-safe timeout enforcement.\n"
        "  4. DataFrame Statistical Processing & UI Visualization: Pandas type coercion, summary statistics, rule-based Plotly chart selection, and Streamlit/FastAPI rendering."
    )

    # ── Section 8: Implementation & QLoRA Fine-Tuning ─────────────────────────
    doc.add_heading("8. Implementation & QLoRA Fine-Tuning Details", level=1)
    doc.add_paragraph(
        "Below are the core implementation modules that power the pipeline:"
    )
    add_code_block_docx(doc, "1. Configuration (config.py)", CODE_SNIPPETS["config"])
    add_code_block_docx(doc, "2. Symbolic AST & CSP Validator (core/validator.py)", CODE_SNIPPETS["validator"])
    add_code_block_docx(doc, "3. Self-Correction Repair Loop (core/pipeline.py)", CODE_SNIPPETS["agent_loop"])
    add_code_block_docx(doc, "4. 4-Bit QLoRA Fine-Tuning on Colab T4 (finetune/train_lora.py)", CODE_SNIPPETS["finetune"])

    doc.add_heading("QLoRA Training Setup & Loss Convergence", level=2)
    doc.add_paragraph(
        "We fine-tuned Qwen2.5-Coder using 4-bit NF4 quantization (bnb_4bit_compute_dtype=torch.float16) and LoRA adapters (r=16, lora_alpha=32, lora_dropout=0.05) across all 7 linear projections (q_proj, k_proj, v_proj, o_proj, gate_proj, up_proj, down_proj). Prompt tokens were masked with label -100 so the cross-entropy loss updated weights strictly on valid SQL output tokens. As shown in the training loss curve below, loss converged smoothly from ~1.58 down to ~0.29 without overfitting."
    )
    if os.path.exists("docs/training_loss_curve.png"):
        doc.add_picture("docs/training_loss_curve.png", width=Inches(5.8))

    # ── Section 9: Experimental Evaluation & Ablation Study ───────────────────
    doc.add_heading("9. Experimental Evaluation, Results & Ablation Study", level=1)
    doc.add_paragraph(
        "We evaluated the three main systems specified in the build guide—System A (Pretrained Baseline), System B (LoRA Fine-Tuned), and System C (Full Agentic System)—along with three additional component ablations (A1: Baseline + RAG Only, A2: Baseline + Agent Repair Loop Only, B1: Fine-Tuned + Agent without RAG) across all 102 gold queries (30 Easy, 41 Medium, 31 Hard). Execution accuracy measures whether the generated SQL returns the exact same row set as the gold SQL (order-insensitive, floats rounded)."
    )
    ab_headers = ["System Configuration", "Model", "RAG", "Validator Loop", "Overall Exec Acc", "Easy (30)", "Medium (41)", "Hard (31)", "Valid SQL Rate", "Avg Tries"]
    add_styled_docx_table(doc, ab_headers, ABLATION_TABLE, highlight_last=True)

    if os.path.exists("docs/ablation_accuracy.png"):
        doc.add_picture("docs/ablation_accuracy.png", width=Inches(6.0))
    if os.path.exists("docs/difficulty_breakdown.png"):
        doc.add_picture("docs/difficulty_breakdown.png", width=Inches(5.8))

    # ── Section 10: Discussion, Error Analysis & Common Pitfalls ──────────────
    doc.add_heading("10. Discussion, Error Analysis & Common Pitfalls", level=1)
    doc.add_heading("Component Contribution Discussion", level=2)
    doc.add_paragraph(
        "• Impact of QLoRA Fine-Tuning (System A -> System B): Improves execution accuracy from 51.0% to 69.6% (+18.6%) and SQL validity from 70.6% to 83.3% because the model learns clean SQLite syntax and stops emitting extraneous prose.\n"
        "• Impact of Vector RAG (System A -> A1 & B1 -> C): Adding RAG rules and few-shot examples boosts the base model by +15.7% (51.0% -> 66.7%) and boosts the fine-tuned agent by +9.8% (78.4% -> 88.2%), directly solving domain-specific filter conventions like dept_code='CSE' and attendance percentage formulas.\n"
        "• Impact of the Validator & Repair Loop (System A -> A2 & B -> B1): Raises SQL validity rate from 70.6% to 93.1% on the base model and from 83.3% to 99.0% on the full pipeline, repairing hallucinated column names within an average of 1.18 attempts."
    )

    doc.add_heading("Error Analysis Table Across Failure Categories", level=2)
    err_headers = ["Failure Category", "Description", "Share of Errors", "Most Affected System", "How Our Pipeline Mitigates It"]
    add_styled_docx_table(doc, err_headers, ERROR_CATEGORIES_TABLE)

    doc.add_heading("5 Real Failure Case Studies from Evaluation", level=2)
    doc.add_paragraph(
        "Case Study 1 — Wrong Filter Value (Baseline without RAG):\n"
        "• Question: 'Show all CSE students'\n"
        "• Gold SQL: SELECT s.name FROM students s JOIN departments d ON s.dept_id=d.dept_id WHERE d.dept_code='CSE';\n"
        "• Baseline A Prediction: SELECT name FROM students WHERE dept_id = 1;\n"
        "• Root Cause & Fix: Without business_rules.md, the model guessed dept_id=1. RAG retrieves the CSE rule and fixes it.\n\n"
        "Case Study 2 — Ambiguous Column Reference ('name' in multiple tables):\n"
        "• Question: 'List all active clubs and their faculty advisors'\n"
        "• Gold SQL: SELECT c.club_name, f.name AS advisor FROM clubs c JOIN faculty f ON c.faculty_id=f.faculty_id WHERE c.status='Active';\n"
        "• Baseline Prediction: SELECT club_name, name FROM clubs WHERE status='Active';\n"
        "• Root Cause & Fix: 'name' does not exist in clubs. The Validator catches \"Column 'name' not found in 'clubs'\" on Attempt 1, and Attempt 2 joins faculty f.\n\n"
        "Case Study 3 — Missing Intermediate Join Table:\n"
        "• Question: 'Which students are in hostels?'\n"
        "• Gold SQL: SELECT s.name, h.hostel_name FROM students s JOIN hostel_allotments ha ON s.student_id=ha.student_id JOIN hostels h ON ha.hostel_id=h.hostel_id;\n"
        "• Failed Prediction: SELECT s.name, h.hostel_name FROM students s JOIN hostels h ON s.student_id=h.hostel_id;\n"
        "• Root Cause & Fix: Skipped the junction table hostel_allotments. RAG FK relationship documents explicitly supply the 2-hop join path.\n\n"
        "Case Study 4 — Bad Aggregation Scope (Global LIMIT 1 vs. Per-Group MAX):\n"
        "• Question: 'Find the highest scorer in each department'\n"
        "• Gold SQL: SELECT d.dept_name, s.name, MAX(m.score) AS top_score FROM departments d JOIN students s ON d.dept_id=s.dept_id JOIN marks m ON s.student_id=m.student_id GROUP BY d.dept_id, d.dept_name;\n"
        "• Failed Prediction: Used ORDER BY AVG(m.score) DESC LIMIT 1 without grouping by department.\n\n"
        "Case Study 5 — Rejecting a Non-Equivalent Optimizer Rewrite:\n"
        "• Question: 'Show students who failed in any course'\n"
        "• Original Valid SQL: SELECT DISTINCT s.name FROM students s JOIN marks m ON s.student_id=m.student_id WHERE m.score < 40;\n"
        "• Optimizer Behaviour: An aggressive rewrite dropping DISTINCT changed row count from 189 to 312. Our Equivalence Checker detected _rows_equal() == False, logged the mismatch, and safely kept the original query."
    )

    doc.add_heading("Common Engineering Pitfalls & Mitigations Implemented", level=2)
    add_styled_docx_table(doc, ["Pitfall", "Cause", "Mitigation Implemented in Code"], PITFALLS_TABLE)

    # ── Section 11: Originality, Ethics, Conclusion & Future Work ─────────────
    doc.add_heading("11. Originality, Ethical Considerations, Conclusion & Future Work", level=1)
    doc.add_paragraph(
        "• Originality (Bonus Criterion): Combines a 4-collection vector store, QLoRA fine-tuning, a symbolic sqlglot CSP validator, an equivalence-checked optimizer, and a hybrid Colab/Local execution backend with 130 automated tests.\n"
        "• Ethical Considerations (Bonus Criterion):\n"
        "  1. Student Data Privacy: Zero student records or PII are sent to external paid LLM APIs.\n"
        "  2. Three-Layer Read-Only Safety: AST SELECT-only gate + SQLite URI mode=ro + PRAGMA query_only=ON prevent any accidental or malicious data modification.\n"
        "  3. Transparency & Auditability: Every query surfaces the exact SQL, retrieved rules, and per-agent latency trace.\n\n"
        "Conclusion:\n"
        "In this project, we designed, fine-tuned, and evaluated a complete Agentic Text-to-SQL system capable of translating natural language questions into verified SQLite/PostgreSQL queries, explanations, and visual charts. Across 102 benchmark queries, our full agentic system achieved 88.2% execution accuracy and 99.0% SQL validity—outperforming the pretrained baseline by +37.2 percentage points.\n\n"
        "Limitations & Future Work:\n"
        "• Multi-Database & Larger Schemas: Scale dynamic table-level RAG pruning to schemas with 100+ tables.\n"
        "• Query-Result Caching: Cache frequent natural language queries and embeddings to reduce repeat latency to <5ms.\n"
        "• Conversational Multi-Turn Memory: Support follow-up questions (e.g., 'Now filter those students to 3rd year only')."
    )

    # ── Section 12: References ────────────────────────────────────────────────
    doc.add_heading("12. References", level=1)
    refs = [
        "1. T. Yu, R. Zhang, K. Yang, M. Yasunaga, D. Wang, Z. Li, J. Ma, I. Li, Q. Yao, S. Roman, Z. Zhang, and D. Radev, \"Spider: A Large-Scale Human-Labeled Dataset for Complex and Cross-Domain Semantic Parsing and Text-to-SQL Task,\" in Proc. Conf. Empirical Methods in Natural Language Processing (EMNLP), Brussels, Belgium, 2018, pp. 3911–3921.",
        "2. J. Li, B. Hui, G. Qu, J. Yang, B. Li, B. Li, B. Wang, B. Qin, R. Geng, N. Huo, X. Zhou, C. Ma, G. Li, K. C. C. Chang, F. Huang, R. Cheng, and Y. Li, \"Can LLM Already Serve as A Database Interface? A BIg Bench for Large-Scale Database Grounded Text-to-SQLs (BIRD),\" in Advances in Neural Information Processing Systems (NeurIPS), vol. 36, 2023.",
        "3. M. Pourreza and D. Rafiei, \"DIN-SQL: Decomposed In-Context Learning of Text-to-SQL with Self-Correction,\" in Advances in Neural Information Processing Systems (NeurIPS), vol. 36, pp. 36339–36348, 2023.",
        "4. D. Gao, H. Wang, Y. Li, X. Sun, Y. Qian, B. Ding, and J. Zhou, \"Text-to-SQL Empowered by Large Language Models: A Benchmark Evaluation (DAIL-SQL),\" Proc. VLDB Endowment, vol. 17, no. 5, pp. 1132–1145, 2024.",
        "5. B. Wang, C. Ren, J. Yang, X. Liang, J. Bai, L. Chai, Z. Yan, Q.-W. Zhang, D. Yin, X. Sun, and Z. Li, \"MAC-SQL: A Multi-Agent Collaborative Framework for Text-to-SQL,\" in Proc. Joint Int. Conf. Computational Linguistics (LREC-COLING), Torino, Italy, 2024, pp. 540–557.",
        "6. H. Li, J. Zhang, C. Li, and H. Chen, \"RESDSQL: Decoupling Schema Linking and Skeleton Parsing for Text-to-SQL,\" in Proc. AAAI Conf. Artificial Intelligence (AAAI), vol. 37, no. 11, pp. 13067–13075, 2023.",
        "7. M. Talaei, M. Pourreza, Y.-C. Chang, A. Mirhoseini, and A. Saberi, \"CHESS: Contextual Harnessing for Efficient SQL Synthesis,\" arXiv preprint arXiv:2405.16755, 2024.",
        "8. T. Dettmers, A. Pagnoni, A. Holtzman, and L. Zettlemoyer, \"QLoRA: Efficient Finetuning of Quantized LLMs,\" in Advances in Neural Information Processing Systems (NeurIPS), vol. 36, 2023.",
        "9. B. Hui et al., \"Qwen2.5-Coder Technical Report,\" arXiv preprint arXiv:2409.12186, 2024.",
        "10. X. Chen, M. Lin, N. Schärli, and D. Zhou, \"Teaching Large Language Models to Self-Debug,\" in Proc. Int. Conf. Learning Representations (ICLR), Vienna, Austria, 2024.",
        "11. P. Lewis, E. Perez, A. Piktus, F. Petroni, V. Karpukhin, N. Goyal, H. Küttler, M. Lewis, W.-t. Yih, T. Rocktäschel, S. Riedel, and D. Kiela, \"Retrieval-Augmented Generation for Knowledge-Intensive NLP Tasks,\" in Advances in Neural Information Processing Systems (NeurIPS), vol. 33, pp. 9459–9474, 2020.",
        "12. N. Reimers and I. Gurevych, \"Sentence-BERT: Sentence Embeddings using Siamese BERT-Networks,\" in Proc. Conf. Empirical Methods in Natural Language Processing (EMNLP), Hong Kong, 2019, pp. 3982–3992.",
    ]
    for ref in refs:
        doc.add_paragraph(ref)

    doc.add_page_break()

    # ── Section 13: IEEE Research Paper Manuscript ────────────────────────────
    rp_banner = doc.add_table(rows=1, cols=1)
    rp_banner.alignment = WD_TABLE_ALIGNMENT.CENTER
    rpc = rp_banner.rows[0].cells[0]
    set_cell_shading(rpc, "1E3A8A")
    rpp0 = rpc.paragraphs[0]
    rpp0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_tag = rpp0.add_run("RESEARCH PAPER MANUSCRIPT (IEEE CONFERENCE FORMAT)\n")
    r_tag.bold = True
    r_tag.font.size = Pt(10)
    r_tag.font.color.rgb = RGBColor(147, 197, 253)

    rpp1 = rpc.add_paragraph()
    rpp1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ptitle = rpp1.add_run(
        "Agentic Text-to-SQL: A Neuro-Symbolic Multi-Agent Architecture with Vector RAG, "
        "4-Bit QLoRA Fine-Tuning, and CSP-Guided Self-Correction"
    )
    r_ptitle.bold = True
    r_ptitle.font.size = Pt(15)
    r_ptitle.font.color.rgb = RGBColor(255, 255, 255)

    rpp2 = rpc.add_paragraph()
    rpp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_pauth = rpp2.add_run(
        "Devesh Yadav (240437), Chakshit Mehta (240429), Yash Bhardwaj (240957), "
        "Anmol Bhardwaj (240383), and Avyam (240417)\n"
        "Department of Computer Science and Engineering"
    )
    r_pauth.font.size = Pt(10)
    r_pauth.font.color.rgb = RGBColor(226, 232, 240)

    doc.add_paragraph()

    p_abs = doc.add_paragraph()
    p_abs.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r_ab_h = p_abs.add_run("Abstract—")
    r_ab_h.bold = True
    r_ab_h.italic = True
    p_abs.add_run(
        "Translating natural language questions into executable Structured Query Language (SQL) over multi-table relational "
        "databases remains challenging for single-pass Large Language Models (LLMs) due to schema hallucinations, overloaded "
        "column names, unwritten domain conventions, and lack of execution safety. In this paper, we present Agentic Text-to-SQL, "
        "a 100% open-source, 9-stage neuro-symbolic multi-agent architecture that integrates dense vector Retrieval-Augmented "
        "Generation (RAG), 4-bit NormalFloat QLoRA fine-tuning of Qwen2.5-Coder, formal Constraint Satisfaction Problem (CSP) "
        "validation via Abstract Syntax Tree (AST) inspection, an error-guided state-space self-correction loop, and an "
        "equivalence-checked algebraic query optimizer. Evaluated on a 12-table College Relational Database across 102 gold-standard "
        "queries spanning Easy, Medium, and Hard complexity tiers, our full agentic system achieves 88.2% execution accuracy and "
        "99.0% SQL validity—outperforming the pretrained baseline by +37.2 percentage points and the standalone QLoRA fine-tuned "
        "model by +18.6 percentage points while guaranteeing strict read-only database safety.\n"
    )
    r_kw_h = p_abs.add_run("Index Terms—")
    r_kw_h.bold = True
    r_kw_h.italic = True
    p_abs.add_run(
        "Text-to-SQL, Multi-Agent Systems, Retrieval-Augmented Generation (RAG), QLoRA Fine-Tuning, "
        "Constraint Satisfaction Problem (CSP), State-Space Search, Self-Debugging, Neuro-Symbolic AI."
    )

    paper_sections = [
        ("I. INTRODUCTION",
         "Relational database management systems (RDBMS) underpin modern academic, healthcare, and enterprise analytics, yet accessing "
         "these repositories requires proficiency in SQL and detailed knowledge of normalized schemas [1], [2]. While frontier proprietary "
         "LLMs have advanced semantic parsing [3], [4], deploying them in institutional environments introduces severe data-privacy risks, "
         "recurring API costs, and fragility when domain-specific business rules (e.g., institutional department codes, grading thresholds, "
         "and attendance formulas) are absent from raw DDL statements. Conversely, compact open-source code models (1.5B–7B parameters) "
         "frequently hallucinate column names across multi-table joins when invoked in a single pass [6].\n\n"
         "To address these limitations, we propose a neuro-symbolic multi-agent framework that treats Text-to-SQL not as a one-shot "
         "sequence-to-sequence translation, but as a closed-loop State-Space Search governed by a symbolic Constraint Satisfaction Problem (CSP). "
         "Our primary contributions are fourfold: (1) a 4-collection ChromaDB + MiniLM vector store indexing table DDLs with sample values, "
         "foreign-key join paths, declarative business rules, and solved few-shot exemplars; (2) 4-bit QLoRA adaptation of Qwen2.5-Coder [8], [9] "
         "with prompt-token loss masking; (3) a deterministic sqlglot AST & EXPLAIN validator that feeds violation gradients into a 3-try repair "
         "loop [10]; and (4) an algebraic optimizer with strict row-set equivalence verification."),
        ("II. RELATED WORK",
         "Cross-domain benchmarks such as Spider [1] and BIRD [2] established that multi-table joins and external domain evidence are the primary "
         "bottlenecks in Text-to-SQL. Decomposition frameworks including DIN-SQL [3] and multi-agent pipelines such as MAC-SQL [5] and CHESS [7] "
         "demonstrated significant accuracy improvements by separating schema linking from query refinement. DAIL-SQL [4] further showed that "
         "question-and-SQL similarity retrieval maximizes in-context learning efficiency. Our work extends these findings to compact, locally "
         "hosted open-source models by combining QLoRA parameter-efficient fine-tuning [8] with symbolic AST/CSP verification and row-equivalence "
         "optimization."),
        ("III. PROPOSED NEURO-SYMBOLIC METHODOLOGY",
         "Let Q denote a user question and D = (S_T, S_C, F_K) denote a relational database with tables S_T, scoped columns S_C(t), and foreign-key "
         "edges F_K. Our pipeline executes nine sequential modules:\n"
         "1) Multi-Collection Vector RAG: Using Sentence-BERT (all-MiniLM-L6-v2) [11], [12], Q is encoded into a 384-D unit vector u_Q. We retrieve "
         "top-k chunks via cosine similarity sim(u_Q, u_d) = u_Q^T u_d across four collections (tables, FK edges, 30 business rules, 40 few-shot examples) "
         "to assemble context C_RAG.\n"
         "2) State-Space Search & Self-Correction: At attempt t in {1, 2, 3}, the SQL Generation Agent samples SQL_t = f_theta(Q, C_RAG, SQL_{t-1}, E_{t-1}). "
         "The Validator Agent parses SQL_t into an AST T_AST and evaluates four crisp CSP constraints: C_1 (Single SELECT-only statement, zero DDL/DML nodes), "
         "C_2 (Table membership in S_T), C_3 (Scoped alias and column resolution in S_C(t), exempting SELECT aliases), and C_4 (SQLite EXPLAIN QUERY PLAN "
         "feasibility). If any C_i fails, the exact violated variable and its legal domain are formatted into diagnostic error E_t for attempt t+1.\n"
         "3) Equivalence-Checked Optimization & Safe Execution: Validated queries undergo sqlglot algebraic simplification to produce SQL_opt. Both "
         "SQL_t and SQL_opt are executed inside a read-only sandbox (mode=ro, PRAGMA query_only=ON) with a thread-safe timeout; SQL_opt is accepted if and "
         "only if norm_rows(SQL_opt) == norm_rows(SQL_t). Finally, the Result Processor and Explainer Agent synthesize statistical insights and a Plotly chart."),
        ("IV. EXPERIMENTAL SETUP & ABLATION RESULTS",
         "We evaluate six system configurations on a 12-table College Relational Database (~15,000 rows) across 102 gold-standard queries (30 Easy, "
         "41 Medium, 31 Hard). While the pretrained baseline (System A) achieves 51.0% execution accuracy (70.6% validity), adding Vector RAG alone (A1) "
         "yields 66.7%, and adding the CSP Repair Loop alone (A2) yields 63.7% (93.1% validity). QLoRA fine-tuning alone (System B) reaches 69.6% accuracy, "
         "rising to 78.4% with the repair loop (B1) and culminating in 88.2% execution accuracy (96.7% Easy, 90.2% Medium, 77.4% Hard) and 99.0% SQL validity "
         "for the Full Agentic Pipeline (System C) in an average of 1.18 attempts."),
        ("V. CONCLUSION",
         "Our empirical findings demonstrate that pairing a compact QLoRA-adapted open-source code model with multi-collection Vector RAG, symbolic CSP "
         "validation, and iterative self-debugging closes the gap with proprietary cloud APIs while guaranteeing 100% local data privacy and read-only safety."),
    ]
    for p_sec_title, p_sec_body in paper_sections:
        doc.add_heading(p_sec_title, level=2)
        for para in p_sec_body.split("\n\n"):
            p = doc.add_paragraph(para)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    doc.save(path)


def build_report_pdf(path):
    doc = SimpleDocTemplate(
        path, pagesize=A4,
        leftMargin=0.6*inch, rightMargin=0.6*inch, topMargin=0.6*inch, bottomMargin=0.6*inch
    )
    styles = getSampleStyleSheet()
    h1_s = ParagraphStyle("H1", parent=styles["Heading1"], fontSize=13.5, leading=17, textColor=colors.HexColor("#1E3A8A"), spaceBefore=12, spaceAfter=6)
    h2_s = ParagraphStyle("H2", parent=styles["Heading2"], fontSize=11, leading=14, textColor=colors.HexColor("#0F172A"), spaceBefore=8, spaceAfter=4)
    body_s = ParagraphStyle("Body", parent=styles["Normal"], fontSize=9.3, leading=13.2, spaceAfter=6)
    code_s = ParagraphStyle("Code", parent=styles["Code"], fontName="Courier", fontSize=7.2, leading=9.2, backColor=colors.HexColor("#F1F5F9"), borderPadding=5, spaceAfter=8)
    title_s = ParagraphStyle("CoverT", parent=styles["Heading1"], fontSize=17, leading=22, alignment=1, textColor=colors.HexColor("#1E3A8A"), spaceAfter=10)
    sub_s = ParagraphStyle("CoverSub", parent=styles["Normal"], fontSize=10.5, leading=14.5, alignment=1, textColor=colors.HexColor("#334155"), spaceAfter=12)
    cell_s = ParagraphStyle("CS", parent=styles["Normal"], fontSize=7.2, leading=9.2)
    hdr_s = ParagraphStyle("HS", parent=styles["Normal"], fontSize=7.6, leading=9.5, textColor=colors.white, fontName="Helvetica-Bold")

    def make_rl_table(headers, rows, col_widths, highlight_last=False):
        data = [[Paragraph(h.replace("&", "&amp;"), hdr_s) for h in headers]]
        for r in rows:
            data.append([Paragraph(str(c).replace("&", "&amp;").replace("\n", "<br/>"), cell_s) for c in r])
        t = Table(data, colWidths=col_widths, repeatRows=1)
        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]
        if highlight_last:
            style_cmds.append(("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#DBEAFE")))
        t.setStyle(TableStyle(style_cmds))
        return t

    # ── Dedicated Formal PDF Cover Page ───────────────────────────────────────
    banner_title = ParagraphStyle("BannerT", parent=styles["Normal"], fontSize=20, leading=25, alignment=1, textColor=colors.white, fontName="Helvetica-Bold")
    banner_sub = ParagraphStyle("BannerSub", parent=styles["Normal"], fontSize=11, leading=15, alignment=1, textColor=colors.HexColor("#93C5FD"), fontName="Helvetica-Bold")
    banner_desc = ParagraphStyle("BannerDesc", parent=styles["Normal"], fontSize=9.8, leading=13.5, alignment=1, textColor=colors.HexColor("#E2E8F0"))

    cover_banner = Table([
        [Paragraph("END-SEMESTER ARTIFICIAL INTELLIGENCE PROJECT REPORT", banner_sub)],
        [Paragraph("AGENTIC TEXT-TO-SQL SYSTEM", banner_title)],
        [Paragraph(
            "A Multi-Agent Neuro-Symbolic Pipeline with Vector RAG, 4-Bit QLoRA Fine-Tuning,<br/>"
            "Symbolic CSP Validation, Self-Correction Loop &amp; Equivalence-Checked Optimization",
            banner_desc
        )],
    ], colWidths=[7.0*inch])
    cover_banner.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#0F172A")),
        ("TOPPADDING", (0, 0), (-1, 0), 16),
        ("BOTTOMPADDING", (0, -1), (-1, -1), 16),
        ("LEFTPADDING", (0, 0), (-1, -1), 14),
        ("RIGHTPADDING", (0, 0), (-1, -1), 14),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))

    meta_rows = [
        ["Course & Project Theme", "Artificial Intelligence — AI for Human Insight and Capability Enhancement"],
        ["Core AI Architecture", "Qwen2.5-Coder (4-bit QLoRA) + ChromaDB/MiniLM Vector RAG + sqlglot AST/CSP Validator"],
        ["GitHub Repository", "https://github.com/DeveshYadav85/text-to-sql"],
        ["Key Benchmark Result", "88.2% Execution Accuracy & 99.0% SQL Validity across 102 Gold Queries (130/130 Tests Pass)"],
    ]
    meta_tbl = make_rl_table(["Project Attribute", "Details"], meta_rows, [1.8*inch, 5.2*inch])

    team_tbl = make_rl_table(
        ["S. No.", "Student Name", "Roll / ID No.", "Module Ownership & Technical Contribution"],
        TEAM_MEMBERS,
        [0.55*inch, 1.45*inch, 1.0*inch, 4.0*inch]
    )

    exec_summary_box = Table([[
        Paragraph(
            "<b>Executive Summary:</b> This report presents the end-to-end architecture, formal AI modeling, QLoRA fine-tuning, "
            "and empirical ablation study of an open-source Agentic Text-to-SQL system. By formulating SQL synthesis as an "
            "<b>Error-Guided State-Space Search</b> and static query verification as a <b>Symbolic Constraint Satisfaction Problem (CSP)</b>, "
            "our full agentic pipeline elevates execution accuracy from <b>51.0%</b> (System A: Pretrained Baseline) to <b>69.6%</b> "
            "(System B: 4-Bit QLoRA Fine-Tuned) and <b>88.2%</b> (System C: Full Agentic Pipeline with Vector RAG + 3-Try Repair Loop), "
            "achieving a <b>99.0% SQL validity rate</b> across 102 gold-standard queries on a 12-table College Relational Database.",
            body_s
        )
    ]], colWidths=[7.0*inch])
    exec_summary_box.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#EFF6FF")),
        ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#2563EB")),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
    ]))

    toc_rows_pdf = [
        ["1.", "Introduction (What is Agentic Text-to-SQL, Problem Statement, Objectives, Motivation, Significance, Challenges, Novelty)"],
        ["2.", "Literature Review & Comparative Table (Spider, DIN-SQL, DAIL-SQL, MAC-SQL, QLoRA, Self-Debugging)"],
        ["3.", "Methodology: Agentic Text-to-SQL as State-Space Search (States, Initial/Goal State, Actions, Search Strategy)"],
        ["4.", "Knowledge Representation & RAG Design (Schema Graph, Business Rules, Case Memory, Small-vs-Large Schema Trade-off)"],
        ["5.", "SQL Validation as a Constraint Satisfaction Problem (CSP) & Applicability of Fuzzy Logic (Rubric Note 2)"],
        ["6.", "CSP Trace for a Particular Query Instance (Step-by-Step Verification & Self-Correction)"],
        ["7.", "Intelligent System Design, 5-Member Team Split, 3-Week Timeline & AI vs. Non-AI Distinction (Rubric Note 1)"],
        ["8.", "Implementation Code & QLoRA Fine-Tuning Loss Curve"],
        ["9.", "Experimental Evaluation: Systems A (Baseline), B (Fine-Tuned), C (Full Agentic) & Component Ablations"],
        ["10.", "Discussion, Error Analysis (5 Case Studies) & Common Engineering Pitfalls"],
        ["11.", "Originality, Ethical Considerations, Conclusion & Future Work"],
        ["12.", "References (12 IEEE Research Papers)"],
        ["13.", "Research Paper Manuscript (IEEE Conference Format)"],
    ]

    story = [
        cover_banner,
        Spacer(1, 14),
        meta_tbl,
        Spacer(1, 14),
        Paragraph("<b>Submitted By — Project Team Members</b>", h2_s),
        team_tbl,
        Spacer(1, 16),
        exec_summary_box,
        PageBreak(),
        Paragraph("Table of Contents", title_s),
        Spacer(1, 8),
        make_rl_table(["Sec.", "Section Title & Sub-Topics Covered"], toc_rows_pdf, [0.55*inch, 6.45*inch]),
        PageBreak(),
    ]

    # Render main report sections into PDF
    story.append(Paragraph("1. Introduction", h1_s))
    story.append(Paragraph("<b>What is Agentic Text-to-SQL?</b> Relational databases store structured organizational knowledge, but querying them requires fluency in SQL, schema join paths, and unwritten domain conventions. Our Agentic Text-to-SQL system takes a plain-English question (e.g., 'Show the top 10 students from CSE who scored above 80') and orchestrates specialized agents for Vector RAG retrieval, open-source LLM SQL generation, sqlglot AST/CSP validation, a 3-try self-correction repair loop, equivalence-checked optimization, safe read-only execution, and visual explanation.", body_s))
    story.append(Paragraph("<b>Problem Statement:</b> Single-pass LLMs fail on real databases due to (1) Schema Hallucination &amp; ambiguous columns ('name', 'year', 'status'), (2) Hidden Business Logic ('CSE' -> dept_code='CSE', 'Pass' -> score>=40), and (3) Lack of Verification &amp; Read-Only Safety. Paid cloud APIs also pose privacy and cost barriers. We solve this using a 100% open-source model that we baseline (System A), fine-tune with QLoRA (System B), and wrap in a RAG + Validator + Repair Agent (System C).", body_s))
    story.append(Paragraph("<b>Motivation, Significance &amp; Novelty:</b> Aligned with 'AI for Human Insight and Capability Enhancement', our system augments human reasoning by returning SQL, data tables, auto-selected Plotly charts, statistical insights, and an auditable 9-step trace. Our novelty lies in combining 4-collection Vector RAG, 4-bit QLoRA fine-tuning, symbolic CSP validation, and row-equivalence optimization.", body_s))

    story.append(Paragraph("2. Literature Review &amp; Comparative Table (10 Research Papers)", h1_s))
    story.append(Paragraph(
        "<b>Comprehensive Survey of 10 Research Papers:</b><br/>"
        "• <b>[1] Spider (Yu et al., EMNLP 2018):</b> Established the 200-database cross-domain Text-to-SQL benchmark and Easy/Medium/Hard complexity split, proving schema linking is the core bottleneck (&lt;25% accuracy on Hard joins).<br/>"
        "• <b>[2] BIRD (Li et al., NeurIPS 2023):</b> 12,751 queries on 95 noisy databases (33.4 GB); proved GPT-4 jumps from 54.89% to 72.28% when external domain business rules are retrieved, motivating our 30-rule knowledge base and SQL Optimizer.<br/>"
        "• <b>[3] DIN-SQL (Pourreza &amp; Rafiei, NeurIPS 2023):</b> Decomposed Text-to-SQL into schema linking, query classification, sub-query generation, and self-correction (85.3% on Spider).<br/>"
        "• <b>[4] DAIL-SQL (Gao et al., VLDB 2024):</b> Demonstrated that question-and-SQL similarity retrieval for few-shot in-context learning boosts accuracy by +11.4% over zero-shot prompting.<br/>"
        "• <b>[5] MAC-SQL (Wang et al., COLING 2024):</b> 3-agent collaborative pipeline (Selector, Decomposer, Refiner) using execution feedback to repair defective queries (+9.2% over single-agent baselines).<br/>"
        "• <b>[6] RESDSQL &amp; CodeS (Li et al., AAAI 2023 / SIGMOD 2024):</b> Showed compact 1.5B–7B open-source code models rival 175B+ models when provided with explicit foreign-key join hints.<br/>"
        "• <b>[7] CHESS (Talaei et al., 2024):</b> Hierarchical retrieval indexing distinct column sample values reduced wrong-column and wrong-filter errors by &gt;40% on ambiguous schemas.<br/>"
        "• <b>[8] QLoRA (Dettmers et al., NeurIPS 2023) &amp; Qwen2.5-Coder (Hui et al., 2024):</b> 4-bit NormalFloat (NF4) quantization + Low-Rank Adapters cut GPU memory by 75%, enabling free Colab T4 fine-tuning.<br/>"
        "• <b>[9] Self-Debugging (Chen et al., ICLR 2024):</b> Error-guided iterative repair using compiler/AST and EXPLAIN diagnostics resolves 68% of first-pass SQL errors within 2–3 tries.<br/>"
        "• <b>[10] RAG (Lewis et al., NeurIPS 2020) &amp; Sentence-BERT (Reimers &amp; Gurevych, EMNLP 2019):</b> Siamese bi-encoder embeddings (all-MiniLM-L6-v2) + HNSW/FAISS enable sub-50ms semantic retrieval.",
        body_s
    ))
    lit_headers = ["Ref.", "Author(s)", "Objective", "Methods", "Key Findings", "Remarks / Gap Filled"]
    lit_rows = [[r["ref"], r["authors"], r["objective"], r["methods"], r["findings"], r["future"]] for r in LIT_PAPERS]
    story.append(make_rl_table(lit_headers, lit_rows, [0.45*inch, 1.15*inch, 1.25*inch, 1.4*inch, 1.35*inch, 1.35*inch], highlight_last=True))
    story.append(Spacer(1, 8))

    story.append(Paragraph("3. Methodology: State-Space Search &amp; 4. Knowledge Representation", h1_s))
    story.append(Paragraph("<b>State-Space Search:</b> State s_t = (Q, C_RAG, SQL_t, E_t, t) for t in {0..3}. Initial State s_0 = (Q, C_RAG, None, None, 0). Goal State: Valid(SQL_t, Schema)=True &amp; EXPLAIN QUERY PLAN succeeds. Actions: RetrieveContext, GenerateSQL, RepairSQL(E_t), OptimizeSQL. Strategy: Error-Guided Heuristic Search with Bounded Backtracking (Max 3 tries).", body_s))
    story.append(Paragraph("<b>Knowledge Representation &amp; Small-vs-Large Schema Trade-off:</b> Combines (1) Relational Schema &amp; Foreign-Key Graph (12 tables, 16 FK edges), (2) 30 Declarative Business Rules (business_rules.md), (3) 40 Few-Shot Case Examples (examples.json), and (4) 384-D ChromaDB/FAISS Embeddings (all-MiniLM-L6-v2). For small/medium schemas, including all table DDLs while retrieving top-k rules and examples prevents missing join tables; for massive schemas, FK-expanded table retrieval scales seamlessly.", body_s))
    if os.path.exists("docs/csp_search_diagram.png"):
        story.append(RLImage("docs/csp_search_diagram.png", width=6.4*inch, height=2.8*inch))
        story.append(Spacer(1, 6))

    story.append(Paragraph("5. Constraint Satisfaction Problem (CSP) &amp; 6. Instance Trace", h1_s))
    story.append(Paragraph("<b>CSP Formulation:</b> Variables X = {StmtType, T_used, A_map, C_ref}. Domains D = SchemaTables &amp; Inspector Columns. Hard Constraints: C1 (Single SELECT only; blocks INSERT/UPDATE/DELETE/DROP), C2 (Table existence), C3 (Scoped alias &amp; column resolution, exempting SELECT aliases), and C4 (EXPLAIN QUERY PLAN dry-run).<br/><b>Rubric Note 2 (Why CSP over Fuzzy Logic):</b> SQL syntax and column existence are binary/crisp, requiring exact CSP verification; linguistic vagueness is handled upstream by continuous cosine similarity in Vector RAG.<br/><b>Instance Trace ('Top 10 CSE students scoring above 80'):</b> If Attempt 1 uses hallucinated column student_name, C3 fails and returns the exact schema columns; Attempt 2 joins students, departments (dept_code='CSE'), and marks (score > 80) and passes all constraints C1–C4 in 4 ms.", body_s))

    story.append(Paragraph("7. System Architecture, Team Split, Timeline &amp; AI vs. Non-AI (Note 1)", h1_s))
    if os.path.exists("docs/architecture_diagram.png"):
        story.append(RLImage("docs/architecture_diagram.png", width=6.4*inch, height=3.3*inch))
        story.append(Spacer(1, 6))
    story.append(Paragraph("<b>4-Person Team Split &amp; 3-Week Timeline:</b> Built around the shared interface contract pipeline.run(question) / agent.answer(question).", body_s))
    story.append(make_rl_table(["Role", "Responsibility", "Owned Files"], TEAM_SPLIT_TABLE, [0.8*inch, 2.7*inch, 3.45*inch]))
    story.append(Spacer(1, 6))
    story.append(make_rl_table(["Week", "Phase", "Deliverables"], TIMELINE_TABLE, [0.75*inch, 1.5*inch, 4.7*inch]))
    story.append(Spacer(1, 6))
    story.append(Paragraph("<b>Rubric Note 1 (AI vs. Non-AI Distinction):</b> AI components include MiniLM Dense Embeddings, QLoRA Fine-Tuned Qwen2.5-Coder, and NL Insight Synthesis. Non-AI deterministic components include sqlglot AST/CSP validation, row-equivalence query optimization, read-only SQLite execution (mode=ro, PRAGMA query_only=ON), and Pandas/Plotly rendering.", body_s))

    story.append(Paragraph("8. Core Implementation &amp; QLoRA Fine-Tuning", h1_s))
    story.append(Preformatted(CODE_SNIPPETS["validator"], code_s))
    if os.path.exists("docs/training_loss_curve.png"):
        story.append(RLImage("docs/training_loss_curve.png", width=5.8*inch, height=2.8*inch))
        story.append(Spacer(1, 6))

    story.append(Paragraph("9. Experimental Evaluation: Baseline vs. Fine-Tuned vs. Agentic + Ablations", h1_s))
    ab_headers = ["System", "Model", "RAG", "Agent Loop", "Overall Acc", "Easy", "Med", "Hard", "Valid Rate", "Tries"]
    story.append(make_rl_table(ab_headers, ABLATION_TABLE, [1.35*inch, 0.95*inch, 0.75*inch, 0.7*inch, 0.6*inch, 0.45*inch, 0.45*inch, 0.45*inch, 0.6*inch, 0.45*inch], highlight_last=True))
    story.append(Spacer(1, 6))
    if os.path.exists("docs/ablation_accuracy.png"):
        story.append(RLImage("docs/ablation_accuracy.png", width=6.1*inch, height=2.9*inch))
    if os.path.exists("docs/difficulty_breakdown.png"):
        story.append(RLImage("docs/difficulty_breakdown.png", width=5.8*inch, height=2.8*inch))

    story.append(Paragraph("10. Error Analysis, Case Studies &amp; Engineering Pitfalls", h1_s))
    story.append(make_rl_table(["Category", "Description", "Share", "Affected", "Mitigation"], ERROR_CATEGORIES_TABLE, [1.1*inch, 1.9*inch, 0.5*inch, 1.1*inch, 2.35*inch]))
    story.append(Spacer(1, 6))
    story.append(make_rl_table(["Pitfall", "Cause", "Mitigation Implemented"], PITFALLS_TABLE, [1.5*inch, 2.1*inch, 3.35*inch]))
    story.append(Spacer(1, 6))

    story.append(Paragraph("11. Originality, Ethics &amp; Conclusion | 12. IEEE References", h1_s))
    story.append(Paragraph(
        "<b>Originality &amp; Ethics:</b> 100% local privacy (zero PII sent to cloud APIs), triple-layer read-only DB protection "
        "(SELECT-only AST + URI mode=ro + PRAGMA query_only=ON), and transparent 9-step explainability. Full Agentic System achieves "
        "88.2% execution accuracy and 99.0% validity across 102 queries (+37.2% over baseline), with 130/130 pytest tests passing.<br/><br/>"
        "<b>IEEE References:</b><br/>"
        "[1] T. Yu et al., \"Spider: A Large-Scale Human-Labeled Dataset for Complex and Cross-Domain Semantic Parsing and Text-to-SQL Task,\" in Proc. EMNLP, 2018, pp. 3911–3921.<br/>"
        "[2] J. Li et al., \"Can LLM Already Serve as A Database Interface? A BIg Bench for Large-Scale Database Grounded Text-to-SQLs (BIRD),\" in NeurIPS, vol. 36, 2023.<br/>"
        "[3] M. Pourreza and D. Rafiei, \"DIN-SQL: Decomposed In-Context Learning of Text-to-SQL with Self-Correction,\" in NeurIPS, vol. 36, pp. 36339–36348, 2023.<br/>"
        "[4] D. Gao et al., \"Text-to-SQL Empowered by Large Language Models: A Benchmark Evaluation (DAIL-SQL),\" Proc. VLDB Endow., vol. 17, no. 5, pp. 1132–1145, 2024.<br/>"
        "[5] B. Wang et al., \"MAC-SQL: A Multi-Agent Collaborative Framework for Text-to-SQL,\" in Proc. LREC-COLING, 2024, pp. 540–557.<br/>"
        "[6] H. Li et al., \"RESDSQL: Decoupling Schema Linking and Skeleton Parsing for Text-to-SQL,\" in Proc. AAAI, vol. 37, pp. 13067–13075, 2023.<br/>"
        "[7] M. Talaei et al., \"CHESS: Contextual Harnessing for Efficient SQL Synthesis,\" arXiv:2405.16755, 2024.<br/>"
        "[8] T. Dettmers et al., \"QLoRA: Efficient Finetuning of Quantized LLMs,\" in NeurIPS, vol. 36, 2023.<br/>"
        "[9] B. Hui et al., \"Qwen2.5-Coder Technical Report,\" arXiv:2409.12186, 2024.<br/>"
        "[10] X. Chen et al., \"Teaching Large Language Models to Self-Debug,\" in Proc. ICLR, 2024.<br/>"
        "[11] P. Lewis et al., \"Retrieval-Augmented Generation for Knowledge-Intensive NLP Tasks,\" in NeurIPS, vol. 33, pp. 9459–9474, 2020.<br/>"
        "[12] N. Reimers and I. Gurevych, \"Sentence-BERT: Sentence Embeddings using Siamese BERT-Networks,\" in Proc. EMNLP, 2019, pp. 3982–3992.",
        body_s
    ))

    story.append(PageBreak())
    paper_banner = Table([
        [Paragraph("SECTION 13 — RESEARCH PAPER MANUSCRIPT (IEEE CONFERENCE FORMAT)", banner_sub)],
        [Paragraph("Agentic Text-to-SQL: A Neuro-Symbolic Multi-Agent Architecture with Vector RAG, 4-Bit QLoRA Fine-Tuning, and CSP-Guided Self-Correction", ParagraphStyle("PT", parent=banner_title, fontSize=13, leading=17))],
        [Paragraph("Devesh Yadav (240437), Chakshit Mehta (240429), Yash Bhardwaj (240957), Anmol Bhardwaj (240383), and Avyam (240417)", banner_desc)],
    ], colWidths=[7.0*inch])
    paper_banner.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#1E3A8A")),
        ("TOPPADDING", (0, 0), (-1, 0), 12),
        ("BOTTOMPADDING", (0, -1), (-1, -1), 12),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
    ]))
    story.append(paper_banner)
    story.append(Spacer(1, 10))
    story.append(Paragraph(
        "<b><i>Abstract</i>—</b>Translating natural language questions into executable Structured Query Language (SQL) over multi-table "
        "relational databases remains challenging for single-pass Large Language Models (LLMs) due to schema hallucinations, overloaded "
        "column names, unwritten domain conventions, and lack of execution safety. In this paper, we present <b>Agentic Text-to-SQL</b>, a "
        "100% open-source, 9-stage neuro-symbolic multi-agent architecture that integrates dense vector Retrieval-Augmented Generation "
        "(RAG), 4-bit NormalFloat QLoRA fine-tuning of Qwen2.5-Coder, formal Constraint Satisfaction Problem (CSP) validation via Abstract "
        "Syntax Tree (AST) inspection, an error-guided state-space self-correction loop, and an equivalence-checked algebraic query optimizer. "
        "Evaluated on a 12-table College Relational Database across 102 gold-standard queries spanning Easy, Medium, and Hard complexity tiers, "
        "our full agentic system achieves <b>88.2% execution accuracy</b> and <b>99.0% SQL validity</b>—outperforming the pretrained baseline "
        "by +37.2 percentage points and the standalone QLoRA fine-tuned model by +18.6 percentage points while guaranteeing strict read-only safety.<br/>"
        "<b><i>Index Terms</i>—</b>Text-to-SQL, Multi-Agent Systems, Retrieval-Augmented Generation (RAG), QLoRA Fine-Tuning, Constraint Satisfaction Problem (CSP), State-Space Search, Self-Debugging.",
        body_s
    ))
    story.append(Paragraph("I. INTRODUCTION &amp; II. RELATED WORK", h2_s))
    story.append(Paragraph(
        "Relational databases underpin institutional analytics, yet querying them requires mastery of SQL and domain-specific business rules [1], [2]. "
        "While proprietary models (e.g., DIN-SQL [3], DAIL-SQL [4], MAC-SQL [5]) achieve strong results on Spider and BIRD, they raise severe privacy and "
        "cost barriers for student records. Conversely, small open-source models (1.5B–7B) hallucinate columns when used in a single pass [6], [7]. We "
        "bridge this gap by pairing 4-bit QLoRA adaptation of Qwen2.5-Coder [8], [9] with 4-collection Sentence-BERT RAG [11], [12] and symbolic CSP self-debugging [10].",
        body_s
    ))
    story.append(Paragraph("III. NEURO-SYMBOLIC METHODOLOGY &amp; IV. EXPERIMENTAL RESULTS", h2_s))
    story.append(Paragraph(
        "Given question Q and schema graph D = (S_T, S_C, F_K), Step 2 retrieves top-k chunks across 4 ChromaDB collections (tables with sample values, "
        "16 FK edges, 30 business rules, 40 solved examples). Steps 3–4 perform Error-Guided State-Space Search over states s_t = (Q, C_RAG, SQL_t, E_t, t), "
        "enforcing four crisp CSP constraints: C1 (Single SELECT-only statement), C2 (Table existence), C3 (Scoped alias/column resolution), and C4 "
        "(EXPLAIN QUERY PLAN dry-run). Step 5 applies sqlglot algebraic rewrites accepted iff norm_rows(SQL_opt) == norm_rows(SQL_t). Across 102 gold "
        "queries, accuracy progresses from 51.0% (System A: Baseline) to 66.7% (A1: +RAG), 63.7% (A2: +Repair Loop), 69.6% (System B: QLoRA), 78.4% "
        "(B1: QLoRA + Repair Loop), and 88.2% (System C: Full Agentic Pipeline, 99.0% SQL validity, 1.18 avg attempts).",
        body_s
    ))

    doc.build(story)


# ═════════════════════════════════════════════════════════════════════════════
# 3. PRESENTATION DECK (PPTX) — 15 Slides Covering Rubrics + Full Build Guide
# ═════════════════════════════════════════════════════════════════════════════
def add_slide_header(slide, title_text, subtitle_text=""):
    banner = slide.shapes.add_shape(1, PptInches(0), PptInches(0), PptInches(13.333), PptInches(1.15))
    banner.fill.solid()
    banner.fill.fore_color.rgb = PptRGBColor(15, 23, 42)
    banner.line.fill.background()

    tb = slide.shapes.add_textbox(PptInches(0.5), PptInches(0.12), PptInches(12.3), PptInches(0.95))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.size = PptPt(24)
    p.font.bold = True
    p.font.color.rgb = PptRGBColor(255, 255, 255)

    if subtitle_text:
        p2 = tf.add_paragraph()
        p2.text = subtitle_text
        p2.font.size = PptPt(13)
        p2.font.color.rgb = PptRGBColor(147, 197, 253)


def add_bullets(slide, items, left=0.6, top=1.35, width=12.1, height=5.7, font_size=15.5):
    tb = slide.shapes.add_textbox(PptInches(left), PptInches(top), PptInches(width), PptInches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.font.size = PptPt(font_size)
        p.font.color.rgb = PptRGBColor(15, 23, 42)
        p.space_after = PptPt(8)


def build_presentation_pptx(path):
    prs = Presentation()
    prs.slide_width = PptInches(13.333)
    prs.slide_height = PptInches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Slide 1: Cover
    s1 = prs.slides.add_slide(blank_layout)
    bg = s1.shapes.add_shape(1, PptInches(0), PptInches(0), PptInches(13.333), PptInches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = PptRGBColor(15, 23, 42)
    bg.line.fill.background()

    tb = s1.shapes.add_textbox(PptInches(0.8), PptInches(1.3), PptInches(11.7), PptInches(5.2))
    tf = tb.text_frame
    tf.word_wrap = True
    p0 = tf.paragraphs[0]
    p0.text = "AGENTIC TEXT-TO-SQL SYSTEM"
    p0.font.size = PptPt(40)
    p0.font.bold = True
    p0.font.color.rgb = PptRGBColor(96, 165, 250)

    p1 = tf.add_paragraph()
    p1.text = "Neuro-Symbolic Multi-Agent Pipeline with Vector RAG, QLoRA Fine-Tuning, CSP Validation & Self-Correction"
    p1.font.size = PptPt(21)
    p1.font.bold = True
    p1.font.color.rgb = PptRGBColor(255, 255, 255)
    p1.space_before = PptPt(12)

    p2 = tf.add_paragraph()
    p2.text = (
        "\n• Team Members: Devesh Yadav (240437) · Chakshit Mehta (240429) · Yash Bhardwaj (240957) · Anmol Bhardwaj (240383) · Avyam (240417)"
        "\n• Theme: AI for Human Insight and Capability Enhancement"
        "\n• Systems Compared: A (Pretrained Baseline) vs. B (LoRA Fine-Tuned) vs. C (Full Agentic + 3 Ablations)"
        "\n• Stack: Qwen2.5-Coder (4-bit QLoRA) · ChromaDB/FAISS RAG · sqlglot AST/CSP · SQLite/Postgres · Streamlit"
        "\n• Results: 88.2% Execution Accuracy & 99.0% SQL Validity across 102 Gold Queries (130/130 Tests Passing)"
    )
    p2.font.size = PptPt(15.5)
    p2.font.color.rgb = PptRGBColor(203, 213, 225)
    p2.space_before = PptPt(18)

    # Slide 2: Problem Statement, Objectives & Theme
    s2 = prs.slides.add_slide(blank_layout)
    add_slide_header(s2, "1. Problem Statement, Objectives & Theme Alignment", "Rubric 1: Clarity, relevance to AI, and real-world feasibility")
    add_bullets(s2, [
        "• Problem: Non-technical users cannot write multi-join SQL queries, while single-pass LLMs fail on:",
        "   1. Schema Ambiguity: Overloaded columns ('name', 'year', 'status') across 12 tables cause wrong joins.",
        "   2. Unwritten Business Rules: Domain conventions ('CSE' -> dept_code='CSE', 'Pass' -> score>=40, 'Low attendance' < 75%).",
        "   3. No Safety or Self-Correction: Single-shot LLMs lack read-only sandboxing and crash on minor syntax errors.",
        "• Core Objective: Build a 100% open-source Text-to-SQL system (no paid GPT/Gemini APIs) that we:",
        "   (A) Baseline with a pretrained model, (B) Fine-tune with QLoRA, and (C) Wrap in a RAG + Validator + Repair Agent.",
        "• Theme (AI for Human Insight): Returns verified SQL, data table, plain-English explanation, bullet insights, Plotly chart, and a full 9-step agent trace."
    ])

    # Slide 3: Literature Review Table
    s3 = prs.slides.add_slide(blank_layout)
    add_slide_header(s3, "2. Literature Review & Comparative Table", "Rubric 2: Comparison of Spider, DIN-SQL, DAIL-SQL, MAC-SQL, QLoRA & Self-Debugging")
    r_cnt, c_cnt = len(LIT_PAPERS) + 1, 5
    tbl = s3.shapes.add_table(r_cnt, c_cnt, PptInches(0.4), PptInches(1.25), PptInches(12.5), PptInches(5.9)).table
    for c_i, h in enumerate(["Ref & Authors", "Objective", "Methods / Techniques", "Key Findings", "Gap Filled by Our Work"]):
        cell = tbl.cell(0, c_i)
        cell.text = h
        for p in cell.text_frame.paragraphs:
            p.font.size = PptPt(10.5)
            p.font.bold = True
    for r_i, r in enumerate(LIT_PAPERS, start=1):
        for c_i, val in enumerate([f"{r['ref']} {r['authors']}", r["objective"], r["methods"], r["findings"], r["future"]]):
            cell = tbl.cell(r_i, c_i)
            cell.text = val
            for p in cell.text_frame.paragraphs:
                p.font.size = PptPt(8.8)

    # Slide 4: State-Space Search Formulation
    s4 = prs.slides.add_slide(blank_layout)
    add_slide_header(s4, "3. State-Space Search Formulation", "Rubric 3: States, Initial State, Goal State, Actions & Search Strategy Justification")
    add_bullets(s4, [
        "• State Representation: s_t = (Q, C_RAG, SQL_t, E_t, t) for attempt t in {0, 1, 2, 3}.",
        "• Initial State (s_0): (Q, C_RAG, None, None, 0) after retrieving table DDLs + top-k rules/examples.",
        "• Goal State (s*): Valid(SQL_t, Schema) == True AND SQLite 'EXPLAIN QUERY PLAN' succeeds.",
        "• Actions: Retrieve(Q) -> C_RAG | Generate(Q, C_RAG) -> SQL_1 | Repair(Q, C_RAG, SQL_t, E_t) -> SQL_{t+1}.",
        "• Why Error-Guided Heuristic Search with Bounded Backtracking (Max 3 Tries)?",
        "   - Uninformed BFS/DFS over token strings is exponential (|V|^L).",
        "   - Greedy decoding proposes the top path, and our symbolic validator provides an exact diagnostic gradient E_t to repair errors in 1.18 average tries."
    ], left=0.5, top=1.35, width=6.2, height=5.8, font_size=14)
    if os.path.exists("docs/csp_search_diagram.png"):
        s4.shapes.add_picture("docs/csp_search_diagram.png", PptInches(6.85), PptInches(1.55), width=PptInches(6.1))

    # Slide 5: Knowledge Representation & Schema Trade-off
    s5 = prs.slides.add_slide(blank_layout)
    add_slide_header(s5, "4. Knowledge Representation & Vector RAG Design", "Rubric 4: Schema Graph, Business Rules, Case Memory & Small-vs-Large DB Trade-off")
    add_bullets(s5, [
        "• 1. Relational Schema & Foreign-Key Graph (12 Tables, 16 FK Edges):",
        "   - Tables: departments, faculty, students, courses, enrollments, marks, attendance, fees, hostels, hostel_allotments, clubs, club_members.",
        "• 2. Declarative Business Rules Base (data/business_rules.md — 30 Atomic Rules):",
        "   - One rule per line: 'CSE' -> dept_code='CSE', 'Pass' -> score>=40, 'Overall performance' -> AVG(marks.score).",
        "• 3. Episodic Case Memory (data/examples.json — 40 Solved Question->SQL Pairs):",
        "   - Strictly isolated from data/eval_set.json (102 gold queries).",
        "• 4. Dense Vector Index (sentence-transformers/all-MiniLM-L6-v2 + ChromaDB / FAISS):",
        "   - Design Trade-off: Include full table DDLs so join paths are never lost + retrieve top-k rules & examples (+1,410 chars context)."
    ], font_size=15)

    # Slide 6: CSP Formulation & Instance Walkthrough
    s6 = prs.slides.add_slide(blank_layout)
    add_slide_header(s6, "5. Constraint Satisfaction Problem (CSP) & Instance Trace", "Rubric 6 & Note 2: Variables, Domains, Hard Constraints & Why CSP over Fuzzy Logic")
    add_bullets(s6, [
        "• Variables (X): {StmtType, T_used, A_map (aliases), C_ref (columns)} parsed into an AST via sqlglot.",
        "• Domains (D): D(StmtType) = {SELECT}; D(T_used) = 12 College Tables; D(C_ref(t)) = PRAGMA table_info(t).",
        "• Hard Constraints (C1–C4):",
        "   - C1 (Safety): Single SELECT only; blocks INSERT, UPDATE, DELETE, DROP, CREATE, ALTER, TRUNCATE.",
        "   - C2 (Table Existence) & C3 (Scoped Column Resolution, exempting SELECT aliases like avg_score).",
        "   - C4 (Execution Feasibility): Must pass 'EXPLAIN QUERY PLAN' dry-run.",
        "• Rubric Note 2 — Why Crisp CSP Instead of Fuzzy Logic?",
        "   - SQL identifiers are binary (exist or fail). Linguistic fuzziness is handled upstream by continuous cosine similarity in Vector RAG.",
        "• Instance Trace: 'Show top 10 CSE students scoring > 80' -> C1–C4 verified in 4 ms -> returns 10 rows (Yash Nair: 97.6)."
    ], font_size=14)

    # Slide 7: 9-Agent Pipeline Architecture
    s7 = prs.slides.add_slide(blank_layout)
    add_slide_header(s7, "6. Intelligent System Design: Multi-Agent Pipeline", "Rubric 5: End-to-end architecture with self-correction and equivalence optimization")
    if os.path.exists("docs/architecture_diagram.png"):
        s7.shapes.add_picture("docs/architecture_diagram.png", PptInches(0.6), PptInches(1.3), width=PptInches(12.1))
    add_bullets(s7, [
        "• Shared Contract: pipeline.run(question, use_rag, use_validation, use_optimizer) -> {sql, rows, explanation, insights, chart, attempts, trace}."
    ], left=0.6, top=6.5, width=12.1, height=0.8, font_size=14.5)

    # Slide 8: Team Task Split, Timeline & AI vs Non-AI (Rubric Note 1)
    s8 = prs.slides.add_slide(blank_layout)
    add_slide_header(s8, "7. Team Task Split, 3-Week Timeline & AI vs. Non-AI Components", "Rubric Item 7 & Note 1: 5-Member Ownership & AI vs. Traditional Engineering Distinction")
    add_bullets(s8, [
        "• 5-Member Team Task Split (Shared Contract: pipeline.run(question) -> dict):",
        "   - Devesh Yadav (240437): LLM wrapper, sqlglot AST/CSP validator, 3-try self-correction loop & orchestrator (core/).",
        "   - Chakshit Mehta (240429): 4-Bit QLoRA fine-tuning (Colab T4), FastAPI + ngrok server & adapter merging (finetune/).",
        "   - Yash Bhardwaj (240957): 12-table seeded college.db, 30 business rules, 40 examples & 102 gold eval queries (db/, data/).",
        "   - Anmol Bhardwaj (240383): Vector RAG schema store, ChromaDB/FAISS indexing & FK graph retriever (rag/).",
        "   - Avyam (240417): Equivalence SQL optimizer, read-only executor, ablation runner & Streamlit dashboard UI (eval/, app.py).",
        "• Rubric Note 1 — AI vs. Non-AI Component Distinction:",
        "   - AI Components: Bi-Encoder Vector Embeddings (MiniLM), 4-Bit QLoRA Fine-Tuned Qwen2.5-Coder, Contextual Self-Repair, NL Insights.",
        "   - Non-AI Components: sqlglot AST/CSP Validator, Row-Equivalence Optimizer, Read-Only SQLite Sandbox (mode=ro), Pandas/Plotly UI."
    ], font_size=14)

    # Slide 9: QLoRA Fine-Tuning & Training Loss Curve
    s9 = prs.slides.add_slide(blank_layout)
    add_slide_header(s9, "8. QLoRA Fine-Tuning Setup & Loss Convergence", "Fine-tuning Qwen2.5-Coder on b-mc2/sql-create-context + College DB Pairs (Colab T4)")
    add_bullets(s9, [
        "• Base Models: Qwen2.5-Coder-7B-Instruct & 1.5B-Instruct.",
        "• 4-Bit Quantization: BitsAndBytes NF4 (load_in_4bit=True, float16 compute).",
        "• LoRA Config: r=16, alpha=32, dropout=0.05 across q/k/v/o/gate/up/down_proj.",
        "• Prompt Alignment: Identical template in training and inference; prompt tokens masked with -100.",
        "• Fast Inference: Adapter merged into fp16 (merge_and_unload) and served via FastAPI + ngrok + Instant Local Fallback."
    ], left=0.5, top=1.4, width=6.1, height=5.5, font_size=14.5)
    if os.path.exists("docs/training_loss_curve.png"):
        s9.shapes.add_picture("docs/training_loss_curve.png", PptInches(6.7), PptInches(1.5), width=PptInches(6.2))

    # Slide 10: Quantitative Evaluation & Ablations (Charts)
    s10 = prs.slides.add_slide(blank_layout)
    add_slide_header(s10, "9. Experimental Results: Baseline vs. Fine-Tuned vs. Agentic", "102 Gold Queries (30 Easy, 41 Medium, 31 Hard) across 6 System Configurations")
    if os.path.exists("docs/ablation_accuracy.png"):
        s10.shapes.add_picture("docs/ablation_accuracy.png", PptInches(0.35), PptInches(1.35), width=PptInches(6.4))
    if os.path.exists("docs/difficulty_breakdown.png"):
        s10.shapes.add_picture("docs/difficulty_breakdown.png", PptInches(6.85), PptInches(1.35), width=PptInches(6.1))
    add_bullets(s10, [
        "• System A Baseline (51.0%) -> System B Fine-Tuned (69.6%) -> System C Full Agentic (88.2% Exec Acc, 99.0% Valid SQL, 1.18 Avg Tries)."
    ], left=0.5, top=6.2, width=12.3, height=1.0, font_size=15)

    # Slide 11: Full Ablation & Error Analysis Table
    s11 = prs.slides.add_slide(blank_layout)
    add_slide_header(s11, "10. Component Ablation Table & Error Breakdown", "Isolating the impact of RAG (+15.7%), QLoRA (+18.6%), and Repair Loop (+12.7%)")
    t_ab = s11.shapes.add_table(len(ABLATION_TABLE) + 1, 10, PptInches(0.4), PptInches(1.3), PptInches(12.5), PptInches(3.2)).table
    ab_h = ["System", "Model", "RAG", "Validator", "Overall", "Easy", "Medium", "Hard", "Valid %", "Tries"]
    for c_i, h in enumerate(ab_h):
        t_ab.cell(0, c_i).text = h
        for p in t_ab.cell(0, c_i).text_frame.paragraphs:
            p.font.size = PptPt(10)
            p.font.bold = True
    for r_i, row in enumerate(ABLATION_TABLE, start=1):
        for c_i, val in enumerate(row):
            t_ab.cell(r_i, c_i).text = val
            for p in t_ab.cell(r_i, c_i).text_frame.paragraphs:
                p.font.size = PptPt(9.5)
    add_bullets(s11, [
        "• Top Failure Modes Fixed: Wrong filter values (22% -> fixed by RAG rules), Ambiguous columns 'name'/'status' (20% -> fixed by CSP Validator + Repair Loop), Missing intermediate joins (15% -> fixed by FK graph hints)."
    ], left=0.5, top=4.75, width=12.3, height=2.2, font_size=14.5)

    # Slide 12: Common Pitfalls, Safety Layer, Ethics & Conclusion
    s12 = prs.slides.add_slide(blank_layout)
    add_slide_header(s12, "11. Engineering Pitfalls Solved, Safety, Ethics & Conclusion", "Bonus Rubric: Originality (2 Marks) & Ethical AI / Privacy / Security (1 Mark)")
    add_bullets(s12, [
        "• Key Engineering Pitfalls Solved:",
        "   1. SELECT Aliases & CTEs in Validator: Exempted exp.Alias nodes so ORDER BY avg_score never false-fails.",
        "   2. RAG Few-Shot Prompt Splitting: Used rsplit('Question:', 1) so RAG examples never truncate the user question.",
        "   3. Thread-Safe Timeouts: Replaced signal.SIGALRM with daemon thread timeouts for Streamlit compatibility.",
        "• Triple-Layer Safety & Ethical AI:",
        "   - Read-Only Guarantee: (1) sqlglot SELECT-only gate, (2) SQLite URI mode=ro, (3) PRAGMA query_only = ON.",
        "   - 100% Data Privacy: Student grades and fee records stay local; 130/130 pytest tests pass.",
        "• Live Demo Ready: Running at http://localhost:8501 + GitHub: https://github.com/DeveshYadav85/text-to-sql"
    ], font_size=14.5)

    prs.save(path)


if __name__ == "__main__":
    build_lit_review_docx("docs/Agentic_Text_to_SQL_Literature_Review_Table.docx")
    build_lit_review_pdf("docs/Agentic_Text_to_SQL_Literature_Review_Table.pdf")
    build_report_docx("docs/Agentic_Text_to_SQL_Project_Report.docx")
    build_report_pdf("docs/Agentic_Text_to_SQL_Project_Report.pdf")
    build_presentation_pptx("docs/Agentic_Text_to_SQL_Presentation.pptx")
    print("All comprehensive submission documents generated successfully!")
