"""
Streamlit UI for Agentic Text-to-SQL.
Run:  streamlit run app.py
"""
import os, sys, json
sys.path.insert(0, os.path.dirname(__file__))

import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Agentic Text-to-SQL",
    page_icon=":material/database:",
    layout="wide",
)

# ── sidebar settings ───────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Settings")
    use_rag        = st.toggle("Use RAG (vector retrieval)", value=True)
    use_validation = st.toggle("Use validation loop", value=True)
    use_optimizer  = st.toggle("Use SQL optimizer", value=True)
    model_name     = st.text_input("Ollama model", value="qwen2.5-coder:7b")
    db_url         = st.text_input("Database URL", value="sqlite:///data/college.db")
    st.divider()
    st.caption("All 9 pipeline steps are shown in expanders below the result.")

# apply sidebar overrides to config at runtime
import config
config.OLLAMA_MODEL = model_name
config.DATABASE_URL = db_url
if "sqlite" in db_url:
    # derive DB_PATH from URL for read-only executor
    config.DB_PATH = db_url.replace("sqlite:///", "")

# ── header ────────────────────────────────────────────────────────────────────
st.title(":material/database: Agentic Text-to-SQL")
st.caption(
    "Ask a plain-English question. The 9-agent pipeline generates SQL, "
    "validates it (with self-correction), optimises it, runs it, and explains the result."
)

# ── input ─────────────────────────────────────────────────────────────────────
question = st.text_input(
    "Your question",
    placeholder="e.g. Show the top 10 students from CSE who scored above 80",
    label_visibility="collapsed",
)

col_run, col_eg = st.columns([1, 5])
with col_run:
    run_btn = st.button(":material/play_arrow: Run", width="stretch")
with col_eg:
    examples = [
        "Show the top 10 students from CSE who scored above 80",
        "Which departments have more than 20 students?",
        "Show students with low attendance",
        "List all fee defaulters",
        "Average score per department",
    ]
    chosen = st.selectbox("Quick examples", [""] + examples, label_visibility="collapsed")
    if chosen and not question:
        question = chosen

if not run_btn or not question.strip():
    st.stop()

# ── run pipeline ──────────────────────────────────────────────────────────────
# check ollama is reachable before spinning
from core.llm import OllamaBackend
backend = OllamaBackend(model=model_name)
if not backend.is_available():
    st.error(
        "**Ollama is not running.**  \n"
        "Start it in a terminal:  `ollama serve`  \n"
        f"Then pull the model:  `ollama pull {model_name}`",
        icon=":material/error:"
    )
    st.stop()

with st.spinner("Running 9-agent pipeline…"):
    from core.pipeline import run as pipeline_run
    result = pipeline_run(
        question,
        use_rag=use_rag,
        use_validation=use_validation,
        use_optimizer=use_optimizer,
    )

# ── result display ────────────────────────────────────────────────────────────
if result["error"]:
    st.error(f"**Error:** {result['error']}", icon=":material/error:")
else:
    st.success("Query complete", icon=":material/check_circle:")

st.divider()

# SQL block
sql_to_show = result.get("optimized_sql") or result.get("sql", "")
st.subheader(":material/code: Generated SQL")
st.code(sql_to_show, language="sql")
if result.get("optimized_sql"):
    st.caption(f"Original SQL was optimised. Attempts: {result['attempts']}")
else:
    st.caption(f"Attempts: {result['attempts']}")

# results table
rows    = result.get("rows", [])
columns = result.get("columns", [])
if rows:
    st.subheader(f":material/table_rows: Results  ({len(rows)} rows)")
    df = pd.DataFrame(rows, columns=columns)
    st.dataframe(df, alt="Query results", use_container_width=False)
elif not result["error"]:
    st.info("Query returned no rows.")

# chart
chart = result.get("chart")
if chart:
    st.subheader(":material/bar_chart: Chart")
    st.plotly_chart(chart, alt="Result chart")

# explanation + insights
explanation = result.get("explanation", "")
insights    = result.get("insights", [])
if explanation:
    st.subheader(":material/lightbulb: Explanation")
    st.write(explanation)
if insights:
    for ins in insights:
        st.markdown(f"- {ins}")

st.divider()

# ── pipeline trace (one expander per step) ────────────────────────────────────
st.subheader(":material/account_tree: Pipeline trace")

STEP_ICONS = {
    "schema_rag":       ":material/search:",
    "sql_generation":   ":material/code:",
    "sql_validation":   ":material/verified:",
    "sql_optimization": ":material/tune:",
    "execution":        ":material/play_circle:",
    "result_processing":":material/table_chart:",
    "explanation":      ":material/lightbulb:",
}

for entry in result.get("trace", []):
    name  = entry["step"]
    icon  = STEP_ICONS.get(name, ":material/circle:")
    t_ms  = int(entry["time_s"] * 1000)
    label = f"{icon} **{name.replace('_', ' ').title()}** — {t_ms} ms"

    with st.expander(label, expanded=False):
        out = entry["output"]
        if name == "schema_rag":
            rt = out.get("retrieval_trace", {})
            st.markdown("**Rules retrieved:**")
            for r in rt.get("rules", []):
                st.caption(r)
            st.markdown("**Examples retrieved:**")
            for e in rt.get("examples", []):
                st.caption(e)
            with st.expander("Full schema context"):
                st.code(rt.get("schema", out.get("context", "")), language="sql")

        elif name == "sql_generation":
            st.markdown("**Generated SQL:**")
            st.code(out.get("sql", ""), language="sql")
            with st.expander("Full prompt"):
                st.text(out.get("prompt", ""))

        elif name == "sql_validation":
            if out.get("valid"):
                st.success("Valid ✓")
            else:
                st.error(f"Invalid: {out.get('error')}")

        elif name == "sql_optimization":
            st.markdown(f"**Changed:** {out.get('changed')}")
            st.caption(out.get("reason", ""))
            if out.get("changed"):
                st.code(out.get("optimized_sql", ""), language="sql")

        elif name == "execution":
            st.markdown(f"**Rows returned:** {out.get('row_count', 0)}")
            if out.get("error"):
                st.error(out["error"])

        elif name == "result_processing":
            st.markdown(f"**Numeric columns:** {out.get('num_cols', [])}")
            st.markdown(f"**Categorical columns:** {out.get('cat_cols', [])}")
            if out.get("stats"):
                st.json(out["stats"])

        elif name == "explanation":
            st.write(out.get("explanation", ""))
            for ins in out.get("insights", []):
                st.caption(f"- {ins}")

        else:
            st.json(out)
