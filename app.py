"""
Streamlit UI for Agentic Text-to-SQL.
Run: streamlit run app.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import streamlit as st
import pandas as pd
from dotenv import load_dotenv

load_dotenv(override=True)

import config

st.set_page_config(
    page_title="Agentic Text-to-SQL",
    page_icon=":material/database:",
    layout="wide",
)

# ── sidebar settings ──────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Settings")
    use_rag = st.toggle("Use RAG (vector retrieval)", value=True)
    use_validation = st.toggle("Use validation loop", value=True)
    use_optimizer = st.toggle("Use SQL optimizer", value=True)

    backends = ["colab", "ollama", "huggingface"]
    default_idx = backends.index(config.LLM_BACKEND) if config.LLM_BACKEND in backends else 0
    selected_backend = st.selectbox("LLM Backend", backends, index=default_idx)
    config.LLM_BACKEND = selected_backend

    if selected_backend == "colab":
        colab_url = st.text_input(
            "Colab ngrok URL",
            value=config.COLAB_API_URL,
            placeholder="https://xxxx.ngrok-free.dev",
        )
        colab_token = st.text_input(
            "Colab API Token",
            value=config.COLAB_API_TOKEN,
            type="password",
        )
        config.COLAB_API_URL = colab_url.strip().rstrip("/")
        config.COLAB_API_TOKEN = colab_token.strip()
        if config.COLAB_API_URL and config.COLAB_API_TOKEN:
            st.success("Colab API configured")
    elif selected_backend == "ollama":
        model_name = st.text_input("Ollama model", value=config.OLLAMA_MODEL)
        config.OLLAMA_MODEL = model_name.strip()
    elif selected_backend == "huggingface":
        st.info(f"Using local HF model: {config.best_hf_model()}")

    db_url = st.text_input("Database URL", value=config.DATABASE_URL)
    config.DATABASE_URL = db_url.strip()
    if "sqlite" in config.DATABASE_URL:
        config.DB_PATH = config.DATABASE_URL.replace("sqlite:///", "")

    st.divider()
    st.caption("All 9 pipeline steps are shown in expanders below the result.")

# ── header ────────────────────────────────────────────────────────────────────
st.title(":material/database: Agentic Text-to-SQL")
st.caption(
    "Ask a plain-English question. The 9-agent pipeline generates SQL, "
    "validates it with self-correction, optimises it, executes it, "
    "and explains the result."
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

# ── backend configuration checks ──────────────────────────────────────────────
if config.LLM_BACKEND == "colab":
    if not config.COLAB_API_URL:
        st.error("COLAB_API_URL is missing. Set it in `.env` or in the sidebar.")
        st.stop()
    if not config.COLAB_API_TOKEN:
        st.error("COLAB_API_TOKEN is missing. Set it in `.env` or in the sidebar.")
        st.stop()

elif config.LLM_BACKEND == "ollama":
    from core.llm import OllamaBackend
    backend = OllamaBackend(model=config.OLLAMA_MODEL)
    if not backend.is_available():
        if config.OLLAMA_MODEL == config.OLLAMA_FINETUNED_MODEL:
            message = (
                "**The fine-tuned Ollama model is not installed.**\n\n"
                "Run the GPU training and GGUF export steps in the README, "
                "copy the GGUF into `finetune/out/`, then create it with:\n\n"
                "`ollama create text-to-sql-qwen2.5-coder:7b -f finetune/Modelfile`"
            )
        else:
            message = (
                "**Ollama is unavailable or the selected model is not installed.**\n\n"
                "Start it in a terminal using `ollama serve`.\n\n"
                f"Then install the model using:\n`ollama pull {config.OLLAMA_MODEL}`"
            )
        st.error(message, icon=":material/error:")
        st.stop()

# ── run pipeline ──────────────────────────────────────────────────────────────
try:
    with st.spinner("Running 9-agent pipeline…"):
        from core.pipeline import run as pipeline_run
        result = pipeline_run(
            question,
            use_rag=use_rag,
            use_validation=use_validation,
            use_optimizer=use_optimizer,
        )
except Exception as e:
    st.error(f"Pipeline execution failed: {str(e)}", icon=":material/error:")
    st.stop()

# ── result display ────────────────────────────────────────────────────────────
if result.get("error"):
    st.error(f"**Error:** {result['error']}", icon=":material/error:")
else:
    st.success("Query complete", icon=":material/check_circle:")

st.divider()

sql_to_show = result.get("optimized_sql") or result.get("sql", "")
st.subheader(":material/code: Generated SQL")
st.code(sql_to_show, language="sql")
if result.get("optimized_sql"):
    st.caption(f"Original SQL was optimised. Attempts: {result.get('attempts', 0)}")
else:
    st.caption(f"Attempts: {result.get('attempts', 0)}")

rows = result.get("rows", [])
columns = result.get("columns", [])
if rows:
    st.subheader(f":material/table_rows: Results ({len(rows)} rows)")
    df = pd.DataFrame(rows, columns=columns)
    st.dataframe(df, alt="Query results")
elif not result.get("error"):
    st.info("Query returned no rows.")

chart = result.get("chart")
if chart:
    st.subheader(":material/bar_chart: Chart")
    st.plotly_chart(chart, alt="Result chart")

explanation = result.get("explanation", "")
insights = result.get("insights", [])
if explanation:
    st.subheader(":material/lightbulb: Explanation")
    st.write(explanation)
if insights:
    st.subheader("Insights")
    for ins in insights:
        st.markdown(f"- {ins}")

st.divider()

# ── pipeline trace ────────────────────────────────────────────────────────────
st.subheader(":material/account_tree: Pipeline trace")

STEP_ICONS = {
    "schema_rag": ":material/search:",
    "sql_generation": ":material/code:",
    "sql_validation": ":material/verified:",
    "sql_optimization": ":material/tune:",
    "execution": ":material/play_circle:",
    "result_processing": ":material/table_chart:",
    "explanation": ":material/lightbulb:",
}

for entry in result.get("trace", []):
    name = entry["step"]
    icon = STEP_ICONS.get(name, ":material/circle:")
    t_ms = int(entry.get("time_s", 0) * 1000)
    label = f"{icon} **{name.replace('_', ' ').title()}** — {t_ms} ms"

    with st.expander(label, expanded=False):
        out = entry.get("output", {})
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
