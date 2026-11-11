import os

# ── database ──────────────────────────────────────────────────────────────────
DB_PATH = os.getenv("DB_PATH", "data/college.db")
DATABASE_URL = os.getenv("DATABASE_URL", f"sqlite:///{DB_PATH}")

# ── llm ───────────────────────────────────────────────────────────────────────
# primary backend: ollama  (no paid APIs, runs fully local)
LLM_BACKEND = os.getenv("LLM_BACKEND", "ollama")   # "ollama" or "huggingface"
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "qwen2.5-coder:7b")
OLLAMA_URL = os.getenv("OLLAMA_URL", "http://localhost:11434")

# huggingface fallback (for machines without ollama)
HF_MODEL = os.getenv("HF_MODEL", "Qwen/Qwen2.5-Coder-3B-Instruct")
HF_FINETUNED = "finetune/out/merged"   # produced by finetune/merge.py

# ── rag ───────────────────────────────────────────────────────────────────────
EMBED_MODEL = "all-MiniLM-L6-v2"
CHROMA_DIR = "data/chroma"
RAG_TOP_K = 3          # top results per doc type (tables, rules, fk, examples)
RAG_EXPAND_FK = True   # expand retrieved tables with their FK neighbours

# ── pipeline ──────────────────────────────────────────────────────────────────
MAX_RETRIES = 3        # self-correction loop max attempts
MAX_ROWS = 200         # row cap for execution
QUERY_TIMEOUT = 30     # seconds

# data files
RULES_PATH = "data/business_rules.md"
EXAMPLES_PATH = "data/examples.json"

def best_hf_model():
    """Return fine-tuned model if available, else base."""
    return HF_FINETUNED if os.path.isdir(HF_FINETUNED) else HF_MODEL
