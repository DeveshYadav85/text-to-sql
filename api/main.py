import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel
import config
from core.agent import Text2SQLAgent
from rag.schema_store import SchemaStore

app = FastAPI(title="Text2SQL API", version="1.0.0")

# ── lazy-load agent so startup is fast ──────────────────────────────────────
_agent: Text2SQLAgent | None = None
_store: SchemaStore | None = None


def get_store() -> SchemaStore:
    global _store
    if _store is None:
        _store = SchemaStore(embed=True)
    return _store


def get_agent() -> Text2SQLAgent:
    global _agent
    if _agent is None:
        _agent = Text2SQLAgent(model_path=config.best_model(), store=get_store())
    return _agent


# ── request / response models ────────────────────────────────────────────────
class QueryRequest(BaseModel):
    question: str


class QueryResponse(BaseModel):
    sql: str
    columns: list[str]
    rows: list[list]
    explanation: str
    attempts: int
    error: str | None


# ── API routes ───────────────────────────────────────────────────────────────
@app.get("/health")
def health():
    return {"status": "ok", "model": config.best_model()}


@app.get("/schema")
def schema():
    """Return table → column mapping for the UI schema panel."""
    store = get_store()
    return {"tables": store.tables}


@app.post("/query", response_model=QueryResponse)
def query(req: QueryRequest):
    if not req.question.strip():
        raise HTTPException(status_code=400, detail="Question cannot be empty")
    try:
        result = get_agent().answer(req.question)
        return QueryResponse(
            sql=result["sql"] or "",
            columns=result["columns"],
            rows=[list(r) for r in result["rows"]],
            explanation=result["explanation"],
            attempts=result["attempts"],
            error=result["error"],
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# ── serve frontend ───────────────────────────────────────────────────────────
_frontend = os.path.join(os.path.dirname(__file__), "..", "frontend")
app.mount("/static", StaticFiles(directory=_frontend), name="static")


@app.get("/")
def root():
    return FileResponse(os.path.join(_frontend, "index.html"))
