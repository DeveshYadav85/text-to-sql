import sqlite3
import numpy as np
import config


class SchemaStore:
    """Schema always goes in full (small DB). RAG is used to pick business rules.
    For a big schema, index the TABLE chunks too and retrieve only the relevant ones."""

    def __init__(self, embed=True):
        self.ddl = self._ddl()
        self.tables = self._tables_with_columns()
        with open(config.RULES_PATH) as f:
            self.rules = [l.strip() for l in f if l.strip()]
        self.model = self.index = None
        if embed:
            import faiss
            from sentence_transformers import SentenceTransformer
            self.model = SentenceTransformer(config.EMBED_MODEL)
            emb = self.model.encode(self.rules, normalize_embeddings=True)
            self.index = faiss.IndexFlatIP(emb.shape[1])
            self.index.add(np.array(emb, dtype="float32"))

    def _ddl(self):
        conn = sqlite3.connect(config.DB_PATH)
        out = [f"TABLE {n}:\n{sql}" for n, sql in
               conn.execute("SELECT name, sql FROM sqlite_master WHERE type='table'")]
        conn.close()
        return out

    def _tables_with_columns(self):
        conn = sqlite3.connect(config.DB_PATH)
        out = {}
        for (t,) in conn.execute("SELECT name FROM sqlite_master WHERE type='table'"):
            out[t] = [r[1] for r in conn.execute(f"PRAGMA table_info({t})")]
        conn.close()
        return out

    def retrieve_rules(self, question, k=config.TOP_K):
        q = self.model.encode([question], normalize_embeddings=True)
        _, idx = self.index.search(np.array(q, dtype="float32"), min(k, len(self.rules)))
        return [self.rules[i] for i in idx[0]]

    def full_schema(self):
        return "\n".join(self.ddl)
