import config
from core.validator import validate
from core.executor import execute


class Text2SQLAgent:
    def __init__(self, model_path=config.BASE_MODEL, use_rag=True, use_agent=True,
                 store=None, llm=None):
        # store / llm can be passed in (used by tests with fakes)
        if store is None:
            from rag.schema_store import SchemaStore
            store = SchemaStore(embed=use_rag)
        if llm is None:
            from core.llm import LLM
            llm = LLM.load(model_path)
        self.store, self.llm = store, llm
        self.use_rag, self.use_agent = use_rag, use_agent

    def _context(self, q):
        ctx = self.store.full_schema()
        if self.use_rag:
            ctx += "\n\nRULES:\n" + "\n".join(self.store.retrieve_rules(q))
        return ctx

    def answer(self, question, explain=True):
        ctx = self._context(question)
        sql, err, attempts = None, None, 0
        retries = config.MAX_RETRIES if self.use_agent else 1

        for _ in range(retries):
            attempts += 1
            sql = self.llm.generate_sql(question, ctx, err, sql)
            try:
                if self.use_agent:
                    validate(sql, self.store.tables)
                cols, rows = execute(sql)
                text = self.llm.explain(question, sql, rows) if (self.use_agent and explain) else ""
                return {"sql": sql, "columns": cols, "rows": rows,
                        "explanation": text, "attempts": attempts, "error": None}
            except Exception as e:
                err = str(e)
        return {"sql": sql, "columns": [], "rows": [], "explanation": "",
                "attempts": attempts, "error": err}
