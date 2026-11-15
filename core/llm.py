"""
LLM wrapper — supports Ollama, HuggingFace, and remote Colab API.
"""
import re
import sys
import requests
import config


class ColabBackend:
    """Calls the fine-tuned Qwen model hosted on Google Colab via FastAPI + ngrok."""

    def __init__(self, base_url=None, api_token=None):
        self.base_url = (base_url or config.COLAB_API_URL).strip().rstrip("/")
        self.api_token = (api_token or config.COLAB_API_TOKEN).strip()

        if not self.base_url:
            raise ValueError("COLAB_API_URL is missing.")
        if not self.api_token:
            raise ValueError("COLAB_API_TOKEN is missing.")

    @staticmethod
    def _extract_schema_and_question(prompt):
        """
        Convert any pipeline prompt (SQL generation, optimization, or explanation)
        into the separate `schema` and `question` fields expected by the Colab `/generate` API.
        Uses rsplit("Question:", 1) so few-shot RAG examples ("Question: ...") inside
        the schema context do NOT break extraction!
        """
        try:
            from core.prompts import SYSTEM
            if prompt.startswith(SYSTEM):
                prompt = prompt[len(SYSTEM):].lstrip()
        except ImportError:
            pass

        if "Question:" in prompt:
            # Always split on the LAST "Question:" because RAG few-shot examples
            # earlier in the prompt also start with "Question:"
            schema_text, remaining = prompt.rsplit("Question:", 1)

            # Remove optional self-correction suffix or final instruction
            # while keeping the error message attached to the question if present
            if "\nThe previous SQL had an error:" in remaining:
                q_part, err_part = remaining.split("\nThe previous SQL had an error:", 1)
                err_part = re.split(r"\nWrite the SQL query:", err_part, maxsplit=1)[0]
                question = f"{q_part.strip()}\n(Fix previous error: {err_part.strip()})"
            else:
                question = re.split(r"\nWrite the SQL query:", remaining, maxsplit=1)[0].strip()

            schema_text = schema_text.strip()
            if not schema_text:
                schema_text = "SQLite database"
            return schema_text, question

        if "Original SQL:" in prompt:
            # Optimization prompt
            schema_text, remaining = prompt.split("Original SQL:", 1)
            return schema_text.strip() or "SQLite database", f"Optimize this SQL:\n{remaining.strip()}"

        # Fallback for any other prompt format
        return "SQLite database", prompt.strip()

    def generate(self, prompt, max_tokens=256):
        schema_text, question = self._extract_schema_and_question(prompt)

        payload = {
            "schema": schema_text,
            "question": question,
            "max_tokens": min(max_tokens, 256),
        }

        headers = {
            "Authorization": f"Bearer {self.api_token}",
            "ngrok-skip-browser-warning": "true",
        }

        try:
            response = requests.post(
                f"{self.base_url}/generate",
                json=payload,
                headers=headers,
                timeout=90,
            )
            response.raise_for_status()
            result = response.json()

            if not result.get("success") or not result.get("sql"):
                raise RuntimeError(f"Colab API returned an unexpected response: {result}")

            return result["sql"]

        except requests.exceptions.Timeout as e:
            raise RuntimeError(
                "Colab API timed out. Check whether the Colab model cell "
                "is running (use the merged float16 cell, not 4-bit) and the ngrok tunnel is active."
            ) from e

        except requests.exceptions.ConnectionError as e:
            raise RuntimeError(
                "Cannot connect to the Colab API. "
                "Check that your Colab runtime and ngrok tunnel are running."
            ) from e

        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response is not None else None
            detail = e.response.text if e.response is not None else str(e)

            if status == 503 and "ERR_NGROK_3004" in detail:
                raise RuntimeError(
                    "The Colab model server returned an invalid response "
                    "(ngrok ERR_NGROK_3004). Restart the model-serving cell "
                    "in Colab, update COLAB_API_URL in .env to the current "
                    "ngrok URL, and retry."
                ) from e

            raise RuntimeError(f"Colab API request failed (HTTP {status}): {detail}") from e


class OllamaBackend:
    """Calls a locally running Ollama server."""

    def __init__(self, model=None, base_url=None):
        self.model = model or config.OLLAMA_MODEL
        self.base_url = (base_url or config.OLLAMA_URL).strip().rstrip("/")

    def _get_local_models(self):
        try:
            r = requests.get(f"{self.base_url}/api/tags", timeout=5)
            r.raise_for_status()
            payload = r.json() or {}
            models = payload.get("models", [])
            return {
                m.get("name")
                for m in models
                if isinstance(m, dict) and m.get("name")
            }
        except Exception:
            return set()

    def generate(self, prompt, max_tokens=512):
        payload = {
            "model": self.model,
            "prompt": prompt,
            "stream": False,
            "options": {"num_predict": max_tokens, "temperature": 0},
        }
        try:
            r = requests.post(
                f"{self.base_url}/api/generate",
                json=payload,
                timeout=120,
            )
            r.raise_for_status()
            return r.json().get("response", "")
        except requests.exceptions.ConnectionError:
            raise RuntimeError(
                "Cannot reach Ollama. Start it with: ollama serve\n"
                f"Then pull the model: ollama pull {self.model}"
            )
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response is not None else None
            if status == 404:
                if self.model == getattr(config, "OLLAMA_FINETUNED_MODEL", ""):
                    raise RuntimeError(
                        f"Ollama model '{self.model}' was not found locally. "
                        "Train and export it using the README instructions."
                    ) from e
                raise RuntimeError(
                    f"Ollama model '{self.model}' was not found locally. "
                    f"Pull it with: ollama pull {self.model}"
                ) from e
            raise RuntimeError(
                f"Ollama request failed for model '{self.model}' (HTTP {status})."
            ) from e

    def is_available(self):
        local = self._get_local_models()
        if not local:
            return False
        return self.model in local or f"{self.model}:latest" in local


class HuggingFaceBackend:
    """Loads a HuggingFace causal-LM locally (CPU / MPS / CUDA)."""

    def __init__(self, model_path=None):
        import torch
        from transformers import AutoTokenizer, AutoModelForCausalLM

        model_path = model_path or config.best_hf_model()
        print(f"[HF] loading {model_path} ...", file=sys.stderr)

        self.tok = AutoTokenizer.from_pretrained(model_path)
        if torch.cuda.is_available():
            dtype, dev = torch.float16, "cuda"
        elif torch.backends.mps.is_available():
            dtype, dev = torch.float16, "mps"
        else:
            dtype, dev = torch.float32, "cpu"

        self.model = AutoModelForCausalLM.from_pretrained(
            model_path, torch_dtype=dtype, low_cpu_mem_usage=True
        ).to(dev)
        self.device = dev

    def generate(self, prompt, max_tokens=512):
        inputs = self.tok(prompt, return_tensors="pt").to(self.device)
        out = self.model.generate(
            **inputs,
            max_new_tokens=max_tokens,
            do_sample=False,
            pad_token_id=self.tok.eos_token_id,
        )
        new_tokens = out[0][inputs.input_ids.shape[1]:]
        return self.tok.decode(new_tokens, skip_special_tokens=True)


class LLM:
    """Single entry point for all LLM calls."""

    def __init__(self, backend=None):
        b = (backend or config.LLM_BACKEND).strip().lower()
        self.backend_name = b
        if b == "ollama":
            self._backend = OllamaBackend()
        elif b == "huggingface":
            self._backend = HuggingFaceBackend()
        elif b == "colab":
            self._backend = ColabBackend()
        else:
            raise ValueError(f"Unsupported LLM_BACKEND: {b}")

    def generate(self, prompt, max_tokens=512):
        return self._backend.generate(prompt, max_tokens)

    @staticmethod
    def extract_sql(text):
        """Pull SQL out of markdown fences or return raw text."""
        m = re.search(r"```(?:sql)?\s*(.*?)```", text, re.S | re.I)
        sql = m.group(1) if m else text
        sql = sql.strip().rstrip(";").strip()
        return sql + ";"


_instance = None
_instance_key = None


def get_llm():
    global _instance, _instance_key
    key = (
        config.LLM_BACKEND.strip().lower(),
        config.COLAB_API_URL.strip(),
        config.COLAB_API_TOKEN.strip(),
        config.OLLAMA_MODEL.strip(),
    )
    if _instance is None or _instance_key != key:
        _instance = LLM()
        _instance_key = key
    return _instance
