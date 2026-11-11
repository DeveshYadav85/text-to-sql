"""
LLM wrapper — primary backend is Ollama (fully local, no paid APIs).
Fallback: HuggingFace model loaded locally (needs torch/transformers installed).

Usage:
    llm = LLM()                      # uses config defaults
    sql = llm.generate(prompt)       # returns plain text
"""
import re
import sys
import requests
import config


class OllamaBackend:
    """Calls a locally running Ollama server."""

    def __init__(self, model=None, base_url=None):
        self.model = model or config.OLLAMA_MODEL
        self.base_url = (base_url or config.OLLAMA_URL).rstrip("/")

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
                "Cannot reach Ollama. Start it with:  ollama serve\n"
                f"Then pull the model:  ollama pull {self.model}"
            )

    def is_available(self):
        try:
            r = requests.get(f"{self.base_url}/api/tags", timeout=5)
            return r.status_code == 200
        except Exception:
            return False


class HuggingFaceBackend:
    """Loads a HuggingFace causal-LM locally (CPU / MPS / CUDA)."""

    def __init__(self, model_path=None):
        import torch
        from transformers import AutoTokenizer, AutoModelForCausalLM

        model_path = model_path or config.HF_MODEL
        print(f"[HF] loading {model_path} ...", file=sys.stderr)

        self.tok = AutoTokenizer.from_pretrained(model_path)
        if torch.cuda.is_available():
            dtype = torch.float16
            dev = "cuda"
        elif torch.backends.mps.is_available():
            dtype = torch.float16
            dev = "mps"
        else:
            dtype = torch.float32
            dev = "cpu"

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
    """
    Single entry point for all LLM calls.
    Backend is selected by config.LLM_BACKEND ('ollama' or 'huggingface').
    """

    def __init__(self, backend=None):
        b = backend or config.LLM_BACKEND
        if b == "ollama":
            self._backend = OllamaBackend()
        else:
            self._backend = HuggingFaceBackend()

    def generate(self, prompt, max_tokens=512):
        """Return raw text from the model."""
        return self._backend.generate(prompt, max_tokens)

    @staticmethod
    def extract_sql(text):
        """Pull SQL out of markdown fences or return raw text."""
        m = re.search(r"```(?:sql)?\s*(.*?)```", text, re.S | re.I)
        sql = m.group(1) if m else text
        # strip any trailing semicolons then add exactly one
        sql = sql.strip().rstrip(";").strip()
        return sql + ";"


# module-level singleton so callers can do:  from core.llm import llm
_instance = None

def get_llm():
    global _instance
    if _instance is None:
        _instance = LLM()
    return _instance
