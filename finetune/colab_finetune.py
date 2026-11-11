"""
Text2SQL — Google Colab Fine-Tuning Script
==========================================
Run this on Google Colab with a free T4 GPU.

Steps:
  1. Open Google Colab → Runtime → Change runtime type → GPU (T4)
  2. Upload the entire text2sql/ folder to Colab (or clone from GitHub)
  3. Run each cell in order

This script fine-tunes Qwen2.5-Coder-7B-Instruct using QLoRA on the
sql-create-context dataset (8000 examples), then merges the LoRA adapter
into the base model and saves it to finetune/out/merged/.
"""

# ── Cell 1: Install dependencies ─────────────────────────────────────────────
# !pip install -q torch transformers accelerate bitsandbytes peft trl datasets sentence-transformers faiss-cpu sqlglot fastapi uvicorn pandas

# ── Cell 2: Mount Drive (optional — to persist model between sessions) ────────
# from google.colab import drive
# drive.mount('/content/drive')
# import os
# os.chdir('/content/drive/MyDrive/text2sql')   # or wherever you uploaded the folder

# ── Cell 3: Verify setup ──────────────────────────────────────────────────────
import torch
print("CUDA available:", torch.cuda.is_available())
print("GPU:", torch.cuda.get_device_name(0) if torch.cuda.is_available() else "None")

import subprocess
result = subprocess.run(["python3", "tests/test_offline.py"], capture_output=True, text=True)
print(result.stdout)
if result.returncode != 0:
    print("ERRORS:", result.stderr)

# ── Cell 4: Fine-tune with LoRA (≈ 45-60 min on T4) ─────────────────────────
# Uses Qwen2.5-Coder-7B-Instruct + 4-bit QLoRA + 8000 SQL examples
# !python finetune/train_lora.py

# ── Cell 5: Merge LoRA adapter into base model ────────────────────────────────
# !python finetune/merge.py
# Saves merged model to finetune/out/merged/

# ── Cell 6: Run full evaluation (all 6 system variants) ──────────────────────
# !python -m eval.run_eval A B C D E F
#
# Output example:
#   A baseline:    exec_acc=42.50%  valid=82.50%  avg_attempts=1.00
#   B fine-tuned:  exec_acc=66.25%  valid=93.75%  avg_attempts=1.00
#   C full agentic:exec_acc=71.25%  valid=95.00%  avg_attempts=1.43
#   D base + rag:  exec_acc=48.75%  valid=85.00%  avg_attempts=1.00
#   E base + agent:exec_acc=52.50%  valid=88.75%  avg_attempts=1.28
#   F tuned + agent:exec_acc=68.75% valid=93.75%  avg_attempts=1.31

# ── Cell 7: Download the merged model (if not using Drive) ────────────────────
# from google.colab import files
# import shutil
# shutil.make_archive('merged_model', 'zip', 'finetune/out/merged')
# files.download('merged_model.zip')

# ── Cell 8: (Back on your Mac) Run the API server ────────────────────────────
# After copying finetune/out/merged/ back locally:
# $ python3 -m uvicorn api.main:app --reload --port 8000

print("""
=== COLAB WORKFLOW SUMMARY ===
1. pip install (Cell 1)
2. [optional] Mount Drive (Cell 2)
3. python tests/test_offline.py  → verify 0 problems
4. python finetune/train_lora.py → ~1h, saves to finetune/out/lora/
5. python finetune/merge.py      → saves to finetune/out/merged/
6. python -m eval.run_eval A B C D E F
7. Download finetune/out/merged/ → copy to your Mac
8. uvicorn api.main:app          → open http://localhost:8000
""")
