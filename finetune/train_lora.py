"""
finetune/train_lora.py
LoRA fine-tuning of Qwen2.5-Coder on SQL pairs from this project's examples file.
Requires a GPU (Colab T4 or better).  Run on Colab or a Linux machine with CUDA.

This file is NOT needed to run the main project.
The project works out of the box with ollama pull qwen2.5-coder:7b.
"""
# NOTE: install on Colab first:
#   pip install transformers peft datasets trl accelerate bitsandbytes torch

import json, sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

PROMPT_TEMPLATE = (
    "You are an expert SQL assistant.\n\n"
    "Database schema:\n{schema}\n\n"
    "Question: {question}\n\n"
    "SQL:"
)
OUTPUT_DIR = "finetune/out/lora"
MODEL_NAME = "Qwen/Qwen2.5-Coder-3B-Instruct"


def load_examples():
    """Load project examples + optionally a HuggingFace dataset."""
    with open("data/examples.json") as f:
        local = json.load(f)
    pairs = [{"question": e["q"], "sql": e["sql"]} for e in local]

    # optionally add b-mc2/sql-create-context (small public dataset)
    try:
        from datasets import load_dataset
        remote = load_dataset("b-mc2/sql-create-context", split="train").select(range(500))
        for row in remote:
            pairs.append({"question": row["question"], "sql": row["answer"],
                          "schema": row.get("context", "")})
        print(f"loaded {len(remote)} rows from b-mc2/sql-create-context")
    except Exception as e:
        print(f"skipping remote dataset: {e}")

    return pairs


def build_dataset(pairs, schema_str):
    from datasets import Dataset

    def fmt(ex):
        schema = ex.get("schema", schema_str)
        text = (
            PROMPT_TEMPLATE.format(schema=schema, question=ex["question"])
            + " " + ex["sql"]
        )
        return {"text": text}

    return Dataset.from_list([fmt(p) for p in pairs])


def train():
    import torch
    from transformers import AutoTokenizer, AutoModelForCausalLM, BitsAndBytesConfig
    from peft import LoraConfig, get_peft_model, TaskType
    from trl import SFTTrainer, SFTConfig

    # get schema
    from sqlalchemy import create_engine, inspect
    import config
    engine = create_engine(config.DATABASE_URL)
    insp   = inspect(engine)
    schema_lines = []
    for t in insp.get_table_names():
        cols = ", ".join(c["name"] for c in insp.get_columns(t))
        schema_lines.append(f"{t}({cols})")
    schema_str = "\n".join(schema_lines)

    pairs   = load_examples()
    dataset = build_dataset(pairs, schema_str)
    print(f"training on {len(dataset)} examples")

    bnb_config = BitsAndBytesConfig(
        load_in_4bit=True,
        bnb_4bit_compute_dtype=torch.float16,
        bnb_4bit_quant_type="nf4",
    )
    tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
    model = AutoModelForCausalLM.from_pretrained(
        MODEL_NAME,
        quantization_config=bnb_config,
        device_map="auto",
    )

    lora_config = LoraConfig(
        r=16, lora_alpha=32,
        target_modules=["q_proj", "v_proj"],
        lora_dropout=0.05,
        bias="none",
        task_type=TaskType.CAUSAL_LM,
    )
    model = get_peft_model(model, lora_config)
    model.print_trainable_parameters()

    sft_config = SFTConfig(
        output_dir=OUTPUT_DIR,
        num_train_epochs=3,
        per_device_train_batch_size=4,
        gradient_accumulation_steps=4,
        learning_rate=2e-4,
        fp16=True,
        logging_steps=10,
        save_strategy="epoch",
        dataset_text_field="text",
        max_seq_length=512,
    )
    trainer = SFTTrainer(
        model=model,
        train_dataset=dataset,
        args=sft_config,
        tokenizer=tokenizer,
    )
    trainer.train()
    trainer.save_model(OUTPUT_DIR)
    print(f"LoRA adapter saved to {OUTPUT_DIR}")
    print("Next: run  python finetune/merge.py  to merge into a full model")


if __name__ == "__main__":
    train()
