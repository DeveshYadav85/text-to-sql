"""
Explanation agent — step 8 of the pipeline.
Uses the LLM to produce a plain-English explanation + key insights.
Also picks the best Plotly chart for the result shape.
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from core.llm import get_llm
from core.prompts import explanation_prompt
import plotly.express as px
import pandas as pd


def _pick_chart(df, num_cols, cat_cols):
    """
    Auto-select chart type from result shape:
      - 1 cat + 1 num  → bar chart
      - 1 num only     → histogram
      - 2+ num         → scatter
      - cat only       → bar of value counts
      - date + num     → line chart
    Returns a plotly Figure or None.
    """
    if df is None or df.empty:
        return None

    n_rows = len(df)
    # detect date-like columns
    date_cols = [c for c in cat_cols
                 if any(k in c.lower() for k in ["date", "year", "month", "day", "semester"])]

    try:
        if date_cols and num_cols:
            # line chart
            fig = px.line(df, x=date_cols[0], y=num_cols[0],
                          title=f"{num_cols[0]} over {date_cols[0]}")

        elif len(cat_cols) >= 1 and len(num_cols) >= 1 and n_rows <= 50:
            # bar chart
            fig = px.bar(df, x=cat_cols[0], y=num_cols[0],
                         title=f"{num_cols[0]} by {cat_cols[0]}",
                         color_discrete_sequence=["steelblue"])

        elif len(num_cols) == 1 and not cat_cols:
            # histogram
            fig = px.histogram(df, x=num_cols[0],
                               title=f"Distribution of {num_cols[0]}",
                               color_discrete_sequence=["steelblue"])

        elif len(num_cols) >= 2:
            # scatter
            fig = px.scatter(df, x=num_cols[0], y=num_cols[1],
                             title=f"{num_cols[0]} vs {num_cols[1]}",
                             color_discrete_sequence=["steelblue"])

        elif cat_cols and not num_cols and n_rows <= 20:
            # value counts as pie
            counts = df[cat_cols[0]].value_counts().reset_index()
            counts.columns = [cat_cols[0], "count"]
            fig = px.pie(counts, names=cat_cols[0], values="count",
                         title=f"Distribution of {cat_cols[0]}")

        else:
            return None

        fig.update_layout(margin=dict(t=40, b=20, l=20, r=20))
        return fig

    except Exception:
        return None


def _parse_insights(text):
    """Extract bullet lines starting with '-' from LLM output."""
    lines = text.split("\n")
    insights = [l.lstrip("- •*").strip() for l in lines if l.strip().startswith("-")]
    # if LLM didn't bullet, split on numbered list
    if not insights:
        import re
        insights = re.findall(r'\d+\.\s+(.+)', text)
    return insights[:4]


def run(question, sql, columns, rows, df, num_cols, cat_cols):
    """
    Returns dict: {explanation, insights, chart}
    """
    llm    = get_llm()
    prompt = explanation_prompt(question, sql, columns, rows)
    raw    = llm.generate(prompt, max_tokens=300)

    # split explanation from insights
    lines = [l.strip() for l in raw.strip().split("\n") if l.strip()]
    # first non-bullet line is the explanation
    explanation = next((l for l in lines if not l.startswith("-")), raw.strip())
    insights    = _parse_insights(raw)

    chart = _pick_chart(df, num_cols, cat_cols)

    return {"explanation": explanation, "insights": insights, "chart": chart}
