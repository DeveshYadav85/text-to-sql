"""
tests/test_offline.py
Offline sanity checks (no LLM needed).
Kept for backwards compatibility — the main tests are in test_validator.py,
test_rag.py, test_pipeline.py, test_gold_queries.py.
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from agents.executor_agent import run as exec_sql
from core.validator import validate

SCHEMA = {
    "students":    ["student_id", "name", "dept_id", "year", "gender", "email", "status"],
    "departments": ["dept_id", "dept_name", "dept_code", "hod_name", "established_year"],
}


def test_basic_select():
    sql = "SELECT name FROM students LIMIT 5;"
    result = exec_sql(sql)
    assert result["error"] is None
    assert result["row_count"] == 5


def test_validate_good():
    sql = "SELECT name FROM students;"
    assert validate(sql, SCHEMA) == sql


def test_validate_bad():
    import pytest
    with pytest.raises(ValueError):
        validate("DELETE FROM students;", SCHEMA)
