"""Shared fixtures for docs tests.

Import strategy: three generate.py files with the same name live in different
skill directories. We use importlib to load them under distinct names.
"""

import importlib.util
import os
import sys

import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Put skills/docs-letter/ on sys.path so ``import generate`` resolves to
# skills/docs-letter/generate.py.
sys.path.insert(0, os.path.join(REPO_ROOT, "skills", "docs-letter"))

import generate as letter_generate  # noqa: E402


def _load_module(name: str, filepath: str):
    """Load a Python module from an explicit file path."""
    spec = importlib.util.spec_from_file_location(name, filepath)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


ord_generate = _load_module(
    "ord_generate",
    os.path.join(REPO_ROOT, "skills", "docs-ord", "generate.py"),
)
memo_generate = _load_module(
    "memo_generate",
    os.path.join(REPO_ROOT, "skills", "docs-memo", "generate.py"),
)
di_generate = _load_module(
    "di_generate",
    os.path.join(REPO_ROOT, "skills", "docs-di", "generate.py"),
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture()
def mock_org_config():
    """Realistic test config matching org_details.md.example structure."""
    return {
        "parent_org": "Тестовое агентство",
        "parent_org_short": "",
        "full_name": "Тестовое учреждение «Центр тестирования»",
        "short_name": "ТУ ЦТ",
        "address": "[адрес организации]",
        "phone_fax": "тел.: +7 (000) 000-00-00",
        "email": "user@example.org",
        "okpo": "00000000",
        "ogrn": "0000000000000",
        "inn_kpp": "0000000000/000000000",
        "leader_title": "Генеральный директор",
        "leader_name_nom": "Т.Т. Тестов",
        "leader_name_gen": "Т.Т. Тестова",
        "author_position": "Начальник отдела тестирования",
        "author_position_gen": "начальника отдела тестирования",
        "author_name_short": "Тестов Т.Т.",
        "author_name_full": "Тестов Тест Тестович",
        "author_phone": "+7 (000) 000-00-00",
        "approver_1_name": "Примеров П.П.",
        "approver_1_position": "Заместитель директора",
        "output_dir_ord": "/tmp",
        "output_dir_letter": "/tmp",
        "output_dir_memo": "/tmp",
        "output_dir_di": "/tmp",
    }


@pytest.fixture(autouse=True)
def patch_org_details(monkeypatch, mock_org_config):
    """Monkeypatch _load_org_details and log_generation in every generator module."""
    monkeypatch.setattr(ord_generate, "_load_org_details", lambda: mock_org_config)
    monkeypatch.setattr(letter_generate, "_load_org_details", lambda: mock_org_config)
    monkeypatch.setattr(memo_generate, "_load_org_details", lambda: mock_org_config)
    monkeypatch.setattr(di_generate, "_load_org_details", lambda: mock_org_config)
    ord_generate.ORG_CONFIGS.clear()
    _noop = lambda *a, **kw: None
    monkeypatch.setattr(ord_generate, "log_generation", _noop)
    monkeypatch.setattr(letter_generate, "log_generation", _noop)
    monkeypatch.setattr(memo_generate, "log_generation", _noop)
    monkeypatch.setattr(di_generate, "log_generation", _noop)
