"""pytest 共通設定。tests/ からプロジェクトルートの court_format_converter を import 可能にし、
examples/ の絶対パスを fixture で提供する。"""

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


@pytest.fixture
def examples_dir():
    return ROOT / "examples"
