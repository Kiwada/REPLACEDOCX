from __future__ import annotations

import sys
from pathlib import Path


def _bootstrap_src_path() -> None:
    root = Path(__file__).resolve().parent
    src = root / "src"
    src_str = str(src)
    if src_str not in sys.path:
        sys.path.insert(0, src_str)


_bootstrap_src_path()

from replacedocx.engine import (  # noqa: E402
    BADGE_TAG,
    ENGINE_LOG,
    GABARITO_RE,
    QUESTION_DIFFICULTY_RE,
    SECTION_TAG,
    gerar_relatorio_dificuldade_por_secao,
    processar_docx,
)

__all__ = [
    "BADGE_TAG",
    "ENGINE_LOG",
    "GABARITO_RE",
    "QUESTION_DIFFICULTY_RE",
    "SECTION_TAG",
    "gerar_relatorio_dificuldade_por_secao",
    "processar_docx",
]
