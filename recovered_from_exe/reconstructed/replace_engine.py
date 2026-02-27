from __future__ import annotations

"""Fachada pública da engine.

Este arquivo mantém compatibilidade com imports existentes enquanto a
implementação fica modularizada em `replace_engine_lib/`.
"""

from replace_engine_lib.common import (
    BADGE_TAG,
    ENGINE_LOG,
    GABARITO_RE,
    QUESTION_DIFFICULTY_RE,
    SECTION_TAG,
)
from replace_engine_lib.pipeline import processar_docx
from replace_engine_lib.report import gerar_relatorio_dificuldade_por_secao

__all__ = [
    "BADGE_TAG",
    "ENGINE_LOG",
    "GABARITO_RE",
    "QUESTION_DIFFICULTY_RE",
    "SECTION_TAG",
    "gerar_relatorio_dificuldade_por_secao",
    "processar_docx",
]
