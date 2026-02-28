"""Core package for ContextoCLI document processing."""

from .engine import gerar_relatorio_dificuldade_por_secao, processar_docx

__version__ = "1.2.0"
__all__ = ["gerar_relatorio_dificuldade_por_secao", "processar_docx"]
