from __future__ import annotations

from pathlib import Path
import shutil
import warnings


def finalize_with_word(
    docx_in: str | Path,
    docx_out: str | Path,
    column_width_cm: float,
    font_name: str,
    font_size: int,
    justify: bool = True,
    force_inline_wrap: bool = True,
    badge_tag: str = "BADGE_REPLACE_DOCX",
) -> None:
    """
    Backend de compatibilidade para macOS/Linux.

    Não usa COM/Word. Mantém a assinatura do backend Windows para permitir
    rodar a engine no Mac sem quebrar. Se `docx_in != docx_out`, apenas copia.
    """
    src = Path(docx_in).resolve()
    dst = Path(docx_out).resolve()

    if src != dst:
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)

    warnings.warn(
        (
            "finalize_with_word (COM) nao esta disponivel neste sistema. "
            "Arquivo salvo somente via python-docx. "
            "Ajustes finais de largura/wrap em imagens podem ficar diferentes."
        ),
        RuntimeWarning,
        stacklevel=2,
    )

    # Mantemos as variáveis para deixar claro o contrato da função.
    _ = (
        column_width_cm,
        font_name,
        font_size,
        justify,
        force_inline_wrap,
        badge_tag,
    )

