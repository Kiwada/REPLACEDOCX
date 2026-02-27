from __future__ import annotations

import sys

if sys.platform.startswith("win"):
    try:
        from word_finalize_win_recovered import finalize_with_word
    except ImportError:  # pragma: no cover
        from .word_finalize_win_recovered import finalize_with_word  # type: ignore
else:
    try:
        from word_finalize_mac import finalize_with_word
    except ImportError:  # pragma: no cover
        from .word_finalize_mac import finalize_with_word  # type: ignore

