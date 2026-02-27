from __future__ import annotations

from pathlib import Path

import win32com.client as win32


def cm_to_pt(cm: float) -> float:
    return cm * 28.3464567


def _get_alt_and_title(obj) -> tuple[str, str]:
    alt = ""
    title = ""

    try:
        alt = (obj.AlternativeText or "").strip()
    except Exception:
        pass

    try:
        title = (obj.Title or "").strip()
    except Exception:
        return alt, title

    return alt, title


def _is_badge(obj, badge_tag: str) -> bool:
    alt, title = _get_alt_and_title(obj)
    return (badge_tag in alt) or (badge_tag in title)


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
    docx_in = str(Path(docx_in).resolve())
    docx_out = str(Path(docx_out).resolve())

    word = win32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        doc = word.Documents.Open(docx_in)

        try:
            doc.Styles("Normal").Font.Name = font_name
            doc.Styles("Normal").Font.Size = font_size
        except Exception:
            pass

        try:
            doc.Content.Font.Name = font_name
            doc.Content.Font.Size = font_size
        except Exception:
            pass

        try:
            doc.Content.ParagraphFormat.Alignment = 3 if justify else 0
        except Exception:
            pass

        target_width_pt = cm_to_pt(float(column_width_cm))

        for shp in doc.InlineShapes:
            try:
                if _is_badge(shp, badge_tag):
                    continue
                shp.LockAspectRatio = True
                shp.Width = target_width_pt
            except Exception:
                continue

        for shp in doc.Shapes:
            try:
                if _is_badge(shp, badge_tag):
                    continue
                shp.LockAspectRatio = True
                shp.Width = target_width_pt
                if force_inline_wrap:
                    try:
                        shp.WrapFormat.Type = 7
                    except Exception:
                        continue
            except Exception:
                continue

        doc.SaveAs(docx_out)
        doc.Close(False)
    finally:
        word.Quit()
