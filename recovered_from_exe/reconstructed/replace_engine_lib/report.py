from __future__ import annotations

from pathlib import Path

from docx import Document

from .common import QUESTION_DIFFICULTY_RE, normalize_text_key
from .docx_utils import iter_paragraphs


def section_aliases_for_report() -> dict[str, list[str]]:
    return {
        "EXERCÍCIOS DE SALA": ["EXERCÍCIOS DE SALA"],
        "EXERCÍCIOS PROPOSTOS": ["EXERCÍCIOS PROPOSTOS"],
        "SEÇÃO ENEM": ["SEÇÃO ENEM"],
        "EXERCÍCIOS DE APROFUNDAMENTO": ["EXERCÍCIOS DE APROFUNDAMENTO"],
        "EXERCÍCIOS REGIONAIS": ["EXERCÍCIOS REGIONAIS"],
        "EXERCÍCIOS DISSERTATIVOS": ["EXERCÍCIOS DISSERTATIVOS", "EXERCÍCIO DISSERTATIVO"],
    }


def match_section_for_report(norm_txt: str) -> str | None:
    aliases = section_aliases_for_report()
    for canonical, raw_aliases in aliases.items():
        for alias in raw_aliases:
            norm_alias = normalize_text_key(alias)
            if norm_txt == norm_alias:
                return canonical
            if norm_txt.startswith(norm_alias + ":"):
                return canonical
            if norm_txt.startswith(norm_alias + " -"):
                return canonical
            if norm_alias in norm_txt and len(norm_txt) <= len(norm_alias) + 20:
                return canonical
    return None


def extract_difficulty(text: str) -> str | None:
    norm_txt = normalize_text_key(text)
    match = QUESTION_DIFFICULTY_RE.match(norm_txt)
    if not match:
        return None
    raw = match.group("level")
    if raw == "FACIL":
        return "facil"
    if raw == "MEDIA":
        return "media"
    if raw == "DIFICIL":
        return "dificil"
    return None


def extract_question_info(text: str) -> tuple[str | None, str] | None:
    norm_txt = normalize_text_key(text)
    match = QUESTION_DIFFICULTY_RE.match(norm_txt)
    if not match:
        return None
    num = match.group("num")
    level = match.group("level")
    if level == "FACIL":
        return num, "facil"
    if level == "MEDIA":
        return num, "media"
    if level == "DIFICIL":
        return num, "dificil"
    return None


def difficulty_label(diff_key: str) -> str:
    if diff_key == "facil":
        return "Fácil"
    if diff_key == "media":
        return "Média"
    if diff_key == "dificil":
        return "Difícil"
    return diff_key


def collect_questions_by_section(doc: Document) -> list[dict]:
    sections: list[dict] = []
    current: dict | None = None
    seq = 0

    for p in list(doc.paragraphs):
        txt = (p.text or "").strip()
        if not txt:
            continue

        norm_txt = normalize_text_key(txt)
        section = match_section_for_report(norm_txt)
        if section:
            current = {
                "secao": section,
                "questoes": [],
            }
            sections.append(current)
            seq = 0
            continue

        if current is None:
            continue

        info = extract_question_info(txt)
        if not info:
            continue

        num, diff = info
        seq += 1
        q_label = f"Questão {num}" if num else f"Questão {seq}"
        current["questoes"].append({"questao": q_label, "dificuldade": diff})

    return sections


def gerar_relatorio_dificuldade_por_secao(
    input_path: str | Path,
    area_conhecimento: str = "biologia",
) -> dict:
    input_path = Path(input_path)
    doc = Document(str(input_path))

    section_order = list(section_aliases_for_report().keys())
    section_counts: dict[str, dict[str, int]] = {
        section: {"facil": 0, "media": 0, "dificil": 0}
        for section in section_order
    }
    section_counts["SEM SEÇÃO"] = {"facil": 0, "media": 0, "dificil": 0}

    current_section = "SEM SEÇÃO"
    for p in iter_paragraphs(doc):
        txt = (p.text or "").strip()
        if not txt:
            continue

        norm_txt = normalize_text_key(txt)
        section = match_section_for_report(norm_txt)
        if section:
            current_section = section
            continue

        difficulty = extract_difficulty(txt)
        if not difficulty:
            continue

        section_counts[current_section][difficulty] += 1

    sections = []
    for section in section_order + ["SEM SEÇÃO"]:
        counts = section_counts[section]
        total = counts["facil"] + counts["media"] + counts["dificil"]
        if total == 0 and section == "SEM SEÇÃO":
            continue
        sections.append(
            {
                "secao": section,
                "facil": counts["facil"],
                "media": counts["media"],
                "dificil": counts["dificil"],
                "total": total,
            }
        )

    total_facil = sum(item["facil"] for item in sections)
    total_media = sum(item["media"] for item in sections)
    total_dificil = sum(item["dificil"] for item in sections)

    return {
        "conteudo": input_path.stem,
        "arquivo_origem": str(input_path),
        "area_conhecimento": area_conhecimento,
        "secoes": sections,
        "totais": {
            "facil": total_facil,
            "media": total_media,
            "dificil": total_dificil,
            "total": total_facil + total_media + total_dificil,
        },
    }
