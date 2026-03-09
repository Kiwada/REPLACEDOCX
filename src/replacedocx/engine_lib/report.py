from __future__ import annotations

import re
from pathlib import Path

from docx import Document

from .common import QUESTION_DIFFICULTY_RE, normalize_text_key
from .docx_utils import iter_paragraphs

INLINE_QUESTION_DIFFICULTY_RE = re.compile(
    r"^\s*(?:(?P<num_a>\d+)\s*[\.\)\-:]?\s*\(?\s*(?P<level_a>FACIL|MEDIA|MEDIO|MEDIAS|DIFICIL)\s*\)?|"
    r"\(?\s*(?P<level_b>FACIL|MEDIA|MEDIO|MEDIAS|DIFICIL)\s*\)?\s*(?P<num_b>\d+)\s*[\.\)\-:]?)"
)
ANSWER_KEY_LINE_RE = re.compile(
    r"^\s*(?:QUESTAO\s*)?(?P<num>\d{1,3})\s*[:\)\.\-]?\s*(?:ALTERNATIVA\s*)?(?:\[|\()?\s*(?P<alt>[A-E])\s*(?:\]|\))?\s*$"
)
ANSWER_KEY_MULTI_PAIR_RE = re.compile(
    r"(?:QUESTAO\s*)?(?P<num>\d{1,3})\s*[:\)\.\-]\s*(?:ALTERNATIVA\s*)?(?:\[|\()?\s*(?P<alt>[A-E])\s*(?:\]|\))?"
    r"(?=\s*(?:$|[;,\|]|(?:\d{1,3}\s*[:\)\.\-])))"
)
ANSWER_KEY_INLINE_ALT_RE = re.compile(
    r"^\s*(?:GABARITO|RESPOSTA(?:\s+CORRETA)?)\s*[:\-]?\s*(?:ALTERNATIVA\s*)?(?:\[|\()?\s*(?P<alt>[A-E])\s*(?:\]|\))?\s*$"
)
QUESTION_LABEL_NUMBER_RE = re.compile(r"(\d+)\s*$")


def section_aliases_for_report() -> dict[str, list[str]]:
    return {
        "EXERCÍCIOS DE SALA": [
            "EXERCÍCIOS DE SALA",
            "EXERCÍCIO DE SALA",
            "QUESTÕES DE SALA",
            "QUESTÃO DE SALA",
        ],
        "EXERCÍCIOS PROPOSTOS": [
            "EXERCÍCIOS PROPOSTOS",
            "EXERCÍCIO PROPOSTO",
            "QUESTÕES PROPOSTAS",
            "QUESTÃO PROPOSTA",
            "PROPOSTOS",
        ],
        "SEÇÃO ENEM": ["SEÇÃO ENEM", "QUESTÕES ENEM", "QUESTÃO ENEM", "ENEM"],
        "EXERCÍCIOS DE APROFUNDAMENTO": [
            "EXERCÍCIOS DE APROFUNDAMENTO",
            "EXERCÍCIO DE APROFUNDAMENTO",
            "QUESTÕES DE APROFUNDAMENTO",
            "QUESTÃO DE APROFUNDAMENTO",
            "APROFUNDAMENTO",
        ],
        "EXERCÍCIOS REGIONAIS": [
            "EXERCÍCIOS REGIONAIS",
            "QUESTÕES REGIONAIS",
            "QUESTÃO REGIONAL",
            "EXERCÍCIO REGIONAL",
            "REGIONAIS",
        ],
        "EXERCÍCIOS DISSERTATIVOS": [
            "EXERCÍCIOS DISSERTATIVOS",
            "EXERCÍCIO DISSERTATIVO",
            "DISSERTATIVOS",
        ],
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
    raw = match.group("level") if match else None
    if raw is None:
        inline = INLINE_QUESTION_DIFFICULTY_RE.match(norm_txt)
        if inline:
            raw = inline.group("level_a") or inline.group("level_b")
    if raw is None:
        return None

    if raw == "FACIL":
        return "facil"
    if raw in {"MEDIA", "MEDIO", "MEDIAS"}:
        return "media"
    if raw == "DIFICIL":
        return "dificil"
    return None


def extract_question_info(text: str) -> tuple[str | None, str] | None:
    norm_txt = normalize_text_key(text)
    match = QUESTION_DIFFICULTY_RE.match(norm_txt)
    if match:
        num = match.group("num")
        level = match.group("level")
    else:
        inline = INLINE_QUESTION_DIFFICULTY_RE.match(norm_txt)
        if not inline:
            return None
        num = inline.group("num_a") or inline.group("num_b")
        level = inline.group("level_a") or inline.group("level_b")

    if level == "FACIL":
        return num, "facil"
    if level in {"MEDIA", "MEDIO", "MEDIAS"}:
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


def _is_answer_key_header(norm_txt: str) -> bool:
    return norm_txt.startswith("GABARITO") or norm_txt.startswith("RESPOSTA")


def _extract_answer_pairs(norm_txt: str) -> list[tuple[str, str]]:
    pairs: list[tuple[str, str]] = []
    if not norm_txt:
        return pairs

    line_match = ANSWER_KEY_LINE_RE.match(norm_txt)
    if line_match:
        pairs.append((str(int(line_match.group("num"))), line_match.group("alt").upper()))
        return pairs

    for match in ANSWER_KEY_MULTI_PAIR_RE.finditer(norm_txt):
        num = str(int(match.group("num")))
        alt = match.group("alt").upper()
        pairs.append((num, alt))
    return pairs


def _match_section_in_answer_key_context(norm_txt: str) -> str | None:
    matched = match_section_for_report(norm_txt)
    if matched:
        return matched

    aliases = section_aliases_for_report()
    for canonical, raw_aliases in aliases.items():
        for alias in raw_aliases:
            norm_alias = normalize_text_key(alias)
            if norm_alias and norm_alias in norm_txt:
                return canonical
    return None


def _extract_answer_keys_by_section(doc: Document) -> dict[str, dict[str, str]]:
    paragraphs = list(iter_paragraphs(doc))
    answers: dict[str, dict[str, str]] = {}
    current_section: str | None = None
    section_seq_counter: dict[str, int] = {}
    last_question_by_section: dict[str, str] = {}

    # Estratégia 1: blocos locais "Gabarito" dentro de cada seção.
    for i, p in enumerate(paragraphs):
        txt = (p.text or "").strip()
        if not txt:
            continue

        norm_txt = normalize_text_key(txt)
        section = _match_section_in_answer_key_context(norm_txt)
        if section:
            current_section = section
            section_seq_counter.setdefault(section, 0)
            continue

        if current_section:
            info = extract_question_info(txt)
            if info:
                qnum = info[0]
                if qnum is None:
                    section_seq_counter[current_section] = section_seq_counter.get(current_section, 0) + 1
                    qnum = str(section_seq_counter[current_section])
                else:
                    qnum = str(int(qnum))
                last_question_by_section[current_section] = qnum

        if not _is_answer_key_header(norm_txt):
            continue

        target_section = current_section or "SEM SEÇÃO"
        collected = 0
        inline_pairs = _extract_answer_pairs(norm_txt)
        if inline_pairs:
            for num, alt in inline_pairs:
                answers.setdefault(target_section, {})[num] = alt
            collected += len(inline_pairs)
        else:
            inline_alt = ANSWER_KEY_INLINE_ALT_RE.match(norm_txt)
            if inline_alt and current_section:
                last_num = last_question_by_section.get(current_section)
                if last_num:
                    answers.setdefault(current_section, {})[last_num] = inline_alt.group("alt").upper()
                    collected += 1

        j = i + 1
        while j < len(paragraphs):
            next_txt = (paragraphs[j].text or "").strip()
            if not next_txt:
                j += 1
                continue

            next_norm = normalize_text_key(next_txt)
            if _match_section_in_answer_key_context(next_norm):
                break

            pairs = _extract_answer_pairs(next_norm)
            if not pairs:
                if collected > 0:
                    break
                j += 1
                continue

            for num, alt in pairs:
                answers.setdefault(target_section, {})[num] = alt
            collected += len(pairs)
            j += 1

    # Estratégia 2: bloco final consolidado de gabarito com subseções.
    last_header_idx = -1
    for i, p in enumerate(paragraphs):
        norm_txt = normalize_text_key((p.text or "").strip())
        if norm_txt and _is_answer_key_header(norm_txt):
            last_header_idx = i

    if last_header_idx >= 0:
        target_section = "SEM SEÇÃO"
        for p in paragraphs[last_header_idx + 1 :]:
            txt = (p.text or "").strip()
            if not txt:
                continue

            norm_txt = normalize_text_key(txt)
            section = _match_section_in_answer_key_context(norm_txt)
            if section:
                target_section = section
                continue

            pairs = _extract_answer_pairs(norm_txt)
            if not pairs:
                continue

            for num, alt in pairs:
                answers.setdefault(target_section, {})[num] = alt

    # Estratégia 3: bloco final de gabarito por seção sem header explícito
    # (ex.: "EXERCÍCIOS DE SALA" seguido de "1. B", "2. A"...).
    tail_start = int(len(paragraphs) * 0.55)
    target_section = "SEM SEÇÃO"
    found_pairs_in_tail = 0
    for p in paragraphs[tail_start:]:
        txt = (p.text or "").strip()
        if not txt:
            continue

        norm_txt = normalize_text_key(txt)
        section = _match_section_in_answer_key_context(norm_txt)
        if section:
            target_section = section
            continue

        pairs = _extract_answer_pairs(norm_txt)
        if not pairs:
            continue

        for num, alt in pairs:
            answers.setdefault(target_section, {})[num] = alt
            found_pairs_in_tail += 1

    _ = found_pairs_in_tail

    return answers


def _question_number_from_label(label: str, fallback_idx: int) -> str:
    match = QUESTION_LABEL_NUMBER_RE.search(label or "")
    if not match:
        return str(fallback_idx)
    return str(int(match.group(1)))


def collect_questions_by_section(doc: Document, include_answer_key: bool = False) -> list[dict]:
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

    if include_answer_key and sections:
        answers_by_section = _extract_answer_keys_by_section(doc)
        fallback_answers = answers_by_section.get("SEM SEÇÃO", {})
        for section_item in sections:
            section_answers = answers_by_section.get(section_item["secao"], {})
            for idx, question in enumerate(section_item.get("questoes", []), start=1):
                qnum = _question_number_from_label(question.get("questao", ""), idx)
                answer = section_answers.get(qnum) or fallback_answers.get(qnum)
                if answer:
                    question["gabarito"] = answer

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
