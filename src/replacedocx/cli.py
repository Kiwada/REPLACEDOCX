from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from .engine import gerar_relatorio_dificuldade_por_secao, processar_docx
from .engine_lib.common import (
    canonicalize_area,
    default_markers_for_area,
    default_section_banners_for_area,
    resolve_path,
    runtime_dir,
)
from .engine_lib.report_outputs import (
    default_output_path,
    force_output_in_base,
    save_difficulty_report,
)

CLI_NAME = "ContextoCLI"
CLI_VERSION = "1.2.0"
CLI_PROG = "contextocli"
SUBCOMMANDS = {"process", "report", "check-assets", "version"}


def _use_ansi_colors() -> bool:
    if os.environ.get("NO_COLOR"):
        return False
    term = (os.environ.get("TERM") or "").lower()
    return sys.stdout.isatty() and term not in {"", "dumb"}


def _ansi(text: str, code: str) -> str:
    if not _use_ansi_colors():
        return text
    return f"\033[{code}m{text}\033[0m"


def _print_contextocli_banner() -> None:
    lines = [
        "╔══════════════════════════════════════════════╗",
        "║                  CONTEXTOCLI                 ║",
        "║      Pipeline DOCX para conteúdo escolar     ║",
        "╚══════════════════════════════════════════════╝",
    ]
    print(_ansi(lines[0], "36;1"))
    print(_ansi(lines[1], "36;1"))
    print(_ansi(lines[2], "90"))
    print(_ansi(lines[3], "36;1"))


def _normalize_cli_path(path_like: Path | str) -> Path:
    raw = str(path_like).strip()
    # Tolera casos em que o usuário envolve o caminho com aspas extras,
    # ex.: "'/Volumes/.../arquivo.docx'" (aspas simples literais).
    for _ in range(2):
        if len(raw) >= 2 and ((raw[0] == raw[-1] == "'") or (raw[0] == raw[-1] == '"')):
            raw = raw[1:-1].strip()
            continue
        break
    return Path(raw).expanduser().resolve()


def _validate_input_docx(path: Path) -> Path:
    input_docx = _normalize_cli_path(path)
    if input_docx.suffix.lower() != ".docx":
        raise ValueError(f"Arquivo precisa ser .docx: {input_docx}")
    if not input_docx.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {input_docx}")
    return input_docx


def _build_common_options(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--area",
        default="biologia",
        help="Área do conhecimento (biologia, quimica, fisica, linguagens, historia, filosofia, sociologia). Também aceita acentos e variações.",
    )


def _configure_process_parser(sub: argparse._SubParsersAction) -> None:
    p = sub.add_parser("process", help="Processa DOCX completo: banners, cápsulas, autoavaliação e relatórios.")
    p.add_argument("input_docx", type=Path, help="Caminho do arquivo .docx de entrada")
    p.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Nome/caminho do .docx de saída (o arquivo será salvo sempre em /Volumes/anthropic_externo/Documentos Processados CONTEXTO).",
    )
    p.add_argument(
        "--finalize-word",
        action="store_true",
        help="Tenta rodar finalização de Word (no macOS/Linux vira fallback sem COM).",
    )
    p.add_argument("--font-name", default="Arial", help="Fonte principal (default: Arial)")
    p.add_argument("--font-size", type=int, default=11, help="Tamanho da fonte (default: 11)")
    p.add_argument(
        "--no-format-text",
        action="store_true",
        help="Não aplica formatação global de fonte/parágrafo no texto original.",
    )
    p.add_argument(
        "--badge-width-cm",
        type=float,
        default=1.3,
        help="Largura dos selos/imagens de marcador em cm (default: 1.3)",
    )
    p.add_argument(
        "--column-width-cm",
        type=float,
        default=7.5,
        help="Largura alvo das imagens/tabelas em cm (default: 7.5)",
    )
    _build_common_options(p)
    p.add_argument(
        "--section-banner-width-cm",
        type=float,
        default=None,
        help="Largura da arte de seção em cm (default: largura da coluna).",
    )
    p.add_argument(
        "--no-section-banners",
        action="store_true",
        help="Desativa inserção automática de artes no início das seções.",
    )
    p.add_argument(
        "--no-question-tables",
        action="store_true",
        help="Desativa inserção das tabelas de autoavaliação por questão ao final das seções.",
    )
    p.add_argument(
        "--preserve-paragraphs",
        action="store_true",
        help="Preserva todos os parágrafos originais, sem remover linhas de gabarito/apêndice final.",
    )
    p.add_argument(
        "--with-report-appendix",
        action="store_true",
        help="Anexa o relatório de dificuldade em A4 no final do DOCX (desativado por padrão).",
    )
    p.add_argument(
        "--no-report-appendix",
        action="store_false",
        dest="with_report_appendix",
        help=argparse.SUPPRESS,
    )
    p.add_argument(
        "--no-report-pdf",
        action="store_true",
        help="Desativa geração do relatório de dificuldade em PDF A4.",
    )
    p.set_defaults(with_report_appendix=False)
    p.set_defaults(handler=_cmd_process)


def _configure_report_parser(sub: argparse._SubParsersAction) -> None:
    p = sub.add_parser("report", help="Gera apenas os relatórios de dificuldade (HTML/PDF).")
    p.add_argument("input_docx", type=Path, help="Caminho do arquivo .docx de entrada")
    p.add_argument(
        "--output-docx",
        type=Path,
        default=None,
        help="Arquivo .docx base para nomear relatórios (sempre salvo em /Volumes/anthropic_externo/Documentos Processados CONTEXTO).",
    )
    _build_common_options(p)
    p.add_argument(
        "--no-pdf",
        action="store_true",
        help="Desativa geração do PDF A4 do relatório.",
    )
    p.set_defaults(handler=_cmd_report)


def _configure_check_assets_parser(sub: argparse._SubParsersAction) -> None:
    p = sub.add_parser("check-assets", help="Valida se os assets da área existem no local esperado.")
    _build_common_options(p)
    p.set_defaults(handler=_cmd_check_assets)


def _configure_version_parser(sub: argparse._SubParsersAction) -> None:
    p = sub.add_parser("version", help="Exibe a versão da CLI.")
    p.set_defaults(handler=_cmd_version)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=CLI_PROG,
        description=f"{CLI_NAME}: processamento de DOCX e geração de relatórios.",
    )
    sub = parser.add_subparsers(dest="command", required=True)
    _configure_process_parser(sub)
    _configure_report_parser(sub)
    _configure_check_assets_parser(sub)
    _configure_version_parser(sub)
    return parser


def _cmd_process(args: argparse.Namespace) -> int:
    input_docx = _validate_input_docx(args.input_docx)
    area = canonicalize_area(args.area)
    format_text_enabled = not bool(args.no_format_text)
    if area == "matematica":
        format_text_enabled = False
    output_docx = (
        force_output_in_base(_normalize_cli_path(args.output))
        if args.output
        else default_output_path(input_docx)
    )

    cfg = {
        "area_conhecimento": area,
        "font_name": args.font_name,
        "font_size": args.font_size,
        "format_text": format_text_enabled,
        "badge_width_cm": args.badge_width_cm,
        "column_width_cm": args.column_width_cm,
        "remove_gabarito": True,
        "preserve_paragraphs": bool(args.preserve_paragraphs),
        "justify": True,
        "finalize_word": bool(args.finalize_word),
        "force_inline_wrap": True,
        "insert_section_banners": not bool(args.no_section_banners),
        "insert_question_tables": not bool(args.no_question_tables),
    }
    if args.section_banner_width_cm is not None:
        cfg["section_banner_width_cm"] = args.section_banner_width_cm

    report = gerar_relatorio_dificuldade_por_secao(
        input_docx,
        area_conhecimento=area,
    )
    cfg["append_difficulty_report"] = bool(args.with_report_appendix)
    cfg["difficulty_report_data"] = report

    result = processar_docx(input_docx, output_docx, cfg)
    report_html, report_pdf = save_difficulty_report(
        report,
        output_docx,
        generate_pdf=not bool(args.no_report_pdf),
    )

    print("Processado com sucesso.")
    print("Área:", area)
    print("Entrada:", input_docx)
    print("Saída:", result["arquivo_saida"])
    print("Relatório HTML:", report_html)
    if report_pdf:
        print("Relatório PDF (A4):", report_pdf)
    elif not args.no_report_pdf:
        print("Relatório PDF (A4): não gerado (instale weasyprint; fallback: reportlab).")
    return 0


def _cmd_report(args: argparse.Namespace) -> int:
    input_docx = _validate_input_docx(args.input_docx)
    area = canonicalize_area(args.area)
    output_docx = (
        force_output_in_base(_normalize_cli_path(args.output_docx))
        if args.output_docx is not None
        else default_output_path(input_docx)
    )
    if output_docx.suffix.lower() != ".docx":
        raise ValueError(f"--output-docx precisa terminar com .docx: {output_docx}")

    report = gerar_relatorio_dificuldade_por_secao(
        input_docx,
        area_conhecimento=area,
    )
    report_html, report_pdf = save_difficulty_report(
        report,
        output_docx,
        generate_pdf=not bool(args.no_pdf),
    )

    print("Relatórios gerados com sucesso.")
    print("Área:", area)
    print("Entrada:", input_docx)
    print("Base de saída:", output_docx)
    print("Relatório HTML:", report_html)
    if report_pdf:
        print("Relatório PDF (A4):", report_pdf)
    elif not args.no_pdf:
        print("Relatório PDF (A4): não gerado (instale weasyprint; fallback: reportlab).")
    return 0


def _cmd_check_assets(args: argparse.Namespace) -> int:
    area = canonicalize_area(args.area)

    marker_paths = sorted(set(default_markers_for_area(area).values()))
    section_paths = sorted(set(default_section_banners_for_area(area).values()))

    required = [("capsula", p) for p in marker_paths] + [("secao", p) for p in section_paths]

    missing = []
    found = []
    for kind, rel_path in required:
        resolved = resolve_path(rel_path)
        if resolved.exists():
            found.append((kind, rel_path, resolved))
        else:
            expected = runtime_dir() / "Assets" / rel_path
            missing.append((kind, rel_path, expected))

    print(f"Área: {area}")
    print(f"Encontrados: {len(found)}")
    for kind, rel_path, resolved in found:
        print(f"  [OK] {kind:<7} {rel_path} -> {resolved}")

    if missing:
        print(f"Ausentes: {len(missing)}")
        for kind, rel_path, resolved in missing:
            print(f"  [MISS] {kind:<7} {rel_path} -> esperado em {resolved}")
        return 1

    print("Todos os assets obrigatórios foram encontrados.")
    return 0


def _cmd_version(_: argparse.Namespace) -> int:
    _print_contextocli_banner()
    print(f"{CLI_NAME} {CLI_VERSION}")
    return 0


def main(argv: list[str] | None = None) -> int:
    args_in = list(argv if argv is not None else sys.argv[1:])

    # Compatibilidade: permite chamada legada sem subcomando.
    # Exemplo antigo: replacedocx <arquivo.docx> --area biologia
    if args_in and args_in[0] not in SUBCOMMANDS and args_in[0] not in {"-h", "--help"}:
        args_in = ["process", *args_in]

    parser = build_parser()
    parsed = parser.parse_args(args_in)
    return int(parsed.handler(parsed))


if __name__ == "__main__":
    raise SystemExit(main())
