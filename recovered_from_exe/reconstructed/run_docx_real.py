from __future__ import annotations

import argparse
import csv
from pathlib import Path

from replace_engine import gerar_relatorio_dificuldade_por_secao, processar_docx


def output_base_dir() -> Path:
    # .../ReplaceDocx/recovered_from_exe/reconstructed/run_docx_real.py
    # -> pasta alvo: .../ReplaceDocx/saida_ok
    return Path(__file__).resolve().parents[2] / "saida_ok"


def default_output_path(input_docx: Path) -> Path:
    base = output_base_dir()
    base.mkdir(parents=True, exist_ok=True)
    return base / f"{input_docx.stem}_ok{input_docx.suffix}"


def default_report_paths(output_docx: Path) -> tuple[Path, Path]:
    csv_path = output_docx.with_name(f"{output_docx.stem}_relatorio_dificuldade.csv")
    txt_path = output_docx.with_name(f"{output_docx.stem}_relatorio_dificuldade.txt")
    return csv_path, txt_path


def save_difficulty_report(report: dict, output_docx: Path) -> tuple[Path, Path]:
    csv_path, txt_path = default_report_paths(output_docx)

    with csv_path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["conteudo", "secao", "facil", "media", "dificil", "total"])
        for row in report["secoes"]:
            writer.writerow(
                [
                    report["conteudo"],
                    row["secao"],
                    row["facil"],
                    row["media"],
                    row["dificil"],
                    row["total"],
                ]
            )
        totals = report["totais"]
        writer.writerow(
            [
                report["conteudo"],
                "TOTAL",
                totals["facil"],
                totals["media"],
                totals["dificil"],
                totals["total"],
            ]
        )

    lines = [
        f"Conteudo: {report['conteudo']}",
        f"Area: {report['area_conhecimento']}",
        f"Arquivo origem: {report['arquivo_origem']}",
        "",
        "Secao | Facil | Media | Dificil | Total",
    ]
    for row in report["secoes"]:
        lines.append(
            f"{row['secao']} | {row['facil']} | {row['media']} | {row['dificil']} | {row['total']}"
        )
    totals = report["totais"]
    lines.extend(
        [
            "",
            f"TOTAL | {totals['facil']} | {totals['media']} | {totals['dificil']} | {totals['total']}",
        ]
    )
    txt_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return csv_path, txt_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Processa um arquivo DOCX usando a engine recuperada do ReplaceDocx."
    )
    parser.add_argument(
        "input_docx",
        type=Path,
        help="Caminho do arquivo .docx de entrada",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Caminho do arquivo .docx de saída (opcional)",
    )
    parser.add_argument(
        "--finalize-word",
        action="store_true",
        help="Tenta rodar finalização de Word (no macOS/Linux vira fallback sem COM).",
    )
    parser.add_argument(
        "--font-name",
        default="Arial",
        help="Fonte principal (default: Arial)",
    )
    parser.add_argument(
        "--font-size",
        type=int,
        default=11,
        help="Tamanho da fonte (default: 11)",
    )
    parser.add_argument(
        "--badge-width-cm",
        type=float,
        default=1.3,
        help="Largura dos selos/imagens de marcador em cm (default: 1.3)",
    )
    parser.add_argument(
        "--column-width-cm",
        type=float,
        default=7.5,
        help="Largura alvo das imagens na finalização Word em cm (default: 7.5)",
    )
    parser.add_argument(
        "--area",
        default="biologia",
        help="Área do conhecimento para buscar as cápsulas em assets/areas/<area>/capsulas (default: biologia).",
    )
    parser.add_argument(
        "--section-banner-width-cm",
        type=float,
        default=None,
        help="Largura da arte de seção em cm (default: largura da coluna).",
    )
    parser.add_argument(
        "--no-section-banners",
        action="store_true",
        help="Desativa inserção automática de artes no início das seções.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    input_docx = args.input_docx.expanduser().resolve()
    if input_docx.suffix.lower() != ".docx":
        raise ValueError(f"Arquivo precisa ser .docx: {input_docx}")
    if not input_docx.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {input_docx}")

    output_docx = (
        args.output.expanduser().resolve()
        if args.output
        else default_output_path(input_docx)
    )

    cfg = {
        "area_conhecimento": args.area,
        "font_name": args.font_name,
        "font_size": args.font_size,
        "badge_width_cm": args.badge_width_cm,
        "column_width_cm": args.column_width_cm,
        "remove_gabarito": True,
        "justify": True,
        "finalize_word": bool(args.finalize_word),
        "force_inline_wrap": True,
        "insert_section_banners": not bool(args.no_section_banners),
    }
    if args.section_banner_width_cm is not None:
        cfg["section_banner_width_cm"] = args.section_banner_width_cm

    result = processar_docx(input_docx, output_docx, cfg)
    report = gerar_relatorio_dificuldade_por_secao(
        input_docx,
        area_conhecimento=args.area,
    )
    report_csv, report_txt = save_difficulty_report(report, output_docx)

    print("Processado com sucesso.")
    print("Área:", args.area)
    print("Entrada:", input_docx)
    print("Saída:", result["arquivo_saida"])
    print("Relatório CSV:", report_csv)
    print("Relatório TXT:", report_txt)


if __name__ == "__main__":
    main()
