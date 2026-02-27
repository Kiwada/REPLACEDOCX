from __future__ import annotations

import argparse
import csv
from datetime import datetime
from html import escape
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


def default_report_paths(output_docx: Path) -> tuple[Path, Path, Path, Path]:
    csv_path = output_docx.with_name(f"{output_docx.stem}_relatorio_dificuldade.csv")
    txt_path = output_docx.with_name(f"{output_docx.stem}_relatorio_dificuldade.txt")
    html_path = output_docx.with_name(f"{output_docx.stem}_relatorio_dificuldade.html")
    pdf_path = output_docx.with_name(f"{output_docx.stem}_relatorio_dificuldade.pdf")
    return csv_path, txt_path, html_path, pdf_path


def _save_difficulty_report_pdf_a4(report: dict, pdf_path: Path) -> bool:
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    except Exception:
        return False

    totals = report["totais"]
    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=A4,
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36,
        title=f"Relatório de Dificuldade - {report['conteudo']}",
    )
    width, _ = A4
    usable_width = width - 72
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleMain",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=15,
        leading=18,
        spaceAfter=8,
        alignment=1,
        textColor=colors.HexColor("#111827"),
    )
    meta_style = ParagraphStyle(
        "Meta",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=11,
        textColor=colors.HexColor("#4B5563"),
    )
    story = [
        Paragraph("Relatório de Dificuldade por Seção", title_style),
        Paragraph(f"Conteúdo: {escape(report['conteudo'])}", meta_style),
        Paragraph(f"Área: {escape(report['area_conhecimento'])}", meta_style),
        Paragraph(
            f"Arquivo origem: {escape(report['arquivo_origem'])}",
            meta_style,
        ),
        Paragraph(
            f"Gerado em: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            meta_style,
        ),
        Spacer(1, 10),
    ]

    cards_data = [
        ["Fácil", "Média", "Difícil", "Total"],
        [
            str(totals["facil"]),
            str(totals["media"]),
            str(totals["dificil"]),
            str(totals["total"]),
        ],
    ]
    card_table = Table(cards_data, colWidths=[usable_width / 4] * 4)
    card_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#DCFCE7")),
                ("BACKGROUND", (1, 0), (1, 0), colors.HexColor("#FEF3C7")),
                ("BACKGROUND", (2, 0), (2, 0), colors.HexColor("#FEE2E2")),
                ("BACKGROUND", (3, 0), (3, 0), colors.HexColor("#E5E7EB")),
                ("TEXTCOLOR", (0, 1), (0, 1), colors.HexColor("#166534")),
                ("TEXTCOLOR", (1, 1), (1, 1), colors.HexColor("#92400E")),
                ("TEXTCOLOR", (2, 1), (2, 1), colors.HexColor("#991B1B")),
                ("TEXTCOLOR", (3, 1), (3, 1), colors.HexColor("#111827")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D1D5DB")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    story.extend([card_table, Spacer(1, 10)])

    table_data = [["Seção", "Fácil", "Média", "Difícil", "Total"]]
    for row in report["secoes"]:
        table_data.append(
            [
                row["secao"],
                str(row["facil"]),
                str(row["media"]),
                str(row["dificil"]),
                str(row["total"]),
            ]
        )
    table_data.append(
        [
            "TOTAL",
            str(totals["facil"]),
            str(totals["media"]),
            str(totals["dificil"]),
            str(totals["total"]),
        ]
    )

    section_table = Table(
        table_data,
        colWidths=[
            usable_width * 0.52,
            usable_width * 0.12,
            usable_width * 0.12,
            usable_width * 0.12,
            usable_width * 0.12,
        ],
        repeatRows=1,
    )
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F2937")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D1D5DB")),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("TEXTCOLOR", (1, 1), (1, -2), colors.HexColor("#166534")),
        ("TEXTCOLOR", (2, 1), (2, -2), colors.HexColor("#92400E")),
        ("TEXTCOLOR", (3, 1), (3, -2), colors.HexColor("#991B1B")),
        ("TEXTCOLOR", (4, 1), (4, -2), colors.HexColor("#111827")),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica-Bold"),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E5E7EB")),
        ("TEXTCOLOR", (0, -1), (-1, -1), colors.HexColor("#111827")),
    ]
    for r in range(1, len(table_data) - 1):
        if r % 2 == 0:
            style.append(("BACKGROUND", (0, r), (-1, r), colors.HexColor("#F8FAFC")))
    section_table.setStyle(TableStyle(style))
    story.append(section_table)

    doc.build(story)
    return True


def save_difficulty_report(
    report: dict,
    output_docx: Path,
    *,
    generate_pdf: bool = True,
) -> tuple[Path, Path, Path, Path | None]:
    csv_path, txt_path, html_path, pdf_path = default_report_paths(output_docx)

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

    max_total = max((row["total"] for row in report["secoes"]), default=1)
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row_html = []
    for row in report["secoes"]:
        bar_width = int((row["total"] / max_total) * 100) if max_total else 0
        total_for_row = row["total"] if row["total"] > 0 else 1
        f_pct = int((row["facil"] / total_for_row) * 100)
        m_pct = int((row["media"] / total_for_row) * 100)
        d_pct = int((row["dificil"] / total_for_row) * 100)
        row_html.append(
            f"""
            <tr>
              <td class="secao">{escape(row["secao"])}</td>
              <td class="n easy">{row["facil"]}</td>
              <td class="n medium">{row["media"]}</td>
              <td class="n hard">{row["dificil"]}</td>
              <td class="n total">{row["total"]}</td>
              <td>
                <div class="bar-bg">
                  <div class="bar-fill" style="width:{bar_width}%"></div>
                </div>
                <div class="bar-split">
                  <span class="easy">{f_pct}%</span>
                  <span class="medium">{m_pct}%</span>
                  <span class="hard">{d_pct}%</span>
                </div>
              </td>
            </tr>
            """
        )

    html = f"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Relatório de Dificuldade - {escape(report["conteudo"])}</title>
  <style>
    :root {{
      --bg: #f6f8fb;
      --card: #ffffff;
      --line: #dde3ec;
      --text: #1f2937;
      --muted: #617287;
      --easy: #0f9d58;
      --medium: #f39c12;
      --hard: #d64545;
      --total: #374151;
      --bar: #2c6fbb;
    }}
    body {{
      margin: 0;
      padding: 24px;
      background: var(--bg);
      color: var(--text);
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
    }}
    .wrap {{
      max-width: 1100px;
      margin: 0 auto;
    }}
    .header {{
      margin-bottom: 16px;
    }}
    .title {{
      margin: 0;
      font-size: 28px;
      line-height: 1.2;
    }}
    .meta {{
      margin-top: 8px;
      color: var(--muted);
      font-size: 14px;
    }}
    .cards {{
      display: grid;
      grid-template-columns: repeat(4, minmax(130px, 1fr));
      gap: 12px;
      margin: 16px 0 20px;
    }}
    .card {{
      background: var(--card);
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 12px 14px;
    }}
    .card .label {{
      font-size: 12px;
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: .04em;
    }}
    .card .value {{
      margin-top: 6px;
      font-size: 26px;
      font-weight: 700;
    }}
    .easy .value {{ color: var(--easy); }}
    .medium .value {{ color: var(--medium); }}
    .hard .value {{ color: var(--hard); }}
    .total .value {{ color: var(--total); }}
    table {{
      width: 100%;
      border-collapse: collapse;
      background: var(--card);
      border: 1px solid var(--line);
      border-radius: 12px;
      overflow: hidden;
    }}
    th, td {{
      padding: 10px 12px;
      border-bottom: 1px solid var(--line);
      vertical-align: middle;
    }}
    th {{
      text-align: left;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: .04em;
      color: var(--muted);
      background: #f3f6fa;
    }}
    td.n {{
      text-align: center;
      font-weight: 700;
    }}
    td.easy {{ color: var(--easy); }}
    td.medium {{ color: var(--medium); }}
    td.hard {{ color: var(--hard); }}
    td.total {{ color: var(--total); }}
    td.secao {{
      font-weight: 600;
      width: 33%;
    }}
    .bar-bg {{
      width: 100%;
      height: 8px;
      border-radius: 99px;
      background: #e6ecf4;
      overflow: hidden;
      margin-bottom: 6px;
    }}
    .bar-fill {{
      height: 100%;
      background: linear-gradient(90deg, #4f9de8, var(--bar));
    }}
    .bar-split {{
      display: flex;
      gap: 10px;
      font-size: 12px;
      font-weight: 600;
    }}
    .bar-split .easy {{ color: var(--easy); }}
    .bar-split .medium {{ color: var(--medium); }}
    .bar-split .hard {{ color: var(--hard); }}
    .footer {{
      margin-top: 12px;
      color: var(--muted);
      font-size: 12px;
    }}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="header">
      <h1 class="title">Relatório de Dificuldade por Seção</h1>
      <div class="meta">
        <div><strong>Conteúdo:</strong> {escape(report["conteudo"])}</div>
        <div><strong>Área:</strong> {escape(report["area_conhecimento"])}</div>
        <div><strong>Arquivo origem:</strong> {escape(report["arquivo_origem"])}</div>
        <div><strong>Gerado em:</strong> {generated_at}</div>
      </div>
    </div>

    <div class="cards">
      <div class="card easy"><div class="label">Fácil</div><div class="value">{totals["facil"]}</div></div>
      <div class="card medium"><div class="label">Média</div><div class="value">{totals["media"]}</div></div>
      <div class="card hard"><div class="label">Difícil</div><div class="value">{totals["dificil"]}</div></div>
      <div class="card total"><div class="label">Total</div><div class="value">{totals["total"]}</div></div>
    </div>

    <table>
      <thead>
        <tr>
          <th>Seção</th>
          <th>Fácil</th>
          <th>Média</th>
          <th>Difícil</th>
          <th>Total</th>
          <th>Visual</th>
        </tr>
      </thead>
      <tbody>
        {''.join(row_html)}
      </tbody>
    </table>
    <div class="footer">Arquivo gerado automaticamente pelo ReplaceDocx.</div>
  </div>
</body>
</html>"""
    html_path.write_text(html, encoding="utf-8")

    pdf_out: Path | None = None
    if generate_pdf and _save_difficulty_report_pdf_a4(report, pdf_path):
        pdf_out = pdf_path

    return csv_path, txt_path, html_path, pdf_out


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
    parser.add_argument(
        "--no-question-tables",
        action="store_true",
        help="Desativa inserção das tabelas de autoavaliação por questão ao final das seções.",
    )
    parser.add_argument(
        "--no-report-appendix",
        action="store_true",
        help="Desativa anexo do relatório de dificuldade em A4 no final do DOCX.",
    )
    parser.add_argument(
        "--no-report-pdf",
        action="store_true",
        help="Desativa geração do relatório de dificuldade em PDF A4.",
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
        "insert_question_tables": not bool(args.no_question_tables),
    }
    if args.section_banner_width_cm is not None:
        cfg["section_banner_width_cm"] = args.section_banner_width_cm

    report = gerar_relatorio_dificuldade_por_secao(
        input_docx,
        area_conhecimento=args.area,
    )
    cfg["append_difficulty_report"] = not bool(args.no_report_appendix)
    cfg["difficulty_report_data"] = report

    result = processar_docx(input_docx, output_docx, cfg)
    report_csv, report_txt, report_html, report_pdf = save_difficulty_report(
        report,
        output_docx,
        generate_pdf=not bool(args.no_report_pdf),
    )

    print("Processado com sucesso.")
    print("Área:", args.area)
    print("Entrada:", input_docx)
    print("Saída:", result["arquivo_saida"])
    print("Relatório CSV:", report_csv)
    print("Relatório TXT:", report_txt)
    print("Relatório HTML:", report_html)
    if report_pdf:
        print("Relatório PDF (A4):", report_pdf)
    elif not args.no_report_pdf:
        print("Relatório PDF (A4): não gerado (instale reportlab na venv).")


if __name__ == "__main__":
    main()
