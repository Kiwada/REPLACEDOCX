from __future__ import annotations

from contextlib import redirect_stderr, redirect_stdout
from datetime import datetime
from html import escape
from io import StringIO
from pathlib import Path


def output_base_dir() -> Path:
    # .../ReplaceDocx/recovered_from_exe/reconstructed/replace_engine_lib/report_outputs.py
    # -> pasta alvo: .../ReplaceDocx/saida_ok
    return Path(__file__).resolve().parents[3] / "saida_ok"


def default_output_path(input_docx: Path) -> Path:
    base = output_base_dir()
    base.mkdir(parents=True, exist_ok=True)
    return base / f"{input_docx.stem}_ok{input_docx.suffix}"


def default_report_paths(output_docx: Path) -> tuple[Path, Path]:
    html_path = output_docx.with_name(f"{output_docx.stem}_relatorio_dificuldade.html")
    pdf_path = output_docx.with_name(f"{output_docx.stem}_relatorio_dificuldade.pdf")
    return html_path, pdf_path


def _build_difficulty_report_html(report: dict) -> str:
    totals = report["totais"]
    max_total = max((row["total"] for row in report["secoes"]), default=1)

    row_html: list[str] = []
    for row in report["secoes"]:
        bar_width = int((row["total"] / max_total) * 100) if max_total else 0
        total_for_row = row["total"] if row["total"] > 0 else 1
        f_pct = int((row["facil"] / total_for_row) * 100)
        m_pct = int((row["media"] / total_for_row) * 100)
        d_pct = int((row["dificil"] / total_for_row) * 100)
        row_html.append(
            f"""
            <tr>
              <td class=\"secao\">{escape(row["secao"])}</td>
              <td class=\"n easy\">{row["facil"]}</td>
              <td class=\"n medium\">{row["media"]}</td>
              <td class=\"n hard\">{row["dificil"]}</td>
              <td class=\"n total\">{row["total"]}</td>
              <td>
                <div class=\"bar-bg\">
                  <div class=\"bar-fill\" style=\"width:{bar_width}%\"></div>
                </div>
                <div class=\"bar-split\">
                  <span class=\"easy\">{f_pct}%</span>
                  <span class=\"medium\">{m_pct}%</span>
                  <span class=\"hard\">{d_pct}%</span>
                </div>
              </td>
            </tr>
            """
        )

    return f"""<!doctype html>
<html lang=\"pt-BR\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>Relatório de Dificuldade - {escape(report["conteudo"])}</title>
  <style>
    @page {{
      size: A4;
      margin: 12mm;
    }}

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

    * {{
      box-sizing: border-box;
    }}

    body {{
      margin: 0;
      padding: 24px;
      background: var(--bg);
      color: var(--text);
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
      print-color-adjust: exact;
      -webkit-print-color-adjust: exact;
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

    @media print {{
      body {{
        padding: 0;
        background: #ffffff;
      }}

      .wrap {{
        max-width: none;
        margin: 0;
      }}
    }}
  </style>
</head>
<body>
  <div class=\"wrap\">
    <div class=\"header\">
      <h1 class=\"title\">Relatório de Dificuldade por Seção</h1>
      <div class=\"meta\">
        <div><strong>Conteúdo:</strong> {escape(report["conteudo"])}</div>
        <div><strong>Área:</strong> {escape(report["area_conhecimento"])}</div>
      </div>
    </div>

    <div class=\"cards\">
      <div class=\"card easy\"><div class=\"label\">Fácil</div><div class=\"value\">{totals["facil"]}</div></div>
      <div class=\"card medium\"><div class=\"label\">Média</div><div class=\"value\">{totals["media"]}</div></div>
      <div class=\"card hard\"><div class=\"label\">Difícil</div><div class=\"value\">{totals["dificil"]}</div></div>
      <div class=\"card total\"><div class=\"label\">Total</div><div class=\"value\">{totals["total"]}</div></div>
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
  </div>
</body>
</html>"""


def _save_difficulty_report_pdf_a4_from_html(html: str, pdf_path: Path, base_url: Path) -> bool:
    # WeasyPrint pode emitir warnings em stderr quando faltam libs nativas;
    # suprimimos ruído e ativamos fallback silenciosamente.
    with redirect_stdout(StringIO()), redirect_stderr(StringIO()):
        try:
            from weasyprint import HTML
        except Exception:
            return False

        try:
            HTML(string=html, base_url=str(base_url)).write_pdf(str(pdf_path))
            return True
        except Exception:
            return False


def _save_difficulty_report_pdf_a4_fallback_reportlab(report: dict, pdf_path: Path) -> bool:
    # Fallback para ambientes sem WeasyPrint.
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
        Paragraph(f"Arquivo origem: {escape(report['arquivo_origem'])}", meta_style),
        Paragraph(f"Gerado em: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", meta_style),
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
) -> tuple[Path, Path | None]:
    html_path, pdf_path = default_report_paths(output_docx)

    html = _build_difficulty_report_html(report)
    html_path.write_text(html, encoding="utf-8")

    pdf_out: Path | None = None
    if generate_pdf:
        # Prioridade: mesmo HTML/CSS -> PDF (visual igual ao relatório HTML).
        if _save_difficulty_report_pdf_a4_from_html(html, pdf_path, html_path.parent):
            pdf_out = pdf_path
        elif _save_difficulty_report_pdf_a4_fallback_reportlab(report, pdf_path):
            pdf_out = pdf_path

    return html_path, pdf_out
