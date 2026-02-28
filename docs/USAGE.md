# ContextoCLI Usage

## Setup

```bash
cd /Users/kiwada/Desktop/ReplaceDocx
python3 -m venv .venv
source .venv/bin/activate
pip install python-docx reportlab
pip install weasyprint
```

## Main commands

```bash
python contextocli.py process "/path/input.docx" --area biologia
python contextocli.py report "/path/input.docx" --area biologia
python contextocli.py check-assets --area biologia
```

## Output

Default output directory:

- `/Volumes/anthropic_externo/Documentos Processados CONTEXTO/`
- local fallback: `saida_ok/`

Generated files:

- `<name>_ok.docx`
- `<name>_ok_relatorio_dificuldade.html`
- `<name>_ok_relatorio_dificuldade.pdf`

## Useful flags

```bash
python contextocli.py process "/path/input.docx" --area biologia --no-format-text
python contextocli.py process "/path/input.docx" --area biologia --no-section-banners
python contextocli.py process "/path/input.docx" --area biologia --no-question-tables
python contextocli.py process "/path/input.docx" --area biologia --no-report-appendix
python contextocli.py process "/path/input.docx" --area biologia --no-report-pdf
python contextocli.py process "/path/input.docx" --area biologia --section-banner-width-cm 15.5
```
