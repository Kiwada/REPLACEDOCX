# Port de `ReplaceDocx` para macOS (base recuperada do `.exe`)

## O que foi recuperado

- `replace_engine.py` (reconstruido a partir da disassembly)
- `word_finalize` Windows (COM/Word)
- `word_finalize` compativel com macOS/Linux (fallback sem COM)

Arquivos nesta pasta:

- `replace_engine.py`
- `word_finalize.py` (seleciona backend por sistema)
- `word_finalize_win_recovered.py`
- `word_finalize_mac.py`

## O que quebra no macOS (e por quê)

O executavel original usa `win32com` + `Word.Application` para a etapa final:

- ajustar largura de imagens (`InlineShapes` / `Shapes`)
- forcar `WrapFormat.Type`
- salvar novamente no Word

Isso é tecnologia COM e so existe no Windows.

## Como testar no macOS (sem UI)

1. Instale Python 3 e dependencias:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install python-docx reportlab
```

2. Rode a engine com um script simples:

```python
from replace_engine import processar_docx

cfg = {
    "area_conhecimento": "biologia",
    "font_name": "Arial",
    "font_size": 11,
    "badge_width_cm": 1.3,
    "column_width_cm": 7.5,
    "remove_gabarito": True,
    "justify": True,
    "finalize_word": False,  # recomendado no macOS
    "force_inline_wrap": True,
}

print(processar_docx("entrada.docx", "saida.docx", cfg))
```

3. Coloque as imagens:

- em `assets/areas/biologia/capsulas/` (recomendado), ou
- em `~/Library/Application Support/ReplaceDocx/assets/areas/biologia/capsulas/`

Estrutura:

```text
assets/areas/<area>/capsulas/facil.png
assets/areas/<area>/capsulas/media.png
assets/areas/<area>/capsulas/dificil.png
assets/areas/<area>/secoes/exercicios_sala.png
assets/areas/<area>/secoes/exercicios_propostos.png
assets/areas/<area>/secoes/secao_enem.png
assets/areas/<area>/secoes/exercicios_aprofundamento.png
assets/areas/<area>/secoes/exercicios_regionais.png
assets/areas/<area>/secoes/exercicios_dissertativos.png
```

Exemplo de CLI:

```bash
python contextocli.py process "/caminho/arquivo.docx" --area biologia
python contextocli.py report "/caminho/arquivo.docx" --area biologia
python contextocli.py check-assets --area biologia
```

Compatibilidade:
- `python replacedocx_cli.py ...` continua válido como alias.
- `python run_docx_real.py ...` continua válido como alias da nova CLI.

Saída padrão:

- pasta: `/Volumes/anthropic_externo/Documentos Processados CONTEXTO/` (fallback: `ReplaceDocx/saida_ok/`)
- nome: `<nome_original>_ok.docx`
- relatório HTML: `<nome_original>_ok_relatorio_dificuldade.html`
- relatório PDF (A4): `<nome_original>_ok_relatorio_dificuldade.pdf` (prioriza `weasyprint`; fallback `reportlab`)
- anexo A4 no próprio DOCX final com resumo por seção

Controle de artes de seção:

```bash
python contextocli.py process "/caminho/arquivo.docx" --area biologia --section-banner-width-cm 15.5
python contextocli.py process "/caminho/arquivo.docx" --area biologia --no-section-banners
python contextocli.py process "/caminho/arquivo.docx" --area biologia --no-format-text
```

Tabelas de dificuldade por questão (autoavaliação):

- Inseridas automaticamente no final do documento.
- Mantêm a ordem das seções encontradas no conteúdo.
- Colunas: `Questão | Dificuldade | Acertou | Errou | Revisar`.
- Para desativar:

```bash
python contextocli.py process "/caminho/arquivo.docx" --area biologia --no-question-tables
```

Relatório final no DOCX/PDF:

- anexo A4 no final do DOCX: ativo por padrão
- para desativar anexo A4 no DOCX:

```bash
python contextocli.py process "/caminho/arquivo.docx" --area biologia --no-report-appendix
```

- para desativar somente o PDF:

```bash
python contextocli.py process "/caminho/arquivo.docx" --area biologia --no-report-pdf
```

## Sobre o `finalize_word` no macOS

- Se `finalize_word=False`: roda so com `python-docx` (mais estavel no Mac)
- Se `finalize_word=True`: o backend `word_finalize_mac.py` nao usa COM; ele apenas preserva o fluxo e avisa que a finalizacao Windows nao foi aplicada

## Próxima etapa (port completo no Mac)

Se você quiser manter a “finalizacao no Word”, a evolução correta é criar um backend real para Mac com:

- `osascript` (AppleScript) + Microsoft Word para Mac, ou
- LibreOffice headless (se a finalizacao puder ser diferente)

## Defaults recuperados da UI original (`App.py`)

- `font_name`: `Arial`
- `font_size`: `11`
- `badge_width_cm`: `1.3`
- `column_width_cm`: `7.5`
- `remove_gabarito`: `True`
- `justify`: `True`
- `finalize_word`: `True`
- `force_inline_wrap`: `True`
- `area_conhecimento`: `biologia` (novo no port)
