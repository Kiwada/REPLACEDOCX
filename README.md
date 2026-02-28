# ContextoCLI (ReplaceDocx)

Ferramenta de terminal para processar arquivos `.docx` de exercicios.

## O que ela faz

- troca marcadores de dificuldade por capsulas visuais
- troca titulos de secao por banners
- remove linhas de resposta (`GABARITO`, `RESPOSTA`) incluindo alternativa
- adiciona tabelas de autoavaliacao no final
- gera relatorio por secao (`facil`, `media`, `dificil`)
- ajusta para fonte 8 apenas referencias abaixo de imagens (`Fonte`, `Referência`)
- cria saida final com sufixo `_ok`

## Requisitos

- macOS
- Python 3.11+
- `python-docx`
- opcional: `weasyprint` (PDF igual ao HTML)
- opcional: `reportlab` (fallback de PDF)

## Instalacao

```bash
cd /Users/kiwada/Desktop/ReplaceDocx/recovered_from_exe/reconstructed
python3 -m venv .venv
source .venv/bin/activate
pip install python-docx reportlab
pip install weasyprint
```

## Uso rapido

Processar documento completo:

```bash
python contextocli.py process "/caminho/arquivo.docx" --area biologia
```

Gerar somente relatorios:

```bash
python contextocli.py report "/caminho/arquivo.docx" --area biologia
```

Validar assets da area:

```bash
python contextocli.py check-assets --area biologia
```

Ajuda:

```bash
python contextocli.py -h
python contextocli.py process -h
```

## Saidas

Por padrao, os arquivos sao salvos em `/Volumes/anthropic_externo/Documentos Processados CONTEXTO/`:

- `<nome>_ok.docx`
- `<nome>_ok_relatorio_dificuldade.html`
- `<nome>_ok_relatorio_dificuldade.pdf` (prioriza `weasyprint` para ficar igual ao HTML)

## Assets (estrutura minima)

```text
Assets/
  areas/
    biologia/
      capsulas/
        facil.png|jpg|jpeg
        media.png|jpg|jpeg
        dificil.png|jpg|jpeg
      secoes/
        exercicios_sala.png|jpg|jpeg
        exercicios_propostos.png|jpg|jpeg
        secao_enem.png|jpg|jpeg
        exercicios_aprofundamento.png|jpg|jpeg
        exercicios_regionais.png|jpg|jpeg
        exercicios_dissertativos.png|jpg|jpeg
```

## Flags uteis (`process`)

- `--area biologia`
- `--no-format-text`
- `--no-section-banners`
- `--no-question-tables`
- `--no-report-appendix`
- `--no-report-pdf`
- `--section-banner-width-cm 7.5`
- `-o /caminho/saida.docx`

## Compatibilidade

Estes comandos antigos continuam funcionando:

- `python replacedocx_cli.py ...`
- `python run_docx_real.py ...`

## Problemas comuns

Caminho com espacos/acentos:

```bash
python contextocli.py process "/Volumes/.../CAPITULO_8_...docx" --area biologia
```

Se JPG nao entrar no DOCX, a engine tenta converter automaticamente para PNG.
