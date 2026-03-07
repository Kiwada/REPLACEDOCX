# ContextoCLI (ReplaceDocx)

Ferramenta de terminal para processar arquivos `.docx` de exercicios.

## O que ela faz

- troca marcadores de dificuldade por capsulas visuais
- troca titulos de secao por banners
- remove linhas de resposta (`GABARITO`, `RESPOSTA`) incluindo alternativa
- adiciona tabelas de autoavaliacao no final
  - formato: `Questão | Nível | Gabarito | Acertei | Errei | Revisar` (todas as áreas)
  - remove automaticamente bloco final de gabarito por seção, quando existir
- gera relatorio por secao (`facil`, `media`, `dificil`)
- ajusta para fonte 8 apenas referencias abaixo de imagens (`Fonte`, `Referência`)
- reconhece variacoes de marcador de dificuldade:
  - `1. (MÉDIA) ...`
  - `(MÉDIA) 1. ...`
- cria saida final com sufixo `_ok`

## Requisitos

- macOS
- Python 3.11+
- `python-docx`
- opcional: `weasyprint` (PDF igual ao HTML)
- opcional: `reportlab` (fallback de PDF)

## Instalacao

```bash
cd /Users/kiwada/Desktop/ReplaceDocx
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

Processar por area:

```bash
python contextocli.py process "/caminho/arquivo.docx" --area quimica
python contextocli.py process "/caminho/arquivo.docx" --area fisica
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

Os arquivos sao salvos em:

`/Volumes/anthropic_externo/Documentos Processados CONTEXTO/`

Se o volume nao estiver montado ou sem permissao de escrita, a CLI retorna erro (nao usa fallback local).

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
- `--with-report-appendix` (anexa o relatório A4 no DOCX; por padrão, fica apenas o quadro de autoavaliação)
- `--no-report-pdf`
- `--section-banner-width-cm 7.5`
- `-o nome_saida.docx` (apenas nome; sempre salva no diretorio de saida fixo)

## Compatibilidade

Estes comandos antigos continuam funcionando:

- `python replacedocx_cli.py ...`
- `python run_docx_real.py ...`

## Estrutura

- `contextocli.py`: entrada principal da CLI
- `src/replacedocx/`: pacote principal (cli, engine, renderizadores e relatorios)
- `scripts/test_run.py`: teste rapido local
- `Assets/`: capsulas e artes por area
- `docs/USAGE.md`: guia objetivo de uso

## Problemas comuns

Caminho com espacos/acentos:

```bash
python contextocli.py process "/Volumes/.../CAPITULO_8_...docx" --area biologia
```

- A CLI aceita apenas `.docx` (nao processa `.doc` diretamente).
- Se JPG nao entrar no DOCX, a engine converte automaticamente para PNG.
- Para estabilidade visual no Word, imagens JPG (inclusive CMYK) sao convertidas para PNG sRGB durante a insercao.
