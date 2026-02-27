# ReplaceDocx (Port para macOS)

Pipeline para processamento de arquivos `.docx` com:

- substituicao de marcadores de dificuldade por capsulas visuais
- substituicao de titulos de secao por artes (banner)
- padronizacao basica de fonte/paragrafo
- remocao de linhas `GABARITO: X`
- geracao de relatorio por secao (facil/media/dificil)

Este projeto foi adaptado de uma versao original Windows e hoje roda via Python no macOS.

## Visao Geral

Entrada:
- um arquivo `.docx` de conteudo
- assets organizados por area (`Assets/areas/<area>/...`)

Saidas:
- `..._ok.docx`
- `..._ok_relatorio_dificuldade.csv`
- `..._ok_relatorio_dificuldade.txt`
- `..._ok_relatorio_dificuldade.html`
- `..._ok_relatorio_dificuldade.pdf` (opcional, se `reportlab` estiver instalado)

Todas as saidas padrao vao para:
- `ReplaceDocx/saida_ok/`

## Estrutura do Projeto

```text
ReplaceDocx/
  Assets/
    areas/
      biologia/
        capsulas/
          facil.png|jpg
          media.png|jpg
          dificil.png|jpg
        secoes/
          exercicios_sala.png|jpg
          exercicios_propostos.png|jpg
          secao_enem.png|jpg
          exercicios_aprofundamento.png|jpg
          exercicios_regionais.png|jpg
          exercicios_dissertativos.png|jpg
  recovered_from_exe/
    reconstructed/
      run_docx_real.py
      replace_engine.py
  saida_ok/
```

## Requisitos

- macOS
- Python 3.11+ (testado com 3.14)
- Dependencia Python:
  - `python-docx`

Opcional:
- Pillow (`PIL`) para conversao de imagens
- sem Pillow, no macOS o script usa `sips` como fallback para JPG nao compativel
- `reportlab` para gerar o relatório em PDF A4

## Instalacao

No terminal:

```bash
cd /Users/kiwada/Desktop/ReplaceDocx/recovered_from_exe/reconstructed
python3 -m venv .venv
source .venv/bin/activate
pip install python-docx reportlab
```

## Uso Rapido

Processar um capitulo:

```bash
cd /Users/kiwada/Desktop/ReplaceDocx/recovered_from_exe/reconstructed
source .venv/bin/activate
python run_docx_real.py "/caminho/arquivo.docx" --area biologia
```

### Parametros principais

- `--area biologia`  
  Define a pasta de assets por area (`Assets/areas/biologia/...`).

- `--no-section-banners`  
  Nao troca os titulos de secao por banners.

- `--no-report-appendix`  
  Nao adiciona o relatório de dificuldade em uma seção A4 ao final do DOCX.

- `--no-report-pdf`  
  Nao gera o arquivo PDF A4 do relatório.

- `--section-banner-width-cm 7.5`  
  Define largura de banner.  
  Default: largura da coluna (`--column-width-cm`, default `7.5`).

- `-o /caminho/saida.docx`  
  Sobrescreve destino da saida.

Ajuda completa:

```bash
python run_docx_real.py -h
```

## Regras de Processamento

### 1) Secoes (banners)

Titulos reconhecidos (normalizados com/sem acento):

- EXERCICIOS DE SALA
- EXERCICIOS PROPOSTOS
- SECAO ENEM
- EXERCICIOS DE APROFUNDAMENTO
- EXERCICIOS REGIONAIS
- EXERCICIO(S) DISSERTATIVO(S)

Comportamento:
- o titulo da secao e substituido completamente pela arte correspondente
- banner fica centralizado
- largura segue `section_banner_width_cm` (ou largura da coluna por default)

### 2) Dificuldade (capsulas)

Padroes reconhecidos para contagem/substituicao:

- `FÁCIL`, `MÉDIA`, `DIFÍCIL`
- com variacoes como:
  - `1. FÁCIL`
  - `(MÉDIA)`
  - `NIVEL: DIFICIL`

### 3) Limpeza e formatacao

- remove linhas `GABARITO: A` ... `GABARITO: E`
- aplica fonte/tamanho configurados
- ajuste de paragrafos basico

## Relatorio por Secao

A cada execucao, alem do DOCX final, sao gerados:

- `*_relatorio_dificuldade.csv`
- `*_relatorio_dificuldade.txt`
- `*_relatorio_dificuldade.html`
- `*_relatorio_dificuldade.pdf` (quando `reportlab` estiver disponível)

No DOCX final:
- o relatório de dificuldade é anexado no fim do documento
- em nova seção A4 (retrato), pronta para exportação/impressão

Conteudo do relatorio:
- nome do conteudo (nome do arquivo de entrada, sem extensao)
- contagem de facil/media/dificil por secao
- total geral

## Convencao de Nomes de Assets

### Capsulas

Pasta:
- `Assets/areas/<area>/capsulas/`

Nomes:
- `facil`
- `media`
- `dificil`

Extensao aceita:
- `.png`, `.jpg`, `.jpeg`

### Secoes

Pasta:
- `Assets/areas/<area>/secoes/`

Nomes:
- `exercicios_sala`
- `exercicios_propostos`
- `secao_enem`
- `exercicios_aprofundamento`
- `exercicios_regionais`
- `exercicios_dissertativos`

Extensao aceita:
- `.png`, `.jpg`, `.jpeg`

## Solucao de Problemas

### Imagem JPG nao entra no DOCX

Alguns JPGs exportados por ferramenta grafica podem nao ser lidos diretamente pelo `python-docx`.

Comportamento atual:
- o motor tenta usar o JPG
- se falhar, converte automaticamente para PNG em cache e segue o processamento

### Arquivo com espacos/acentos no caminho

Use aspas no caminho:

```bash
python run_docx_real.py "/Volumes/.../CAPÍTULO_7_....docx" --area biologia
```

## Arquivos Principais

- `recovered_from_exe/reconstructed/run_docx_real.py`  
  CLI principal.

- `recovered_from_exe/reconstructed/replace_engine.py`  
  Motor de processamento (banners, capsulas, relatorio).

- `recovered_from_exe/reconstructed/PORT_MACOS.md`  
  Guia tecnico complementar do port.
