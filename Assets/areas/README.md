# Estrutura de Assets por Área

Use esta estrutura para as cápsulas por área do conhecimento:

```text
Assets/
  areas/
    biologia/
      capsulas/
        facil.png
        media.png
        dificil.png
      secoes/
        exercicios_sala.png
        exercicios_propostos.png
        secao_enem.png
        exercicios_aprofundamento.png
        exercicios_regionais.png
        exercicios_dissertativos.png
    fisica/
      capsulas/
        facil.png
        media.png
        dificil.png
      secoes/
        exercicios_sala.png
        exercicios_propostos.png
        secao_enem.png
        exercicios_aprofundamento.png
        exercicios_regionais.png
        exercicios_dissertativos.png
    quimica/
      capsulas/
        facil.png
        media.png
        dificil.png
      secoes/
        exercicios_sala.png
        exercicios_propostos.png
        secao_enem.png
        exercicios_aprofundamento.png
        exercicios_regionais.png
        exercicios_dissertativos.png
```

## Como o script resolve

- Ao executar com `--area biologia`, ele busca:
  - `assets/areas/biologia/capsulas/facil.png`
  - `assets/areas/biologia/capsulas/media.png`
  - `assets/areas/biologia/capsulas/dificil.png`
  - `assets/areas/biologia/secoes/exercicios_sala.png`
  - `assets/areas/biologia/secoes/exercicios_propostos.png`
  - `assets/areas/biologia/secoes/secao_enem.png`
  - `assets/areas/biologia/secoes/exercicios_aprofundamento.png`
  - `assets/areas/biologia/secoes/exercicios_regionais.png`
  - `assets/areas/biologia/secoes/exercicios_dissertativos.png`

## Regra de nomes

- Pasta da área: minúscula e sem acentos (ex.: `linguagens`, `matematica`)
- Arquivos das cápsulas: sempre `facil.png`, `media.png`, `dificil.png`
- Arquivos de seção: `exercicios_sala.png`, `exercicios_propostos.png`, `secao_enem.png`, `exercicios_aprofundamento.png`, `exercicios_regionais.png`, `exercicios_dissertativos.png`
