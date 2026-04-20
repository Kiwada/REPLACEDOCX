# Estrutura de Assets por Área

Use esta estrutura para as cápsulas por área do conhecimento:

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
    fisica/
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
    quimica/
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
    linguagens/
      capsulas/
        facil.png|jpg|jpeg
        media.png|jpg|jpeg
        dificil.png|jpg|jpeg
      secoes/
        exercicios_sala.png|jpg|jpeg
        exercicios_propostos.png|jpg|jpeg
        secao_enem.png|jpg|jpeg
    historia|filosofia|sociologia/
      capsulas/
        facil.png|jpg|jpeg
        media.png|jpg|jpeg
        dificil.png|jpg|jpeg
      secoes/
        exercicios_basicos.png|jpg|jpeg
        exercicios_propostos.png|jpg|jpeg
        secao_enem.png|jpg|jpeg
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
- Arquivos das cápsulas: sempre `facil`, `media`, `dificil` (`.png`, `.jpg` ou `.jpeg`)
- Arquivos de seção: normalmente `exercicios_sala`, `exercicios_propostos`, `secao_enem`, `exercicios_aprofundamento`, `exercicios_regionais`, `exercicios_dissertativos` (`.png`, `.jpg` ou `.jpeg`)
- Em `linguagens`, use apenas `exercicios_sala`, `exercicios_propostos` e `secao_enem`
- Em `historia`, `filosofia` e `sociologia`, use apenas `exercicios_basicos`, `exercicios_propostos` e `secao_enem`
- Equivalência de título (especial): `Questão Regional`, `Questões Regionais` e `Exercício Regional` são tratados como `Exercícios Regionais`.
