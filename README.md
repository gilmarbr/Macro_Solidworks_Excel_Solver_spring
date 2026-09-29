# Otimização de Mola de Compressão (Estampo/Molde) — Excel Solver + SolidWorks

Projeto de exemplo que mostra como usar o **Solver do Excel** para otimizar o
projeto de uma mola de compressão (peso mínimo, respeitando força, tensão e
espaço disponível) e sincronizar automaticamente o resultado com um modelo
paramétrico no **SolidWorks** — incluindo cor por classe de força (padrão
ISO 10243) e captura de uma vista isométrica da peça direto na planilha.

Feito para acompanhar [este vídeo](#) *(troque pelo link do seu vídeo)*.

## O que este projeto faz

1. Você preenche os dados do seu projeto (força exigida, espaço disponível,
   material) numa planilha Excel.
2. O **Solver** encontra a combinação de diâmetro do arame, diâmetro médio
   da espira e número de espiras que resulta na **menor massa possível**,
   respeitando todas as restrições de engenharia.
3. Uma macro em VBA:
   - atualiza automaticamente a peça da mola no **SolidWorks** com os
     valores otimizados;
   - lê de volta a **massa real** calculada pelo SolidWorks (considerando o
     material atribuído à peça), para conferência;
   - pinta a mola de acordo com a classe de força (verde/azul/vermelho/amarelo,
     inspirado na ISO 10243 para molas de estampo);
   - captura uma **vista isométrica** da peça e insere a imagem direto na
     planilha.

## Pré-requisitos

- Microsoft Excel (Windows), com o suplemento **Solver** ativado
  (`Arquivo > Opções > Suplementos > Gerenciar: Suplementos do Excel > Ir... > Solver`).
- No editor VBA (`Alt+F11`), a referência **SOLVER** marcada em
  `Ferramentas > Referências`.
- SolidWorks instalado na mesma máquina, com uma peça de mola modelada por
  Helix + Sweep, controlada pelas Global Variables `d_wire`, `D_mean`, `Na`
  (ver seção [Modelo no SolidWorks](#modelo-no-solidworks) abaixo).

## Estrutura do repositório

```
├── planilha/
│   └── MolaOtimizacao.xlsx      # planilha com fórmulas, nomes definidos e modelo do Solver
├── macro/
│   └── OtimizacaoMola.bas       # módulo VBA para importar no Excel
└── README.md
```

## Como usar

1. Abra `planilha/MolaOtimizacao.xlsx` e leia a aba **LEIAME** (passo a passo
   completo está lá também).
2. No SolidWorks, modele a mola com as Global Variables `d_wire`, `D_mean`,
   `Na` e `L_free` (`Equations > Add`), e ligue `d_wire`, `D_mean` e `Na` à
   planilha via `Equations > Link Values to External File`, apontando para
   as células de mesmo nome na aba **Projeto**.
3. No editor VBA (`Alt+F11`), importe `macro/OtimizacaoMola.bas`
   (`Arquivo > Importar Arquivo`).
4. Ajuste a constante `CAMINHO_PECA`, no topo do módulo, para o caminho do
   seu arquivo `.SLDPRT`.
5. Salve a planilha como **.xlsm** (Excel Habilitado para Macro) para manter
   o código importado.
6. Preencha as células amarelas na aba **Projeto** com os dados do seu
   projeto e do material.
7. Rode `Main_OtimizarEAtualizar` (`Alt+F8`).

## Modelo no SolidWorks

| Global Variable | Significado | Alimenta |
|---|---|---|
| `d_wire` | diâmetro do arame (mm) | diâmetro do perfil da Sweep |
| `D_mean` | diâmetro médio da espira (mm) | parâmetro *Diameter* da Helix |
| `Na` | número de espiras ativas | usado para calcular `Nt = Na + 2` |
| `L_free` | comprimento livre (mm) | parâmetro *Height* da Helix |

Os nomes são **idênticos** aos nomes definidos na planilha, de propósito —
uma única fonte de verdade entre os dois programas.

## O modelo de otimização

**Minimizar:** massa do fio (`ρ · π²/4 · d² · D · Nt`)

**Variáveis:** diâmetro do arame, diâmetro médio da espira, número de
espiras ativas

**Restrições:**
- força calculada ≥ força de trabalho exigida
- tensão de cisalhamento ≤ tensão admissível / fator de segurança
- diâmetro externo ≤ diâmetro do furo de alojamento
- índice da mola (D/d) entre 4 e 12
- comprimento sólido + folga mínima ≤ comprimento livre − curso de trabalho

## Avisos importantes

- **Classificação de força por cor:** a norma ISO 10243 classifica uma mola
  comparando-a com o catálogo do fabricante para o **mesmo diâmetro e
  comprimento livre** — não existe um valor de força absoluto universal. Os
  limites usados neste projeto (`lim_leve_max`, `lim_media_max`,
  `lim_pesada_max`) são uma aproximação didática; confirme a classe real no
  catálogo do fabricante escolhido antes de usar em produção.
- **Material:** para a massa calculada no Excel bater com a massa real do
  SolidWorks, a densidade e o módulo de cisalhamento usados nas fórmulas
  precisam ser exatamente os mesmos do material atribuído à peça.
- **Aviso do Windows/antivírus ao baixar o `.xlsm`:** isso é comum para
  qualquer planilha com macro baixada da internet, mesmo sendo inofensiva.
  Se o Windows bloquear, clique em **Mais informações > Executar assim
  mesmo** — o código completo da macro está neste repositório para você
  conferir antes.
- Este é um projeto **educacional**, feito para demonstrar o uso do Solver
  e a integração Excel–SolidWorks. Revise as fórmulas e os fatores de
  segurança antes de aplicar a um projeto real.

## Licença

Distribuído sob a licença MIT — veja o arquivo `LICENSE`. Sinta-se livre
para usar, adaptar e redistribuir, mantendo os créditos.

## Créditos

Projeto criado para fins didáticos, acompanhando um vídeo tutorial sobre
otimização mecânica com o Solver do Excel.
