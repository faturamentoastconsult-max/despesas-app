# Conector de Faturamento

Ferramenta para substituir o preenchimento manual da planilha **Controle de Faturamento**
a partir dos TXTs. Abra `faturamento.html` direto no navegador — não precisa instalar nada,
não precisa de servidor e nenhum arquivo sai da máquina.

## O desenho

Duas camadas, e uma regra que amarra as duas:

| Camada | Origem | O que traz |
|---|---|---|
| **Cabeçalho** (a nota) | XML da NFe / NFSe | nº da NF, série, cliente, CNPJ, emissão, **valor total** |
| **Itens** (os eventos) | Rateio em `.xlsx` ou `.csv` | evento, **valor do evento** |

> **Regra central:** a soma dos eventos do rateio tem de bater com o total da nota.
> O que não bate vira divergência e é sinalizado — não entra escondido na planilha.

É essa conferência que hoje você faz no olho. A ferramenta assume ela.

## Passo a passo

1. **Competência** — o mês do lote (padrão `08/2026`), a tolerância de arredondamento
   (padrão R$ 0,02) e o que fazer com notas emitidas fora do mês.
2. **Notas fiscais** — arraste todos os XMLs de uma vez. Duplicados são detectados pela
   chave de acesso; arquivos de evento/cancelamento são ignorados.
3. **Rateio** — arraste a planilha. A ferramenta detecta os cabeçalhos e propõe o de-para
   das colunas (NF, evento, valor, cliente). **O mapeamento fica salvo no navegador**,
   então nos meses seguintes já vem preenchido.
4. **Conciliação** — cada nota é cruzada com o seu rateio e recebe um status:

   | Status | Significado |
   |---|---|
   | `OK` | soma do rateio = total da nota |
   | `DIVERGENTE` | soma do rateio ≠ total da nota (diferença mostrada) |
   | `SEM RATEIO` | nota importada sem nenhum evento correspondente |
   | `SEM NF` | evento de rateio sem XML correspondente |
   | `FORA DA COMPETENCIA` | nota emitida em outro mês |

5. **Saída** — `.xlsx` com quatro abas, ou CSV, ou copiar as linhas para colar direto.

## O arquivo gerado

| Aba | Conteúdo |
|---|---|
| `Controle_Faturamento` | uma linha por NF × evento — o que vai para a planilha principal |
| `Consolidado_Eventos` | total de cada evento no mês, com quantidade de notas |
| `Divergencias` | só o que precisa de ação, com o motivo |
| `Notas` | as notas lidas, para auditoria contra os XMLs |

A coluna **`ID_Unico`** (`competência | NF | evento`) é a chave de reimportação: reprocessar
o mesmo mês atualiza a linha em vez de duplicar. Se o mesmo evento aparecer duas vezes na
mesma nota, a segunda ocorrência ganha sufixo `#2`.

## Ajustando às suas colunas

Os cabeçalhos da aba `Controle_Faturamento` estão em `COLS_CONTROLE` e a montagem das
linhas em `linhasControle()`, dentro de `faturamento.html`. Para casar exatamente com a sua
planilha atual, basta reordenar/renomear os dois na mesma ordem.

## Limitações conhecidas

- `.xls` antigo (binário) não é lido — salve como `.xlsx` ou CSV.
- NFSe não tem layout único no país; o leitor cobre o padrão ABRASF. Se o seu município
  fugir do padrão, o arquivo aparece como "layout não reconhecido" e o parser precisa de um
  ajuste pontual.
- A leitura de `.xlsx` usa `DecompressionStream`, disponível em navegadores modernos
  (Chrome/Edge 103+, Firefox 113+, Safari 16.4+). Em navegador antigo, use CSV — um aviso
  aparece no topo da tela.
