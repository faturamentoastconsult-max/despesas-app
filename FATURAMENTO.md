# Conector de Faturamento

Substitui o preenchimento manual da planilha **Controle de Faturamento** a partir de TXTs.
Abra `faturamento.html` no navegador — não precisa instalar nada, não precisa de servidor,
nenhum arquivo sai da máquina.

## O desenho

A descoberta que definiu a ferramenta: **o XML mensal já traz tudo**. Um único
`ConsultarNfseFaixaResposta` contém as notas do mês inteiro, e a `<Discriminacao>` de cada
nota já carrega a quebra por evento:

```
VALOR SALARIOS/ENCARGOS.........R$ 415.812,01
VALOR DE TAXA DE SERVICOS.......R$  52.392,13
VALOR DE ENCARGOS S/TRIBUTOS....R$  86.636,55
DEDUCOES EVENTO 671/672 ....    R$ -25.955,29
BASE DE CALCULO P/ IMPOSTOS.....R$ 528.885,40
```

Não é preciso um arquivo de rateio por nota para montar o controle — os eventos vêm de dentro
da própria nota.

> **Regra central:** a soma dos eventos da discriminação tem de bater com o valor da nota.
> O que não bate é sinalizado em vez de entrar calado na planilha.

Os **demonstrativos de faturamento** (`.xlsx`, um por NF) continuam úteis, mas como segunda
camada opcional: eles conferem a nota contra a folha e trazem o detalhe por **verba**
(`002-SALARIO MES`, `091-HORA EXTRAS 65%`…) e por **centro de custo**.

| Camada | Origem | Traz | Necessária? |
|---|---|---|---|
| Nota | XML mensal da NFSe | NF, cliente, CNPJ, valores, vencimento, pedido, **eventos** | sim — resolve o mês |
| Folha | Demonstrativo `.xlsx` por NF | verbas, centros de custo, retenções | opcional |

## Passo a passo

1. **Competência** — mês do lote (padrão `08/2026`), tolerância de arredondamento
   (padrão R$ 0,02) e qual valor vai para o controle (serviços ou líquido).
2. **Notas fiscais** — arraste o XML do mês. Um arquivo, todas as notas.
3. **Demonstrativos** *(opcional)* — arraste quantos quiser; cada um é casado com a sua NF
   pelo número no título.
4. **Conciliação** — cada nota recebe um status:

   | Status | Significado |
   |---|---|
   | `OK` | os eventos somam o valor da nota |
   | `DISCRIMINACAO NAO FECHA` | os eventos da própria nota não somam o valor dela |
   | `SEM EVENTOS` | a discriminação não trouxe nenhum evento reconhecível |
   | `DEMONSTRATIVO DIVERGENTE` | o demonstrativo totaliza diferente da nota |
   | `SEM NF` | demonstrativo sem a nota correspondente no XML |
   | `FORA DA COMPETENCIA` | nota emitida em outro mês |

5. **Saída** — `.xlsx`, CSV, ou copiar as linhas para colar.

## O arquivo gerado

| Aba | Conteúdo |
|---|---|
| `Controle_Faturamento` | **uma linha por NF**, com uma coluna para cada evento do mês |
| `Eventos_Detalhe` | formato longo: uma linha por NF × evento |
| `Consolidado_Eventos` | total de cada evento no mês e o quanto representa do faturado |
| `Divergencias` | só o que precisa de ação, com o motivo |
| `Demonstrativos`, `Centros_de_Custo`, `Verbas_Detalhe` | quando houver demonstrativos |

A aba principal é um **pivot**: as colunas de evento são descobertas a partir dos dados do
mês, então um evento novo aparece sozinho como coluna nova. `ID_Unico` (`competência | NF`)
é a chave de reimportação — reprocessar o mês atualiza a linha em vez de duplicar.

## Como os eventos são reconhecidos

A discriminação é texto livre, e o layout varia entre notas: o separador pode ser `|` ou
quebra de linha, e há rótulos com acento, sem espaço ou com espaço a mais
(`TOTAL DAS DEDUÇÕES`, `VALOR DE ENCARGOSS/TRIBUTOS`, `BASECBS/IBS`).

O parser trata isso com uma **chave canônica** — maiúscula, sem acento, sem espaço nem
pontuação — de modo que as variações de grafia caiam no mesmo evento no consolidado.

A classificação é por exclusão: rótulos derivados (`BASE DE CALCULO`, `VALOR LIQUIDO A PAGAR`,
`VALOR BRUTO`, `TOTAL`) não entram na soma; linhas informativas (`BASE CBS/IBS`, `VENCTO`,
`DEPOSITO`, `PEDIDO`, `DS`) são descartadas; **todo o resto é evento**. Assim um rótulo novo
entra sozinho, sem mexer no código — e se for classificado errado, a conferência de soma
denuncia na hora.

## Validado contra dados reais

Sobre o XML de 08/2026 (210 notas, 148 clientes, R$ 8.667.370,58):

- **202 notas fecham** — a soma dos eventos bate com o valor da nota
- **8 notas não fecham** por inconsistência da própria discriminação, não do parser
  (ex.: a NF 27279 declara base de R$ 349.705,23 mas os componentes somam R$ 314.667,67)
- o demonstrativo da NF 27382 totaliza R$ 528.885,38 contra R$ 528.885,40 do XML —
  2 centavos, dentro da tolerância padrão

## Ajustando às suas colunas

Os cabeçalhos fixos estão em `colsControle()` e a montagem das linhas em `linhasControle()`,
dentro de `faturamento.html`. Para casar exatamente com a sua planilha, reordene/renomeie os
dois na mesma ordem — as colunas de evento entram automaticamente entre eles.

## Limitações conhecidas

- `.xls` antigo (binário) não é lido — salve como `.xlsx`.
- NFSe não tem layout único no país; o leitor cobre o padrão ABRASF (validado na 2.02).
- Grafias que diferem por uma palavra inteira não são unificadas: `REEMB.UNIFORMES,EPIS E
  OUTROS` e `REEMB UNIFORMES,EPIS,OUTROS` aparecem como dois eventos. É proposital — juntar
  por semelhança arriscaria fundir eventos que são mesmo distintos.
- A leitura de `.xlsx` usa `DecompressionStream` (Chrome/Edge 103+, Firefox 113+,
  Safari 16.4+). Em navegador antigo, só o passo 3 fica indisponível.
