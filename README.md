# despesas-app

Duas páginas independentes, sem build e sem instalação:

| Arquivo | O que é |
| --- | --- |
| `index.html` | App de gestão de despesas (login, recibos, saldos, aprovações) — usa Google Sheets via Apps Script |
| `consolidador.html` | **Link compartilhado** para upload de Excel que gera o resumo consolidado de faturamento |

---

## Consolidador de Demonstrativos (`consolidador.html`)

Página pública para quem precisa transformar o Excel **Consolidado** (uma aba por
cliente, no formato "Demonstrativo de Faturamento") no resumo com
**Cliente / Total / RITM / PEDIDO**.

### Como funciona

1. A pessoa abre o link e arrasta o arquivo `.xlsx`, `.xlsm` ou `.xls`.
2. A página lê cada aba e monta o consolidado na hora.
3. A coluna **PEDIDO** é digitada na tela (não existe no arquivo de origem).
4. Botão **Baixar resumo (.xlsx)** gera a planilha final; também há export em CSV.

O arquivo é lido **dentro do navegador** — nada é enviado para servidores, então não
há dados armazenados em nenhum lugar depois que a aba é fechada.

### O que é extraído de cada aba

| Campo | Origem na planilha |
| --- | --- |
| Cliente | texto após `CLIENTE:` |
| Total | coluna `TOTAL` da linha `TOTAL:  N Funcionários` |
| RITM | qualquer célula com `RITM` + números |
| DS | texto após `Nº DS.:` (exibido só na tela) |
| Período | texto após `MÊS/ANO:` (usado no nome do arquivo baixado) |

Regras aplicadas:

- Abas sem `CLIENTE:` são ignoradas e listadas no painel de avisos.
- Uma aba pode conter mais de um cliente: cada `CLIENTE:` inicia um novo bloco.
- Se o mesmo cliente tiver mais de uma linha de totais, os valores são somados.
- Vários RITMs no mesmo cliente saem separados por ` / `.
- Nada é descartado em silêncio: total ou RITM faltando vira aviso na tela.

As posições de linha/coluna variam entre abas, por isso tudo é localizado por
conteúdo e não por célula fixa.

### Arquivo gerado

Mesmo layout do resumo modelo (`ResumoDs.xlsx`):

```
A1            Consolidado - Faturamento
A2:D2         Cliente | Total | RITM | PEDIDO
A3:D...       uma linha por cliente
última linha  TOTAL GERAL | soma
```

Nome do arquivo: `Resumo_DS_<mês>-<ano>.xlsx` (ex.: `Resumo_DS_04-2026.xlsx`).

### Publicando o link compartilhado

1. **Settings > Pages > Source:** "Deploy from a branch"
2. **Branch:** `main` / `(root)` > **Save**
3. O link fica: `https://SEU-USUARIO.github.io/despesas-app/consolidador.html`

O botão **Copiar link** no topo da página copia esse endereço.

> Quem tiver o link consegue abrir a página. Como nada é gravado em servidor, não há
> risco de acesso a arquivos já enviados — mas, para restringir quem abre, publique
> em um site interno ou atrás do login da empresa.

### Detalhes técnicos

- A leitura de Excel usa **SheetJS 0.18.5**, com cópia local em
  `vendor/xlsx.full.min.js` para funcionar em redes que bloqueiam CDN. Se o arquivo
  local faltar, a página recorre à CDN automaticamente.
- Os números de pedido digitados ficam salvos no `localStorage` do navegador, por
  código de cliente, e voltam sozinhos no próximo envio.
- Para publicar, os arquivos necessários são `consolidador.html` e a pasta `vendor/`.

---

## App de despesas (`index.html`)

As instruções de configuração do Google Sheets + Apps Script + Drive estão no
comentário no final do próprio `index.html`.
