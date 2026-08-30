// =====================================================================
//  Consolidado de Rateios — consulta única, sem lista fixa de colunas
//  Cole em: Dados > Obter Dados > Consulta em Branco > Editor Avançado
//  Nome sugerido da consulta: fRateios
// =====================================================================
let
    // ── 1. Parâmetros (o único trecho que muda de ano para ano) ──────
    PastaOrigem       = "C:\Users\Jaqueline\OneDrive - Steintemp Gestao de Pessoas\FTD CONSOLIDADA\2026",
    ArquivoConsolidado = "2026.xlsx",

    // ── 2. Arquivos de rateio da pasta ──────────────────────────────
    Fonte = Folder.Files(PastaOrigem),
    Arquivos = Table.SelectRows(Fonte, each
           [Attributes]?[Hidden]? <> true
       and not Text.StartsWith([Name], "~$")
       and List.Contains({".xlsx", ".xlsm", ".xls"}, Text.Lower([Extension]))
       and [Name] <> ArquivoConsolidado),

    // ── 3. Colunas de identificação (tudo o que NÃO for isso vira rubrica) ──
    Chaves = {"ANO", "MÊS", "Emp", "Fil", "Cod", "C.Custo", "C.Custo (Descrição)",
              "Nome", "CPF", "Função", "Salário", "Dt Admissão", "Dt Demissão",
              "Taxa Fat", "Período Bnf"},

    // Colunas agregadas/repetidas que não podem entrar no unpivot.
    // "Total Geral" é soma das demais: mantê-la dobraria o faturamento.
    Redundantes = {"TOTAL GERAL", "NOME", "CENTRO DE CUSTO", "VALOR", "TAXA FAT"},
    EhRedundante = (n as text) as logical =>
        let semSufixo = Text.Upper(Text.Trim(Text.Replace(n, ".1", "")))
        in List.Contains(Redundantes, semSufixo),

    // ── 4. Leitura de um arquivo, seja qual for o conjunto de colunas ──
    LerRateio = (Conteudo as binary, NomeArquivo as text) as table =>
        let
            Livro    = Excel.Workbook(Conteudo, true, true),
            Planilha = Table.SelectRows(Livro, each [Kind] = "Sheet"){0}[Data],
            Nomes    = Table.ColumnNames(Planilha),
            SemLixo  = Table.RemoveColumns(Planilha, List.Select(Nomes, each EhRedundante(_))),
            ChavesAqui = List.Intersect({Chaves, Table.ColumnNames(SemLixo)}),
            Longo    = Table.UnpivotOtherColumns(SemLixo, ChavesAqui, "Atributo", "Valor"),
            ComNF    = Table.AddColumn(Longo, "NF",
                          each Text.BeforeDelimiter(NomeArquivo, ".", {0, RelativePosition.FromEnd}), type text)
        in
            ComNF,

    Lidos     = Table.AddColumn(Arquivos, "Dados", each LerRateio([Content], [Name])),
    // Table.Combine faz a união das colunas: rubrica nova entra sozinha.
    Empilhado = Table.Combine(Lidos[Dados]),

    // ── 5. Classificação por FORMATO, não por lista fixa ────────────
    //     Rubrica de folha = três dígitos no início ("002-SALARIO MES").
    Classificado = Table.AddColumn(Empilhado, "Tipo", each
        let a = Text.From([Atributo] ?? "")
        in  if a = "Taxa"           then "Taxa"
            else if a = "Tributo Fiscal" then "Tributo"
            else if Text.Length(a) >= 3
                 and Text.Select(Text.Start(a, 3), {"0".."9"}) = Text.Start(a, 3)
                                    then "Rubrica"
            else "Não classificado", type text),

    // ── 6. Tipagem e limpeza ────────────────────────────────────────
    ComValor = Table.TransformColumnTypes(Classificado, {{"Valor", type number}}, "en-US"),
    ComSalario = Table.AddColumn(ComValor, "Salário (R$)", each
        try Number.FromText(Text.Replace(Text.Replace(
              Text.BeforeDelimiter(Text.Trim(Text.From([Salário])), " /"), ".", ""), ",", "."))
        otherwise null, type number),
    Tipado = Table.TransformColumnTypes(ComSalario, {
        {"ANO", Int64.Type}, {"MÊS", Int64.Type}, {"Emp", type text}, {"Fil", type text},
        {"Cod", type text}, {"C.Custo", type text}, {"C.Custo (Descrição)", type text},
        {"Nome", type text}, {"CPF", type text}, {"Função", type text}, {"NF", type text}}),

    // ── 7. Centro de custo normalizado (código + descrição padronizada) ──
    Normalizado = Table.AddColumn(Tipado, "Centro de Custo", each
        let bruto = Text.Upper(Text.Trim(Text.From([#"C.Custo (Descrição)"] ?? "SEM CENTRO DE CUSTO"))),
            um    = Text.Combine(List.Select(Text.Split(bruto, " "), each _ <> "")," "),
            comHifen = if Text.Contains(um, "-")
                       then Text.Trim(Text.BeforeDelimiter(um, "-")) & " - " & Text.Trim(Text.AfterDelimiter(um, "-"))
                       else um,
            corr  = Text.Replace(Text.Replace(comHifen, "LOGITICA", "LOGISTICA"), "LOGISTICAS", "LOGISTICA"),
            corr2 = Text.Replace(Text.Replace(corr, "ARMAZENAMENTO E EXPEDIÇÃO DE P.ACABADOS", "ARMAZENAGEM E EXPEDIÇÃO"),
                                 "ARMAZENAMENTO E EXPEDIÇÃO", "ARMAZENAGEM E EXPEDIÇÃO")
        in  Text.Replace(corr2, "ARMAZENAGEM EXPEDIÇÃO", "ARMAZENAGEM E EXPEDIÇÃO"), type text),

    // ── 8. Só o que é faturamento, sem os zeros (≈52% das linhas) ───
    Final = Table.SelectRows(Normalizado, each
                [Tipo] <> "Não classificado" and [MÊS] <> null
            and [Valor] <> null and [Valor] <> 0)
in
    Final


// =====================================================================
//  CONSULTA 2 — auditoria. Nome: fRateios_NaoClassificado
//  Carregue numa aba pequena. Se vier vazia, a leitura está limpa;
//  se aparecer algo, é coluna nova ou cabeçalho fora do lugar.
// =====================================================================
// let
//     Base = fRateios_Origem,   // aponte para a etapa "Normalizado" acima
//     Rejeitados = Table.SelectRows(Base, each [Tipo] = "Não classificado"),
//     Resumo = Table.Group(Rejeitados, {"NF", "Atributo"}, {
//                  {"Linhas", each Table.RowCount(_), Int64.Type},
//                  {"Soma",   each List.Sum([Valor]),  type number}})
// in
//     Resumo
