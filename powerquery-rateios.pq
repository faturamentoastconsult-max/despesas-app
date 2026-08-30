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


// =====================================================================
//  CONSULTA 3 — alimenta a aba "NF x Centro de Custo".
//  Nome: qNFxCentroCusto      Carregar em: tabela nesta pasta
// =====================================================================
// let
//     Agrupado = Table.Group(fRateios, {"MÊS", "NF", "Emp", "Centro de Custo"}, {
//         {"Folha",           each List.Sum(Table.SelectRows(_, each [Tipo] = "Rubrica")[Valor]) ?? 0, type number},
//         {"Taxa e tributos", each List.Sum(Table.SelectRows(_, each [Tipo] <> "Rubrica")[Valor]) ?? 0, type number},
//         {"Total",           each List.Sum([Valor]) ?? 0, type number}}),
//     Ordenado = Table.Sort(Agrupado, {
//         {"MÊS", Order.Ascending}, {"NF", Order.Ascending}, {"Centro de Custo", Order.Ascending}}),
//     // numera os centros dentro de cada nota: é a chave que a aba
//     // "Consulta por NF" usa para montar a distribuição de uma nota só
//     PorNota = Table.Group(Ordenado, {"NF"}, {{"n", each
//                   Table.AddIndexColumn(_, "Ordem na nota", 1, 1, Int64.Type), type table}}),
//     Aberto  = Table.ExpandTableColumn(Table.SelectColumns(PorNota, {"n"}), "n",
//                   {"MÊS", "NF", "Emp", "Centro de Custo", "Folha", "Taxa e tributos", "Total", "Ordem na nota"}),
//     ComChave = Table.AddColumn(Aberto, "Chave",
//                   each [NF] & "|" & Text.From([Ordem na nota]), type text)
// in
//     ComChave


// =====================================================================
//  CONSULTA 4 — alimenta a aba "Rubricas".
//  Os meses viram colunas em ordem crescente, então um mês novo entra
//  sozinho como coluna nova, e uma rubrica nova como linha nova.
//  Nome: qRubricaMes          Carregar em: tabela nesta pasta
// =====================================================================
// let
//     Sel    = Table.SelectColumns(fRateios, {"Grupo", "Código", "Rubrica", "MÊS", "Valor"}),
//     Texto  = Table.TransformColumns(Sel, {{"MÊS", Text.From, type text}}),
//     Agrup  = Table.Group(Texto, {"Grupo", "Código", "Rubrica", "MÊS"},
//                  {{"Valor", each List.Sum([Valor]) ?? 0, type number}}),
//     Meses  = List.Transform(List.Sort(List.Distinct(Agrup[MÊS]), each Number.From(_)), Text.From),
//     Pivot  = Table.Pivot(Agrup, Meses, "MÊS", "Valor", List.Sum),
//     SemNulo = Table.ReplaceValue(Pivot, null, 0, Replacer.ReplaceValue, Meses),
//     ComTotal = Table.AddColumn(SemNulo, "Total",
//                  each List.Sum(List.Transform(Meses, (m) => Record.Field(_, m))), type number),
//     Ordenado = Table.Sort(ComTotal, {
//         {"Grupo", Order.Ascending}, {"Código", Order.Ascending}, {"Rubrica", Order.Ascending}})
// in
//     Ordenado


// =====================================================================
//  CONSULTA 5 — conferência: soma calculada x Total Geral do arquivo.
//  Alimenta as colunas de conferência da aba "Notas Fiscais".
//  Nome: qConferenciaNF       Carregar em: tabela nesta pasta
//
//  Depende de uma variação da consulta principal que NÃO descarta a
//  coluna "Total Geral" — duplique fRateios, tire "TOTAL GERAL" da
//  lista Redundantes, chame de fRateios_ComTotal e use abaixo.
// =====================================================================
// let
//     Informado = Table.Group(
//         Table.SelectRows(fRateios_ComTotal, each [Atributo] = "Total Geral"),
//         {"NF"}, {{"Total Geral informado", each List.Sum([Valor]) ?? 0, type number}}),
//     Calculado = Table.Group(fRateios, {"NF"},
//         {{"Total calculado", each List.Sum([Valor]) ?? 0, type number}}),
//     Junto = Table.NestedJoin(Calculado, {"NF"}, Informado, {"NF"}, "i", JoinKind.FullOuter),
//     Aberto = Table.ExpandTableColumn(Junto, "i", {"Total Geral informado"}),
//     ComDif = Table.AddColumn(Aberto, "Diferença",
//         each Number.Round(([Total calculado] ?? 0) - ([Total Geral informado] ?? 0), 2), type number),
//     Situacao = Table.AddColumn(ComDif, "Situação",
//         each if Number.Abs([Diferença]) < 0.05 then "OK" else "Verificar", type text)
// in
//     Situacao
