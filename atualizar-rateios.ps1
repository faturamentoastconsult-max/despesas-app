<#
    atualizar-rateios.ps1
    ---------------------------------------------------------------------------
    Atualiza a pasta consolidada de rateios sem ninguém abrir o Excel:
    atualiza as consultas, roda as conferências, e SÓ ENTÃO gera e publica a
    cópia do cliente. Se alguma conferência acusar, ele para e não publica.

    Rode UMA VEZ à mão antes de agendar. O Excel precisa estar instalado na
    máquina que executa, e essa máquina precisa estar ligada no horário.

    Uso:
      powershell -ExecutionPolicy Bypass -File .\atualizar-rateios.ps1
      powershell -ExecutionPolicy Bypass -File .\atualizar-rateios.ps1 -Simular
#>

[CmdletBinding()]
param(
    # A sua cópia de trabalho, com a aba Rateios e as consultas.
    [string] $PastaTrabalho = "C:\Rateios\Rateios_FTD_2026.xlsx",

    # Onde a cópia do cliente deve ser publicada (a pasta sincronizada do drive).
    [string] $DestinoCliente = "C:\Users\Jaqueline\OneDrive - Steintemp Gestao de Pessoas\FTD CONSOLIDADA\Rateios_FTD_2026_cliente.xlsx",

    # Aba de detalhe que não vai para o cliente.
    [string] $AbaDetalhe = "Rateios",

    [string] $PastaLog = "C:\Rateios\log",

    # Passa por tudo e mostra o resultado das conferências, sem publicar nada.
    [switch] $Simular
)

$ErrorActionPreference = 'Stop'
$inicio = Get-Date

if (-not (Test-Path $PastaLog)) { New-Item -ItemType Directory -Path $PastaLog -Force | Out-Null }
$log = Join-Path $PastaLog ("rateios-{0:yyyy-MM-dd_HHmm}.log" -f $inicio)

function Escrever([string] $texto, [string] $nivel = 'INFO') {
    $linha = "{0:HH:mm:ss}  {1,-5}  {2}" -f (Get-Date), $nivel, $texto
    Write-Host $linha
    Add-Content -Path $log -Value $linha -Encoding UTF8
}

$excel = $null; $wb = $null; $wbCliente = $null
# GetTempPath() sempre resolve; $env:TEMP vem nulo quando o Agendador roda
# a tarefa numa conta de serviço.
$temporario = Join-Path ([System.IO.Path]::GetTempPath()) ("rateios-cliente-{0:yyyyMMddHHmmss}.xlsx" -f $inicio)

try {
    if (-not (Test-Path $PastaTrabalho)) { throw "Não achei a pasta de trabalho em $PastaTrabalho" }
    Escrever "Início. Pasta de trabalho: $PastaTrabalho"

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    Escrever "Abrindo a pasta…"
    $wb = $excel.Workbooks.Open($PastaTrabalho, 0, $false)

    # ── 1. Atualizar ───────────────────────────────────────────────────────
    Escrever "Atualizando as consultas…"
    $wb.RefreshAll()
    # Sem isto o script segue antes do Power Query terminar e lê número velho.
    $excel.CalculateUntilAsyncQueriesDone()
    $excel.CalculateFullRebuild()
    Escrever "Consultas atualizadas."

    # ── 2. Conferências ────────────────────────────────────────────────────
    # Lê pelo NOME da coluna, não pela letra: se a ordem mudar, o script
    # reclama em vez de conferir a célula errada.
    function Coluna([object] $aba, [string] $titulo, [int] $linhaCabecalho = 4) {
        for ($c = 1; $c -le 60; $c++) {
            if ([string]$aba.Cells.Item($linhaCabecalho, $c).Text -eq $titulo) { return $c }
        }
        throw "Não achei a coluna '$titulo' na aba '$($aba.Name)'. A ordem das colunas mudou?"
    }
    function UltimaLinha([object] $aba, [int] $col) {
        return $aba.Cells.Item($aba.Rows.Count, $col).End(-4162).Row   # xlUp
    }

    $problemas = New-Object System.Collections.Generic.List[string]

    $abaNF   = $wb.Worksheets.Item("Notas Fiscais")
    $colSit  = Coluna $abaNF "Situação"
    $ultNF   = UltimaLinha $abaNF $colSit
    $verificar = 0
    for ($r = 5; $r -le $ultNF; $r++) {
        if ([string]$abaNF.Cells.Item($r, $colSit).Text -eq "Verificar") { $verificar++ }
    }
    Escrever "Conferência 1 · notas marcadas como Verificar: $verificar"
    # Duas são conhecidas e continuam abertas; a terceira é novidade.
    if ($verificar -gt 2) { $problemas.Add("$verificar notas não fecham contra o Total Geral (esperado: 2 conhecidas)") }

    foreach ($nome in @("Centro de Custo", "Notas Fiscais")) {
        $aba = $wb.Worksheets.Item($nome)
        $achou = $false
        for ($r = 5; $r -le ($aba.UsedRange.Rows.Count + 10); $r++) {
            for ($c = 1; $c -le 4; $c++) {
                if ([string]$aba.Cells.Item($r, $c).Text -like "Não listado*") {
                    $achou = $true
                    $soma = 0.0
                    for ($cc = 1; $cc -le 30; $cc++) {
                        $v = $aba.Cells.Item($r, $cc).Value2
                        if ($v -is [double]) { $soma += [math]::Abs($v) }
                    }
                    Escrever ("Conferência 2 · {0}, linha Não listado: {1:N2}" -f $nome, $soma)
                    if ($soma -gt 0.05) { $problemas.Add("$nome tem R$ $($soma.ToString('N2')) fora da lista") }
                    break
                }
            }
            if ($achou) { break }
        }
        if (-not $achou) { $problemas.Add("Não achei a linha 'Não listado' na aba $nome") }
    }

    $abaAud = $wb.Worksheets | Where-Object { $_.Name -like "*Classificad*" } | Select-Object -First 1
    if ($abaAud) {
        $linhas = [math]::Max(0, $abaAud.UsedRange.Rows.Count - 4)
        Escrever "Conferência 3 · linhas não classificadas: $linhas"
        if ($linhas -gt 0) { $problemas.Add("$linhas linhas caíram em 'Não classificado' — coluna nova ou cabeçalho fora do lugar") }
    } else {
        Escrever "Conferência 3 · aba de auditoria não encontrada, pulando." "AVISO"
    }

    $abaResumo = $wb.Worksheets.Item("Resumo")
    $total = $abaResumo.Range("A5").Value2
    Escrever ("Total faturado no período: {0:N2}" -f $total)
    if (-not ($total -is [double]) -or $total -le 0) { $problemas.Add("O total do Resumo não é um número positivo") }

    # ── 3. Salvar a de trabalho ────────────────────────────────────────────
    if ($Simular) {
        Escrever "Modo simulação: nada foi salvo." "AVISO"
    } else {
        $wb.Save()
        Escrever "Pasta de trabalho salva."
    }

    if ($problemas.Count -gt 0) {
        foreach ($p in $problemas) { Escrever $p "ERRO" }
        throw "Conferência acusou $($problemas.Count) problema(s). A cópia do cliente NÃO foi publicada."
    }
    Escrever "As conferências passaram."

    # ── 4. Cópia do cliente ────────────────────────────────────────────────
    if ($Simular) {
        Escrever "Modo simulação: cópia do cliente não gerada." "AVISO"
    } else {
        Escrever "Gerando a cópia do cliente…"
        $wb.SaveCopyAs($temporario)
        $wbCliente = $excel.Workbooks.Open($temporario, 0, $false)

        # As consultas não vão junto: o cliente não tem acesso à pasta de origem
        # e um "Atualizar" lá só produziria erro.
        foreach ($con in @($wbCliente.Connections)) {
            try { $con.OLEDBConnection.RefreshOnFileOpen = $false } catch { }
            try { $con.Delete() } catch { }
        }
        $wbCliente.Worksheets.Item($AbaDetalhe).Delete()
        $wbCliente.Worksheets.Item(1).Activate()

        $pastaDestino = Split-Path $DestinoCliente -Parent
        if (-not (Test-Path $pastaDestino)) { throw "A pasta de destino não existe: $pastaDestino" }
        $wbCliente.SaveAs($DestinoCliente, 51)   # 51 = xlsx
        $wbCliente.Close($false); $wbCliente = $null

        $tam = [math]::Round((Get-Item $DestinoCliente).Length / 1MB, 2)
        Escrever "Cópia do cliente publicada em $DestinoCliente ($tam MB)."
        if ($tam -gt 4) { Escrever "A cópia ficou acima de 4 MB — a aba de detalhe saiu mesmo?" "AVISO" }
    }

    Escrever ("Fim. {0:N0} segundos." -f ((Get-Date) - $inicio).TotalSeconds)
    exit 0
}
catch {
    Escrever $_.Exception.Message "ERRO"
    Escrever "Nada foi publicado. O log completo está em $log" "ERRO"
    exit 1
}
finally {
    foreach ($obj in @($wbCliente, $wb)) {
        if ($obj) { try { $obj.Close($false) } catch { } }
    }
    if ($excel) { try { $excel.Quit() } catch { } }
    foreach ($obj in @($wbCliente, $wb, $excel)) {
        if ($obj) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch { } }
    }
    if (Test-Path $temporario) { Remove-Item $temporario -Force -ErrorAction SilentlyContinue }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
