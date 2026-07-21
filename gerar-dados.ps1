# Gera Site/data.json e Ritmo/data.json a partir das bases em Bases/
# Reexecutar sempre que uma base for atualizada (Excel COM trava os .xls entre execucoes;
# se der erro de acesso, feche instancias fantasmas: taskkill /F /IM EXCEL.EXE)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$basesPath = Join-Path $root "Bases"
$sitePath = Join-Path $root "Site"
$ritmoPath = Join-Path $root "Ritmo"

function Remove-Diacritics([string]$s) {
    if (-not $s) { return $s }
    $normalized = $s.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $normalized.ToCharArray()) {
        $cat = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
        if ($cat -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($ch) }
    }
    return $sb.ToString().Normalize([Text.NormalizationForm]::FormC)
}

Write-Host "Abrindo Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# ---------------------------------------------------------------------------
# 1. Fornecedores (codigo -> nome/cor). Mapa CODPROD -> fornecedor vem do catalogo
#    mestre "ESTOQUE API (1).xls" (coluna Cod.Fornec.), lido na secao seguinte.
# ---------------------------------------------------------------------------
$fornecedores = [ordered]@{
    "AB MAURI"  = @{ nome = "Ab Mauri";  cor = "#2a78d6" }
    "ARCOR"     = @{ nome = "Arcor";     cor = "#008300" }
    "BAGLEY"    = @{ nome = "Bagley";    cor = "#e87ba4" }
    "CONDOR"    = @{ nome = "Condor";    cor = "#eda100" }
    "INGLEZA"   = @{ nome = "Ingleza";   cor = "#1baf7a" }
    "PANASONIC" = @{ nome = "Panasonic"; cor = "#eb6834" }
    "WOW"       = @{ nome = "WOW";       cor = "#4a3aa7" }
}

Write-Host "Lendo catalogo mestre ESTOQUE API (mapa produto -> fornecedor, autoritativo via Cod.Fornec.)..."
$fornCodeToId = @{ "6101"="INGLEZA"; "20047"="PANASONIC"; "24185"="ARCOR"; "25987"="BAGLEY"; "35116"="CONDOR"; "28037"="AB MAURI"; "38293"="WOW" }
$codprodToForn = @{}
$wbCat = $excel.Workbooks.Open((Join-Path $basesPath "ESTOQUE API (1).xls"), $null, $true)
$wsCat = $wbCat.Worksheets.Item(1)
$rowsCat = $wsCat.UsedRange.Rows.Count
$arrCat = $wsCat.UsedRange.Value2
for ($r = 2; $r -le $rowsCat; $r++) {
    $cp = "$($arrCat[$r,2])"
    $fcRaw = $arrCat[$r,40]
    if (-not $cp -or $cp -eq "" -or -not ($fcRaw -is [double])) { continue }
    $fc = "$([int]$fcRaw)"
    if ($fornCodeToId.ContainsKey($fc)) { $codprodToForn[$cp] = $fornCodeToId[$fc] }
}
$wbCat.Close($false)
Write-Host "  Mapeados via ESTOQUE API: $($codprodToForn.Count) codigos de produto"
# Nota: produtos vendidos que nao aparecem aqui (ex.: linha "Marcas Exclusivas", cod.
# fornecedor 17097/24318 - batatas/salgadinhos Calira/Skive/Mitbit/Zipoca) sao de proposito
# excluidos do escopo dos 7 fornecedores API - caem em "outrosItensValor", nao em nenhum
# fornecedor. Confirmado com o usuario em 2026-07-21 que esses SKUs realmente nao pertencem
# a nenhum dos 7 (nao aparecem no catalogo mestre sob nenhum dos 7 codigos).

# ---------------------------------------------------------------------------
# 2. Catalogo "Cortes - Geral" (produtos de alocacao controlada / corte)
# ---------------------------------------------------------------------------
Write-Host "Lendo catalogo de cortes..."
$wbCorte = $excel.Workbooks.Open((Join-Path $basesPath "Cortes - Geral.xls"), $null, $true)
$wsCorte = $wbCorte.Worksheets.Item(1)
$rowsCorte = $wsCorte.UsedRange.Rows.Count
$arrCorte = $wsCorte.UsedRange.Value2
$corteInfo = @{}  # codprod -> @{ fornecedor, categoria, alocado }
for ($r = 1; $r -le $rowsCorte; $r++) {
    $cp = "$($arrCorte[$r,1])"
    if (-not $cp -or $cp -eq "") { continue }
    $corteInfo[$cp] = @{
        categoria = "$($arrCorte[$r,12])"
        alocado   = [double]$arrCorte[$r,13]
    }
}
$wbCorte.Close($false)
Write-Host "  Produtos de corte: $($corteInfo.Count)"

# ---------------------------------------------------------------------------
# 3. Metas Julho.xlsx (FATURAMENTO e POSITIVACAO por Setor/Vendedor)
# ---------------------------------------------------------------------------
Write-Host "Lendo metas..."
function Read-MetaSheet($wb, $sheetName) {
    $ws = $wb.Worksheets.Item($sheetName)
    $rows = $ws.UsedRange.Rows.Count
    $arr = $ws.UsedRange.Value2
    $headerRow = 0
    for ($r = 1; $r -le [Math]::Min(5, $rows); $r++) {
        if ("$($arr[$r,1])" -eq "Setor") { $headerRow = $r; break }
    }
    $result = @{}
    for ($r = $headerRow + 1; $r -le $rows; $r++) {
        $setor = $arr[$r,1]
        $vendedor = "$($arr[$r,2])"
        if (-not $setor -or "$setor" -eq "") { continue }
        $setorStr = "$([int]$setor)"
        $result[$setorStr] = @{
            vendedor = $vendedor.Trim()
            WOW = [double]($arr[$r,3])
            PANASONIC = [double]($arr[$r,4])
            INGLEZA = [double]($arr[$r,5])
            CONDOR = [double]($arr[$r,6])
            BAGLEY = [double]($arr[$r,7])
            ARCOR = [double]($arr[$r,8])
            "AB MAURI" = [double]($arr[$r,9])
            geral = [double]($arr[$r,10])
        }
    }
    return $result
}
$supervisorCodesFixos = @("80","23","24","100","27","28")
function Read-MetaBlocos($wb, $sheetName) {
    # A planilha lista vendedores em blocos seguidos da linha-total do supervisor
    # (Setor = codigo do supervisor). Isso da o organograma canonico vendedor->supervisor,
    # mais confiavel que o COD.SUPERVISOR por linha da 8014 (que tem ruido/transferencias).
    $ws = $wb.Worksheets.Item($sheetName)
    $rows = $ws.UsedRange.Rows.Count
    $arr = $ws.UsedRange.Value2
    $headerRow = 0
    for ($r = 1; $r -le [Math]::Min(5, $rows); $r++) {
        if ("$($arr[$r,1])" -eq "Setor") { $headerRow = $r; break }
    }
    $canon = @{}
    $bloco = New-Object System.Collections.Generic.List[string]
    for ($r = $headerRow + 1; $r -le $rows; $r++) {
        $setor = $arr[$r,1]
        if (-not $setor -or "$setor" -eq "") { continue }
        $setorStr = "$([int]$setor)"
        if ($supervisorCodesFixos -contains $setorStr) {
            foreach ($cv in $bloco) { $canon[$cv] = $setorStr }
            $bloco.Clear()
        } else {
            $bloco.Add($setorStr)
        }
    }
    return $canon
}

$wbMeta = $excel.Workbooks.Open((Join-Path $basesPath "Metas Julho.xlsx"), $null, $true)
$metaFat = Read-MetaSheet $wbMeta "FATURAMENTO"
$metaPos = Read-MetaSheet $wbMeta "POSITIVAÇÃO"
$vendedorSupervisorCanonico = Read-MetaBlocos $wbMeta "FATURAMENTO"
$wbMeta.Close($false)
Write-Host "  Metas faturamento: $($metaFat.Count) setores | positivacao: $($metaPos.Count) setores"
Write-Host "  Organograma canonico (planilha de metas): $($vendedorSupervisorCanonico.Count) vendedores"

# ---------------------------------------------------------------------------
# 4. Estrutura de supervisores (fixa, confirmada na base 8014 - COD.SUPERVISOR)
# ---------------------------------------------------------------------------
# Codigos "dobrados" para dentro do supervisor real (mesmo padrao ja usado no
# projeto Acompanhamento Arcor para esta mesma equipe de vendas):
#   25 Denilson Rocha Guimaraes  -> 100 Washington (volume real)
#   14 Rodrigo Stelleo Costa Leite -> 27 Rodrigo (volume residual)
# Codigos internos/residuais (63,64,68,87) somados apenas no total geral, como "Outros".
$foldSupervisor = @{ "25" = "100"; "14" = "27" }
$supervisoresInfo = [ordered]@{
    "80"  = @{ nome = "Alessandro"; cor = "#008300" }
    "23"  = @{ nome = "Anderson";   cor = "#e87ba4" }
    "24"  = @{ nome = "Arildo";     cor = "#eda100" }
    "100" = @{ nome = "Washington"; cor = "#1baf7a" }
    "27"  = @{ nome = "Rodrigo";    cor = "#e34948" }
    "28"  = @{ nome = "Sueli";      cor = "#8a5a3b" }
}
$residuaisInterno = @("63","64","68","87")

# ---------------------------------------------------------------------------
# 5. Base 8014 - Geral (linhas de venda)
# ---------------------------------------------------------------------------
Write-Host "Lendo base 8014 - Geral (pode demorar)..."
$wb8 = $excel.Workbooks.Open((Join-Path $basesPath "8014 - Geral.xls"), $null, $true)
$ws8 = $wb8.Worksheets.Item(1)
$rows8 = $ws8.UsedRange.Rows.Count
$arr8 = $ws8.UsedRange.Value2
Write-Host "  $($rows8-1) linhas"

# Descobre o mes vigente (maior data encontrada) para filtrar a janela movel do export
$maxDate = [datetime]::MinValue
for ($r = 2; $r -le $rows8; $r++) {
    $v = $arr8[$r,1]
    if ($v -is [double]) {
        $d = [datetime]::FromOADate($v)
        if ($d -gt $maxDate) { $maxDate = $d }
    }
}
$primeiroDiaMes = Get-Date -Year $maxDate.Year -Month $maxDate.Month -Day 1
$ultimoDiaMes = $primeiroDiaMes.AddMonths(1).AddDays(-1)
Write-Host "  Periodo vigente: $($primeiroDiaMes.ToString('MM/yyyy')) (dados ate $($maxDate.ToString('dd/MM/yyyy')))"

# Pre-passada: alguns vendedores aparecem com COD.SUPERVISOR diferente entre linhas
# (transferencia no meio do mes / ruido de digitacao). Um vendedor so pode pertencer a
# UM supervisor no dashboard (senao a meta dele seria contada duas vezes); prioridade:
# 1) organograma da planilha de metas (canonico), 2) supervisor majoritario nas linhas.
Write-Host "Resolvendo supervisor de cada vendedor (votacao por maioria como fallback)..."
$votos = @{}
for ($r = 2; $r -le $rows8; $r++) {
    $cvv = $arr8[$r,15]
    $css = $arr8[$r,17]
    if (-not ($cvv -is [double]) -or -not ($css -is [double])) { continue }
    $cv = "$([int]$cvv)"
    $cs = "$([int]$css)"
    if ($foldSupervisor.ContainsKey($cs)) { $cs = $foldSupervisor[$cs] }
    if (-not $votos.ContainsKey($cv)) { $votos[$cv] = @{} }
    $votos[$cv][$cs] = ([int]$votos[$cv][$cs]) + 1
}
$vendedorSupervisorFinal = @{}
foreach ($cv in $votos.Keys) {
    if ($vendedorSupervisorCanonico.ContainsKey($cv)) {
        $vendedorSupervisorFinal[$cv] = $vendedorSupervisorCanonico[$cv]
    } else {
        $melhor = $null; $melhorCount = -1
        foreach ($cs in $votos[$cv].Keys) {
            if ($votos[$cv][$cs] -gt $melhorCount -and ($supervisoresInfo.Contains($cs))) { $melhor = $cs; $melhorCount = $votos[$cv][$cs] }
        }
        if ($melhor) { $vendedorSupervisorFinal[$cv] = $melhor }
    }
}
$divergentes = @($votos.Keys | Where-Object { $votos[$_].Count -gt 1 })
if ($divergentes.Count -gt 0) {
    Write-Host "  Vendedores com supervisor divergente entre linhas (resolvido por vendedor, nao por linha): $($divergentes -join ', ')"
}

Write-Host "Processando linhas..."
$linhas = New-Object System.Collections.Generic.List[object]

for ($r = 2; $r -le $rows8; $r++) {
    $vData = $arr8[$r,1]
    if (-not ($vData -is [double])) { continue }
    $data = [datetime]::FromOADate($vData)
    if ($data -lt $primeiroDiaMes -or $data -gt $ultimoDiaMes) { continue }

    $codCliente = "$($arr8[$r,2])"
    $nomeCliente = "$($arr8[$r,3])"
    $cnpj = "$($arr8[$r,4])"
    $segmento = "$($arr8[$r,5])"
    $numPedRCA = "$($arr8[$r,10])"
    $numPedWinthor = "$($arr8[$r,9])"
    $statusFat = Remove-Diacritics("$($arr8[$r,14])").ToUpper().Trim()
    $codVendedor = "$([int]$arr8[$r,15])"
    $nomeVendedor = "$($arr8[$r,16])"
    $codprod = "$($arr8[$r,22])"
    $descProduto = "$($arr8[$r,23])"
    $unidVendidas = [double]$arr8[$r,24]
    $valor = [double]$arr8[$r,29]
    $tipoVenda = Remove-Diacritics("$($arr8[$r,30])").ToUpper().Trim()

    $codSupervisor = "OUTROS"
    if ($vendedorSupervisorFinal.ContainsKey($codVendedor)) { $codSupervisor = $vendedorSupervisorFinal[$codVendedor] }

    $segTeam = "Varejo"
    if ($segmento.ToUpper().StartsWith("AS")) { $segTeam = "AS" }

    $forn = $null
    if ($codprodToForn.ContainsKey($codprod)) { $forn = $codprodToForn[$codprod] }

    $linhas.Add([pscustomobject]@{
        Data = $data
        CodCliente = $codCliente
        NomeCliente = $nomeCliente.Trim()
        CNPJ = $cnpj
        SegTeam = $segTeam
        NumPedRCA = $numPedRCA
        NumPedWinthor = $numPedWinthor
        StatusFat = $statusFat
        CodVendedor = $codVendedor
        NomeVendedor = $nomeVendedor.Trim()
        CodSupervisor = $codSupervisor
        CodProd = $codprod
        DescProduto = $descProduto.Trim()
        UnidVendidas = $unidVendidas
        Valor = $valor
        TipoVenda = $tipoVenda
        Fornecedor = $forn
    })
}
$wb8.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "  $($linhas.Count) linhas no periodo vigente"

# ---------------------------------------------------------------------------
# 6. Funcoes de agregacao
# ---------------------------------------------------------------------------
function New-Agg {
    $o = [ordered]@{
        metaFaturamento = 0.0
        metaPositivacao = 0.0
        faturado = 0.0
        aFaturar = 0.0
        devolucao = 0.0
        bonificacao = 0.0
        clientesVenda = New-Object System.Collections.Generic.HashSet[string]
        porFornecedor = [ordered]@{}
    }
    return $o
}
function Get-FornBucket($agg, $fid) {
    if (-not $agg.porFornecedor.Contains($fid)) {
        $agg.porFornecedor[$fid] = [ordered]@{
            metaFaturamento = 0.0; metaPositivacao = 0.0
            faturado = 0.0; aFaturar = 0.0; devolucao = 0.0; bonificacao = 0.0
            clientesVenda = New-Object System.Collections.Generic.HashSet[string]
        }
    }
    return $agg.porFornecedor[$fid]
}
function Add-Linha($agg, $linha) {
    if ($linha.Fornecedor) {
        $b = Get-FornBucket $agg $linha.Fornecedor
        if ($linha.TipoVenda -eq "VENDA") {
            if ($linha.StatusFat -eq "FATURADO") { $b.faturado += $linha.Valor } else { $b.aFaturar += $linha.Valor }
            [void]$b.clientesVenda.Add($linha.CNPJ)
        } elseif ($linha.TipoVenda -eq "DEVOLUCAO" -or $linha.TipoVenda -eq "DEVOLUÇÃO") {
            $b.devolucao += [Math]::Abs($linha.Valor)
        } elseif ($linha.TipoVenda -eq "BONIFICACAO" -or $linha.TipoVenda -eq "BONIFICAÇÃO") {
            $b.bonificacao += $linha.Valor
        }
    }
    if ($linha.TipoVenda -eq "VENDA") {
        if ($linha.StatusFat -eq "FATURADO") { $agg.faturado += $linha.Valor } else { $agg.aFaturar += $linha.Valor }
        [void]$agg.clientesVenda.Add($linha.CNPJ)
    } elseif ($linha.TipoVenda -eq "DEVOLUCAO" -or $linha.TipoVenda -eq "DEVOLUÇÃO") {
        $agg.devolucao += [Math]::Abs($linha.Valor)
    } elseif ($linha.TipoVenda -eq "BONIFICACAO" -or $linha.TipoVenda -eq "BONIFICAÇÃO") {
        $agg.bonificacao += $linha.Valor
    }
}
function ConvertTo-JsonAgg($agg) {
    $out = [ordered]@{
        metaFaturamento = [Math]::Round($agg.metaFaturamento,2)
        metaPositivacao = $agg.metaPositivacao
        faturadoLiquido = [Math]::Round($agg.faturado - $agg.devolucao,2)
        aFaturar = [Math]::Round($agg.aFaturar,2)
        devolucao = [Math]::Round($agg.devolucao,2)
        bonificacao = [Math]::Round($agg.bonificacao,2)
        positivacaoRealizado = $agg.clientesVenda.Count
        porFornecedor = [ordered]@{}
    }
    foreach ($fid in $agg.porFornecedor.Keys) {
        $b = $agg.porFornecedor[$fid]
        $out.porFornecedor[$fid] = [ordered]@{
            metaFaturamento = [Math]::Round($b.metaFaturamento,2)
            metaPositivacao = $b.metaPositivacao
            faturadoLiquido = [Math]::Round($b.faturado - $b.devolucao,2)
            aFaturar = [Math]::Round($b.aFaturar,2)
            devolucao = [Math]::Round($b.devolucao,2)
            bonificacao = [Math]::Round($b.bonificacao,2)
            positivacaoRealizado = $b.clientesVenda.Count
        }
    }
    return $out
}

Write-Host "Agregando totais..."
$totalGeral = New-Agg
$geralAS = New-Agg
$geralVarejo = New-Agg
$supAgg = @{}
foreach ($cs in $supervisoresInfo.Keys) { $supAgg[$cs] = @{ geral = New-Agg; AS = New-Agg; Varejo = New-Agg; vendedores = @{} } }
$supAgg["OUTROS"] = @{ geral = New-Agg; AS = New-Agg; Varejo = New-Agg; vendedores = @{} }

foreach ($l in $linhas) {
    Add-Linha $totalGeral $l
    if ($l.SegTeam -eq "AS") { Add-Linha $geralAS $l } else { Add-Linha $geralVarejo $l }

    $cs = $l.CodSupervisor
    Add-Linha $supAgg[$cs].geral $l
    Add-Linha $supAgg[$cs][$l.SegTeam] $l

    if (-not $supAgg[$cs].vendedores.ContainsKey($l.CodVendedor)) {
        $supAgg[$cs].vendedores[$l.CodVendedor] = @{
            nome = $l.NomeVendedor
            geral = New-Agg
            pedidos = New-Object 'System.Collections.Generic.Dictionary[string,object]'
        }
    }
    $vAgg = $supAgg[$cs].vendedores[$l.CodVendedor]
    Add-Linha $vAgg.geral $l

    if ($l.TipoVenda -eq "VENDA") {
        $pedKey = "$($l.NumPedWinthor)|$($l.StatusFat)"
        if (-not $vAgg.pedidos.ContainsKey($pedKey)) {
            $vAgg.pedidos[$pedKey] = [ordered]@{
                numeroPedido = $l.NumPedWinthor
                numeroPedidoRCA = $l.NumPedRCA
                data = $l.Data.ToString("yyyy-MM-dd")
                codCliente = $l.CodCliente
                razaoSocial = $l.NomeCliente
                status = $l.StatusFat
                valor = 0.0
            }
        }
        $vAgg.pedidos[$pedKey].valor = [Math]::Round($vAgg.pedidos[$pedKey].valor + $l.Valor, 2)
    }
}

# Aplica metas (faturamento e positivacao) por vendedor e soma nos supervisores/geral/segmentos
# Nota: os totais de segmento (AS/Varejo) nao tem meta propria na planilha de metas (a meta e
# por vendedor, sem quebra por segmento do cliente) - por isso a meta so e somada nos niveis
# vendedor / supervisor / total geral, nunca em geralAS/geralVarejo (que ficam so com realizado).
foreach ($cs in $supAgg.Keys) {
    foreach ($cv in $supAgg[$cs].vendedores.Keys) {
        $vAgg = $supAgg[$cs].vendedores[$cv]
        if ($metaFat.ContainsKey($cv)) {
            $mf = $metaFat[$cv]; $mp = $metaPos[$cv]
            $vAgg.geral.metaFaturamento = $mf.geral
            $vAgg.geral.metaPositivacao = $mp.geral
            foreach ($fid in $fornecedores.Keys) {
                $b = Get-FornBucket $vAgg.geral $fid
                $b.metaFaturamento = $mf[$fid]
                $b.metaPositivacao = $mp[$fid]
            }
            $totalGeral.metaFaturamento += $mf.geral
            $totalGeral.metaPositivacao += $mp.geral
            $supAgg[$cs].geral.metaFaturamento += $mf.geral
            $supAgg[$cs].geral.metaPositivacao += $mp.geral
            foreach ($fid in $fornecedores.Keys) {
                (Get-FornBucket $totalGeral $fid).metaFaturamento += $mf[$fid]
                (Get-FornBucket $totalGeral $fid).metaPositivacao += $mp[$fid]
                (Get-FornBucket $supAgg[$cs].geral $fid).metaFaturamento += $mf[$fid]
                (Get-FornBucket $supAgg[$cs].geral $fid).metaPositivacao += $mp[$fid]
            }
        }
    }
}

# ---------------------------------------------------------------------------
# 7. Devolucoes (aba dedicada, por RCA/supervisor)
# ---------------------------------------------------------------------------
Write-Host "Montando devolucoes..."
$devLinhas = $linhas | Where-Object { $_.TipoVenda -eq "DEVOLUCAO" -or $_.TipoVenda -eq "DEVOLUÇÃO" }
$devPorSup = @{}
foreach ($l in $devLinhas) {
    $cs = $l.CodSupervisor
    if (-not $devPorSup.ContainsKey($cs)) { $devPorSup[$cs] = New-Object System.Collections.Generic.List[object] }
    $devPorSup[$cs].Add([ordered]@{
        data = $l.Data.ToString("yyyy-MM-dd")
        numeroPedido = $l.NumPedRCA
        codCliente = $l.CodCliente
        razaoSocial = $l.NomeCliente
        produto = $l.DescProduto
        fornecedor = $l.Fornecedor
        vendedor = $l.NomeVendedor
        valor = [Math]::Round([Math]::Abs($l.Valor),2)
    })
}

# ---------------------------------------------------------------------------
# 8. Cortes (produtos de alocacao controlada)
# ---------------------------------------------------------------------------
Write-Host "Montando cortes..."
$corteLinhas = $linhas | Where-Object { $_.TipoVenda -eq "VENDA" -and $corteInfo.ContainsKey($_.CodProd) }
$cortePorSup = @{}
$corteConsumoProduto = @{}
foreach ($l in $corteLinhas) {
    $cs = $l.CodSupervisor
    if (-not $cortePorSup.ContainsKey($cs)) { $cortePorSup[$cs] = New-Object System.Collections.Generic.List[object] }
    $cortePorSup[$cs].Add([ordered]@{
        data = $l.Data.ToString("yyyy-MM-dd")
        numeroPedido = $l.NumPedRCA
        codCliente = $l.CodCliente
        razaoSocial = $l.NomeCliente
        produto = $l.DescProduto
        fornecedor = $l.Fornecedor
        vendedor = $l.NomeVendedor
        status = $l.StatusFat
        valor = [Math]::Round($l.Valor,2)
    })
    if (-not $corteConsumoProduto.ContainsKey($l.CodProd)) {
        $ci = $corteInfo[$l.CodProd]
        $corteConsumoProduto[$l.CodProd] = [ordered]@{
            produto = $l.DescProduto
            categoria = $ci.categoria
            fornecedor = $l.Fornecedor
            alocado = $ci.alocado
            vendidoUnid = 0.0
            valor = 0.0
        }
    }
    $corteConsumoProduto[$l.CodProd].valor = [Math]::Round($corteConsumoProduto[$l.CodProd].valor + $l.Valor,2)
    $corteConsumoProduto[$l.CodProd].vendidoUnid += $l.UnidVendidas
}

# ---------------------------------------------------------------------------
# 9. Periodo / dias uteis (para o Ritmo)
# ---------------------------------------------------------------------------
$hoje = $maxDate.Date
$diasUteisRestantes = 0
$diasUteisTotais = 0
$diasUteisPassados = 0
for ($d = $primeiroDiaMes; $d -le $ultimoDiaMes; $d = $d.AddDays(1)) {
    if ($d.DayOfWeek -ne "Sunday") {
        $diasUteisTotais++
        if ($d -le $hoje) { $diasUteisPassados++ }
        if ($d -gt $hoje) { $diasUteisRestantes++ }
    }
}

Write-Host "Serializando JSON..."

# Monta a arvore de supervisores/vendedores para o JSON final
$supervisoresOut = New-Object System.Collections.Generic.List[object]
foreach ($cs in $supervisoresInfo.Keys) {
    $info = $supervisoresInfo[$cs]
    $sAgg = $supAgg[$cs]
    $vendOut = New-Object System.Collections.Generic.List[object]
    foreach ($cv in ($sAgg.vendedores.Keys | Sort-Object { $sAgg.vendedores[$_].geral.faturado + $sAgg.vendedores[$_].geral.aFaturar } -Descending)) {
        $vAgg = $sAgg.vendedores[$cv]
        $pedidosArr = @($vAgg.pedidos.Values | Sort-Object data -Descending)
        $vendOut.Add([ordered]@{
            codigo = $cv
            nome = $vAgg.nome
            geral = ConvertTo-JsonAgg $vAgg.geral
            pedidos = $pedidosArr
        })
    }
    $supervisoresOut.Add([ordered]@{
        codigo = $cs
        nome = $info.nome
        cor = $info.cor
        geral = ConvertTo-JsonAgg $sAgg.geral
        segmentos = [ordered]@{
            AS = ConvertTo-JsonAgg $sAgg.AS
            Varejo = ConvertTo-JsonAgg $sAgg.Varejo
        }
        vendedores = $vendOut
    })
}

$fornecedoresOut = New-Object System.Collections.Generic.List[object]
foreach ($fid in $fornecedores.Keys) {
    $fornecedoresOut.Add([ordered]@{ id = $fid; nome = $fornecedores[$fid].nome; cor = $fornecedores[$fid].cor })
}

$devolucoesOut = [ordered]@{
    valorTotal = [Math]::Round(($devLinhas | Measure-Object -Property Valor -Sum).Sum,2) | ForEach-Object { [Math]::Abs($_) }
    porSupervisor = New-Object System.Collections.Generic.List[object]
}
$devolucoesOut.valorTotal = [Math]::Abs([Math]::Round(($devLinhas | Measure-Object -Property Valor -Sum).Sum,2))
foreach ($cs in $devPorSup.Keys) {
    $nomeSup = if ($supervisoresInfo.Contains($cs)) { $supervisoresInfo[$cs].nome } else { "Outros/Interno" }
    $itens = @($devPorSup[$cs] | Sort-Object data -Descending)
    $devolucoesOut.porSupervisor.Add([ordered]@{
        codigo = $cs
        nome = $nomeSup
        valorTotal = [Math]::Round((($itens | Measure-Object -Property valor -Sum).Sum),2)
        itens = $itens
    })
}

$cortesOut = [ordered]@{
    valorTotal = [Math]::Round(($corteLinhas | Measure-Object -Property Valor -Sum).Sum,2)
    produtos = @($corteConsumoProduto.Values | Sort-Object valor -Descending)
    porSupervisor = New-Object System.Collections.Generic.List[object]
}
foreach ($cs in $cortePorSup.Keys) {
    $nomeSup = if ($supervisoresInfo.Contains($cs)) { $supervisoresInfo[$cs].nome } else { "Outros/Interno" }
    $itens = @($cortePorSup[$cs] | Sort-Object data -Descending)
    $cortesOut.porSupervisor.Add([ordered]@{
        codigo = $cs
        nome = $nomeSup
        valorTotal = [Math]::Round((($itens | Measure-Object -Property valor -Sum).Sum),2)
        itens = $itens
    })
}

$outrosValor = ($linhas | Where-Object { -not $_.Fornecedor -and $_.TipoVenda -eq "VENDA" } | Measure-Object -Property Valor -Sum).Sum
if (-not $outrosValor) { $outrosValor = 0 }

$dataJson = [ordered]@{
    geradoEm = (Get-Date).ToString("yyyy-MM-dd HH:mm")
    periodo = [ordered]@{
        mes = $primeiroDiaMes.Month
        ano = $primeiroDiaMes.Year
        label = (Get-Culture).TextInfo.ToTitleCase($primeiroDiaMes.ToString("MMMM/yyyy", [Globalization.CultureInfo]::GetCultureInfo("pt-BR")))
        ultimaAtualizacao = $maxDate.ToString("dd/MM/yyyy")
        diasUteisTotais = $diasUteisTotais
        diasUteisPassados = $diasUteisPassados
        diasUteisRestantes = $diasUteisRestantes
    }
    fornecedores = $fornecedoresOut
    totalGeral = ConvertTo-JsonAgg $totalGeral
    geralAS = ConvertTo-JsonAgg $geralAS
    geralVarejo = ConvertTo-JsonAgg $geralVarejo
    supervisores = $supervisoresOut
    devolucoes = $devolucoesOut
    cortes = $cortesOut
    outrosItensValor = [Math]::Round($outrosValor,2)
}

$json = $dataJson | ConvertTo-Json -Depth 12 -Compress
Set-Content -Path (Join-Path $sitePath "data.json") -Value $json -Encoding UTF8

# Alem do .json (util para depuracao), gera um .js com os dados embutidos como variavel
# global. Vendedores/supervisores abrem o HTML clicando duas vezes (file://), e navegadores
# bloqueiam fetch() de arquivo local por CORS - um <script src="data.js"> nao tem essa
# restricao, entao o site funciona tanto local quanto hospedado num servidor.
Set-Content -Path (Join-Path $sitePath "data.js") -Value "window.DASHBOARD_DATA = $json;" -Encoding UTF8
Write-Host "OK: Site/data.json e Site/data.js gerados ($([Math]::Round((Get-Item (Join-Path $sitePath 'data.json')).Length/1KB,1)) KB)"

Write-Host "Concluido."
