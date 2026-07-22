# Gera Site/data.json e Ritmo/data.json a partir das bases em Bases/
# Reexecutar sempre que uma base for atualizada (Excel COM trava os .xls entre execucoes;
# se der erro de acesso, feche instancias fantasmas: taskkill /F /IM EXCEL.EXE)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$basesPath = Join-Path $root "Bases"
$sitePath = Join-Path $root "Site"
$ritmoPath = Join-Path $root "Ritmo"
$consultaPath = Join-Path $root "Consulta"

# Usuario reorganizou a pasta Bases em 2026-07-22 (depois de compartilhar upload com
# terceiro): arquivos que mudam pouco (Metas, catalogo ESTOQUE API) foram movidos para
# uma subpasta "Armazenamento", deixando a raiz de Bases so com o que muda com frequencia
# (8022 - geral.xls, Cortes Geral.xls). Resolve-BaseFile procura primeiro na raiz, depois
# em Armazenamento, pra pipeline funcionar nas duas organizacoes sem precisar mover nada.
function Resolve-BaseFile([string]$nomeArquivo) {
    $naRaiz = Join-Path $basesPath $nomeArquivo
    if (Test-Path $naRaiz) { return $naRaiz }
    $noArmazenamento = Join-Path (Join-Path $basesPath "Armazenamento") $nomeArquivo
    if (Test-Path $noArmazenamento) { return $noArmazenamento }
    throw "Arquivo nao encontrado em Bases/ nem em Bases/Armazenamento/: $nomeArquivo"
}

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

# Na base 8022 as colunas VALOR R$ NF e UNIDADES VENDIDAS vem como TEXTO formatado em
# pt-BR ("1.251,00"), nao como numero (Value2 retorna string em vez de double para
# praticamente todas as linhas) - diferente da 8014, onde essas colunas eram numericas.
# [double]"1.251,00" falha (o parser nao sabe que "." e separador de milhar aqui), entao
# converte manualmente antes de castar.
function Parse-ValorBR($v) {
    if ($v -is [double]) { return $v }
    if (-not $v -or "$v" -eq "") { return 0.0 }
    $s = "$v".Trim().Replace(".", "").Replace(",", ".")
    $out = 0.0
    if ([double]::TryParse($s, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$out)) { return $out }
    return 0.0
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
$wbCat = $excel.Workbooks.Open((Resolve-BaseFile "ESTOQUE API (1).xls"), $null, $true)
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
# 2. (era o catalogo estatico "Cortes - Geral.xls" - substituido em 2026-07-22 pelo
# relatorio real de corte "Cortes Geral.xls", lido mais abaixo na secao 8 depois da
# 8022, pois precisa do lookup pedido->cliente construido a partir dela.)
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

$wbMeta = $excel.Workbooks.Open((Resolve-BaseFile "Metas Julho.xlsx"), $null, $true)
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
# 5. Base 8022 - Geral (linhas de venda) - substituiu a 8014 em 2026-07-21.
# Mesmas informacoes da 8014 + STATUS PEDIDO (igual ao antigo "A FATURAR") e STATUS
# BLOQUEIO (novo - status operacional do pedido: PENDENTE/LIBERADO/MONTADO/FATURADO/
# BLOQUEADO/DEVOLVIDO). Colunas deslocadas em relacao a 8014: tem "FILIAL" a mais no
# inicio (+1) e "STATUS BLOQUEIO" a mais entre ORIGEM_PEDIDO e COD.VENDEDOR (+1), e
# CAIXAS VENDIDAS/UNIDADES VENDIDAS vem trocadas de ordem entre si.
# ---------------------------------------------------------------------------
Write-Host "Lendo base 8022 - Geral (pode demorar)..."
$wb8 = $excel.Workbooks.Open((Resolve-BaseFile "8022 - geral.xls"), $null, $true)
$ws8 = $wb8.Worksheets.Item(1)
$rows8 = $ws8.UsedRange.Rows.Count
$arr8 = $ws8.UsedRange.Value2
Write-Host "  $($rows8-1) linhas"

# Descobre o mes vigente (maior data encontrada) para filtrar a janela movel do export
$maxDate = [datetime]::MinValue
for ($r = 2; $r -le $rows8; $r++) {
    $v = $arr8[$r,2]
    if ($v -is [double]) {
        $d = [datetime]::FromOADate($v)
        if ($d -gt $maxDate) { $maxDate = $d }
    }
}
$primeiroDiaMes = (Get-Date -Year $maxDate.Year -Month $maxDate.Month -Day 1).Date
$ultimoDiaMes = $primeiroDiaMes.AddMonths(1).AddDays(-1)
Write-Host "  Periodo vigente: $($primeiroDiaMes.ToString('MM/yyyy')) (dados ate $($maxDate.ToString('dd/MM/yyyy')))"

# Pre-passada: alguns vendedores aparecem com COD.SUPERVISOR diferente entre linhas
# (transferencia no meio do mes / ruido de digitacao). Um vendedor so pode pertencer a
# UM supervisor no dashboard (senao a meta dele seria contada duas vezes); prioridade:
# 1) organograma da planilha de metas (canonico), 2) supervisor majoritario nas linhas.
Write-Host "Resolvendo supervisor de cada vendedor (votacao por maioria como fallback)..."
$votos = @{}
for ($r = 2; $r -le $rows8; $r++) {
    $cvv = $arr8[$r,17]
    $css = $arr8[$r,19]
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
    $vData = $arr8[$r,2]
    if (-not ($vData -is [double])) { continue }
    $data = [datetime]::FromOADate($vData)
    if ($data -lt $primeiroDiaMes -or $data -gt $ultimoDiaMes) { continue }

    $codCliente = "$($arr8[$r,3])"
    $nomeCliente = "$($arr8[$r,4])"
    $cnpj = "$($arr8[$r,5])"
    $segmento = "$($arr8[$r,6])"
    $numPedRCA = "$($arr8[$r,11])"
    $numPedWinthor = "$($arr8[$r,10])"
    $statusFat = Remove-Diacritics("$($arr8[$r,15])").ToUpper().Trim()
    $statusBloqueio = Remove-Diacritics("$($arr8[$r,16])").ToUpper().Trim()
    $codVendedor = "$([int]$arr8[$r,17])"
    $nomeVendedor = "$($arr8[$r,18])"
    $codprod = "$($arr8[$r,24])"
    $descProduto = "$($arr8[$r,25])"
    $unidVendidas = Parse-ValorBR $arr8[$r,27]
    $valor = Parse-ValorBR $arr8[$r,31]
    $tipoVenda = Remove-Diacritics("$($arr8[$r,32])").ToUpper().Trim()

    # Bucketing de vendas/devolucao/corte usa o supervisor DA PROPRIA LINHA (so com os folds
    # numericos ja conhecidos: Denilson->Washington, Rodrigo Stelleo->Rodrigo) - nao o
    # "canonico" por vendedor. Motivo (achado 2026-07-21): vendedor 260 tem TODAS as devolucoes
    # do mes marcadas com supervisor=Arildo na base, mas a maioria das vendas marcadas com
    # Alessandro (parece transferencia no meio do mes) - usar o canonico pra bucketing jogava
    # a devolucao inteira pro Alessandro, R$14.519,53 a mais do que um filtro direto na base
    # mostra. Um filtro simples (supervisor = X, tipo = devolucao) precisa bater exatamente com
    # o que o dashboard soma. O canonico (vendedorSupervisorFinal) continua existindo e e usado
    # só para decidir quem recebe a META do vendedor (evita o double-count que motivou criar
    # o canonico originalmente) - ver o loop de metas mais abaixo.
    $codSupervisorRawVal = $arr8[$r,19]
    $codSupervisor = "OUTROS"
    if ($codSupervisorRawVal -is [double]) {
        $csFolded = "$([int]$codSupervisorRawVal)"
        if ($foldSupervisor.ContainsKey($csFolded)) { $csFolded = $foldSupervisor[$csFolded] }
        if ($supervisoresInfo.Contains($csFolded)) { $codSupervisor = $csFolded }
    }

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
        StatusBloqueio = $statusBloqueio
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

# Lookup Pedido Winthor -> dados do cliente, usado para enriquecer o relatorio de Cortes
# (que nao tem CNPJ/cod.cliente/razao social - so pedido, produto e RCA). Um pedido Winthor
# pertence a um unico cliente, entao o primeiro encontrado ja resolve.
$pedidoWinthorToCliente = @{}
foreach ($l in $linhas) {
    if (-not $pedidoWinthorToCliente.ContainsKey($l.NumPedWinthor)) {
        $pedidoWinthorToCliente[$l.NumPedWinthor] = @{ CodCliente = $l.CodCliente; CNPJ = $l.CNPJ; NomeCliente = $l.NomeCliente }
    }
}

# ---------------------------------------------------------------------------
# 5b. Bases 8022 de meses anteriores (ex.: "8022 - Jun.xls") - so para a Consulta
# de Pedidos. O dashboard/devolucoes/cortes continuam restritos ao mes vigente
# (o que esta em "8022 - geral.xls"), mas o vendedor precisa achar pedidos
# faturados em meses anteriores na Consulta - motivo pelo qual o usuario pediu
# para incluir essas bases extras (2026-07-22). Mesmo layout de colunas do
# "8022 - geral.xls" (confirmado pelo usuario). Qualquer arquivo "8022 - *.xls"
# (exceto o "geral") na pasta Bases entra automaticamente aqui, para nao precisar
# tocar no pipeline de novo no proximo mes historico que for adicionado.
#
# Descoberta ao ligar (2026-07-22): "8022 - Jun.xls" nao e um export limpo so de
# Junho - e uma janela movel que mistura Junho E Julho (33903 linhas de 06/2026 +
# 31632 de 07/2026 nesse arquivo). Detectar "o mes da data maxima" (como se faz
# pra base principal) pegaria Julho de novo - duplicando o que "8022 - geral.xls"
# ja cobre (e esse e' o mais atualizado pro mes vigente, roda toda hora). Por
# isso o corte aqui e' fixo: mantém so linhas ANTERIORES ao mes vigente
# ($primeiroDiaMes, calculado a partir da base principal la em cima) - pega
# Junho (e qualquer outro mes mais antigo que apareca em outro arquivo historico)
# sem duplicar Julho.
# ---------------------------------------------------------------------------
$linhasHistorico = New-Object System.Collections.Generic.List[object]
$armazenamentoPath = Join-Path $basesPath "Armazenamento"
$arquivosHistorico = @(Get-ChildItem -Path $basesPath -Filter "8022 - *.xls" -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne "8022 - geral.xls" })
if (Test-Path $armazenamentoPath) {
    $arquivosHistorico += @(Get-ChildItem -Path $armazenamentoPath -Filter "8022 - *.xls" -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne "8022 - geral.xls" })
}
foreach ($arqH in $arquivosHistorico) {
    Write-Host "Lendo base historica $($arqH.Name) (so para Consulta de Pedidos)..."
    $excelH = New-Object -ComObject Excel.Application
    $excelH.Visible = $false
    $excelH.DisplayAlerts = $false
    $wbH = $excelH.Workbooks.Open($arqH.FullName, $null, $true)
    $wsH = $wbH.Worksheets.Item(1)
    $rowsH = $wsH.UsedRange.Rows.Count
    $arrH = $wsH.UsedRange.Value2

    $lidasH = 0
    for ($r = 2; $r -le $rowsH; $r++) {
        $vData = $arrH[$r,2]
        if (-not ($vData -is [double])) { continue }
        $data = [datetime]::FromOADate($vData)
        if ($data -ge $primeiroDiaMes) { continue }

        $codSupervisorRawValH = $arrH[$r,19]
        $codSupervisorH = "OUTROS"
        if ($codSupervisorRawValH -is [double]) {
            $csFoldedH = "$([int]$codSupervisorRawValH)"
            if ($foldSupervisor.ContainsKey($csFoldedH)) { $csFoldedH = $foldSupervisor[$csFoldedH] }
            if ($supervisoresInfo.Contains($csFoldedH)) { $codSupervisorH = $csFoldedH }
        }
        $codprodH = "$($arrH[$r,24])"
        $fornH = $null
        if ($codprodToForn.ContainsKey($codprodH)) { $fornH = $codprodToForn[$codprodH] }

        $linhasHistorico.Add([pscustomobject]@{
            Data = $data
            CodCliente = "$($arrH[$r,3])"
            NomeCliente = "$($arrH[$r,4])".Trim()
            CNPJ = "$($arrH[$r,5])"
            SegTeam = $null
            NumPedRCA = "$($arrH[$r,11])"
            NumPedWinthor = "$($arrH[$r,10])"
            StatusFat = Remove-Diacritics("$($arrH[$r,15])").ToUpper().Trim()
            StatusBloqueio = Remove-Diacritics("$($arrH[$r,16])").ToUpper().Trim()
            CodVendedor = "$([int]$arrH[$r,17])"
            NomeVendedor = "$($arrH[$r,18])".Trim()
            CodSupervisor = $codSupervisorH
            CodProd = $codprodH
            DescProduto = "$($arrH[$r,25])".Trim()
            UnidVendidas = Parse-ValorBR $arrH[$r,27]
            Valor = Parse-ValorBR $arrH[$r,31]
            TipoVenda = Remove-Diacritics("$($arrH[$r,32])").ToUpper().Trim()
            Fornecedor = $fornH
        })
        $lidasH++
    }
    $wbH.Close($false)
    $excelH.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelH) | Out-Null
    Write-Host "  $lidasH linhas lidas (anteriores a $($primeiroDiaMes.ToString('MM/yyyy')))"
}

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
        clientesFaturado = New-Object System.Collections.Generic.HashSet[string]
        clientesAFaturar = New-Object System.Collections.Generic.HashSet[string]
        clientesDevolucao = New-Object System.Collections.Generic.HashSet[string]
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
            clientesFaturado = New-Object System.Collections.Generic.HashSet[string]
            clientesAFaturar = New-Object System.Collections.Generic.HashSet[string]
            clientesDevolucao = New-Object System.Collections.Generic.HashSet[string]
        }
    }
    return $agg.porFornecedor[$fid]
}
function Add-Linha($agg, $linha) {
    if ($linha.Fornecedor) {
        $b = Get-FornBucket $agg $linha.Fornecedor
        if ($linha.TipoVenda -eq "VENDA") {
            if ($linha.StatusFat -eq "FATURADO") { $b.faturado += $linha.Valor; [void]$b.clientesFaturado.Add($linha.CNPJ) }
            else { $b.aFaturar += $linha.Valor; [void]$b.clientesAFaturar.Add($linha.CNPJ) }
            [void]$b.clientesVenda.Add($linha.CNPJ)
        } elseif ($linha.TipoVenda -eq "DEVOLUCAO" -or $linha.TipoVenda -eq "DEVOLUÇÃO") {
            $b.devolucao += [Math]::Abs($linha.Valor)
            [void]$b.clientesDevolucao.Add($linha.CNPJ)
        } elseif ($linha.TipoVenda -eq "BONIFICACAO" -or $linha.TipoVenda -eq "BONIFICAÇÃO") {
            $b.bonificacao += $linha.Valor
        }
    }
    if ($linha.TipoVenda -eq "VENDA") {
        if ($linha.StatusFat -eq "FATURADO") { $agg.faturado += $linha.Valor; [void]$agg.clientesFaturado.Add($linha.CNPJ) }
        else { $agg.aFaturar += $linha.Valor; [void]$agg.clientesAFaturar.Add($linha.CNPJ) }
        [void]$agg.clientesVenda.Add($linha.CNPJ)
    } elseif ($linha.TipoVenda -eq "DEVOLUCAO" -or $linha.TipoVenda -eq "DEVOLUÇÃO") {
        $agg.devolucao += [Math]::Abs($linha.Valor)
        [void]$agg.clientesDevolucao.Add($linha.CNPJ)
    } elseif ($linha.TipoVenda -eq "BONIFICACAO" -or $linha.TipoVenda -eq "BONIFICAÇÃO") {
        $agg.bonificacao += $linha.Valor
    }
}
# Positivacao = clientes com venda no periodo QUE NAO tiveram nenhuma devolucao no mesmo
# escopo (fornecedor ou geral) - decisao explicita do usuario 2026-07-21: uma devolucao,
# mesmo parcial, tira o cliente da contagem de positivados (nao so quando o liquido fica <=0).
# Quebrado em faturado/a-faturar (um cliente pode entrar nos dois se tiver pedidos nos dois
# status) para o usuario ver a composicao, nao so o total.
function Count-Set($set, $excluir) {
    $count = 0
    foreach ($cnpj in $set) { if (-not $excluir.Contains($cnpj)) { $count++ } }
    return $count
}
function Count-Positivados($agg) { return Count-Set $agg.clientesVenda $agg.clientesDevolucao }
function ConvertTo-JsonAgg($agg) {
    $out = [ordered]@{
        metaFaturamento = [Math]::Round($agg.metaFaturamento,2)
        metaPositivacao = [Math]::Round($agg.metaPositivacao)
        faturadoLiquido = [Math]::Round($agg.faturado - $agg.devolucao,2)
        aFaturar = [Math]::Round($agg.aFaturar,2)
        devolucao = [Math]::Round($agg.devolucao,2)
        bonificacao = [Math]::Round($agg.bonificacao,2)
        positivacaoRealizado = Count-Positivados $agg
        positivacaoFaturado = Count-Set $agg.clientesFaturado $agg.clientesDevolucao
        positivacaoAFaturar = Count-Set $agg.clientesAFaturar $agg.clientesDevolucao
        clientesComDevolucao = $agg.clientesDevolucao.Count
        porFornecedor = [ordered]@{}
    }
    foreach ($fid in $agg.porFornecedor.Keys) {
        $b = $agg.porFornecedor[$fid]
        $out.porFornecedor[$fid] = [ordered]@{
            metaFaturamento = [Math]::Round($b.metaFaturamento,2)
            metaPositivacao = [Math]::Round($b.metaPositivacao)
            faturadoLiquido = [Math]::Round($b.faturado - $b.devolucao,2)
            aFaturar = [Math]::Round($b.aFaturar,2)
            devolucao = [Math]::Round($b.devolucao,2)
            bonificacao = [Math]::Round($b.bonificacao,2)
            positivacaoRealizado = Count-Positivados $b
            positivacaoFaturado = Count-Set $b.clientesFaturado $b.clientesDevolucao
            positivacaoAFaturar = Count-Set $b.clientesAFaturar $b.clientesDevolucao
            clientesComDevolucao = $b.clientesDevolucao.Count
        }
    }
    return $out
}

Write-Host "Agregando totais..."
$totalGeral = New-Agg
$supAgg = @{}
foreach ($cs in $supervisoresInfo.Keys) { $supAgg[$cs] = @{ geral = New-Agg; vendedores = @{} } }
$supAgg["OUTROS"] = @{ geral = New-Agg; vendedores = @{} }

foreach ($l in $linhas) {
    Add-Linha $totalGeral $l

    $cs = $l.CodSupervisor
    Add-Linha $supAgg[$cs].geral $l

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
                numeroPedidoWinthor = $l.NumPedWinthor
                numeroPedidoClube = $l.NumPedRCA
                codigoRCA = $l.CodVendedor
                data = $l.Data.ToString("yyyy-MM-dd")
                codCliente = $l.CodCliente
                cnpj = $l.CNPJ
                razaoSocial = $l.NomeCliente
                status = $l.StatusFat
                statusBloqueio = $l.StatusBloqueio
                valor = 0.0
            }
        }
        $vAgg.pedidos[$pedKey].valor = [Math]::Round($vAgg.pedidos[$pedKey].valor + $l.Valor, 2)
    }
}

# Rede de seguranca: garante que todo vendedor com meta tenha uma entrada no bucket do seu
# supervisor CANONICO, mesmo se nenhuma linha dele no periodo tiver sido marcada com esse
# supervisor na base (caso raro, mas sem isso a meta dele seria perdida silenciosamente em
# vez de somada - pior que double-count).
foreach ($cv in $metaFat.Keys) {
    if (-not $vendedorSupervisorFinal.ContainsKey($cv)) { continue }
    $csCanon = $vendedorSupervisorFinal[$cv]
    if (-not $supAgg.ContainsKey($csCanon)) { continue }
    if (-not $supAgg[$csCanon].vendedores.ContainsKey($cv)) {
        $supAgg[$csCanon].vendedores[$cv] = @{
            nome = $metaFat[$cv].vendedor
            geral = New-Agg
            pedidos = New-Object 'System.Collections.Generic.Dictionary[string,object]'
        }
    }
}

# Aplica metas (faturamento e positivacao) por vendedor e soma nos supervisores/geral.
# So no bucket do supervisor CANONICO do vendedor (nao em todo bucket onde ele aparece) -
# um vendedor com linhas divididas entre dois supervisores (ver nota acima) pode ter uma
# entrada em cada supAgg[cs].vendedores, mas a meta so pode ser somada uma vez.
foreach ($cs in $supAgg.Keys) {
    foreach ($cv in $supAgg[$cs].vendedores.Keys) {
        $vAgg = $supAgg[$cs].vendedores[$cv]
        $ehCanonico = $vendedorSupervisorFinal.ContainsKey($cv) -and $vendedorSupervisorFinal[$cv] -eq $cs
        if ($ehCanonico -and $metaFat.ContainsKey($cv)) {
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
# Measure-Object -Property NAO funciona em [ordered]@{} (hashtable) - so em PSCustomObject/
# propriedades reais de .NET. Os grupos de devolucoes/cortes sao hashtables (para controlar
# a ordem das chaves no JSON), entao toda soma de valorTotal precisa ser manual (Sum-GrupoValor,
# mais abaixo). Bug real encontrado 2026-07-21: todo "valorTotal" por supervisor de Devolucoes
# e Cortes estava sempre 0 desde a primeira versao do site por causa disso (Measure-Object
# falhava silenciosamente, sem erro).

# Devolucoes e cortes sao agrupados por cliente+pedido (nao mais uma linha por produto):
# cada grupo tem os dados do cliente/pedido/vendedor uma vez so, e uma lista "produtos"
# com os itens que compoem aquele pedido (cada um com seu proprio codigo de produto).
# Terminologia (corrigida 2026-07-21 apos confirmacao do usuario):
#   - "Codigo RCA" = codigo do VENDEDOR (nao numero de pedido) - o mesmo codigo que aparece
#     na coluna "Setor" da planilha de metas, ao lado do nome do vendedor.
#   - "Pedido Winthor" (NUMERO PED. WINTHOR) e "Pedido Clube" (NUMERO PED. RCA, apesar do
#     nome da coluna na base ter "RCA" - e o pedido do sistema do "Clube da Venda", numeracao
#     independente da do Winthor) sao dois numeros de pedido DIFERENTES para a mesma venda -
#     por isso as duas colunas aparecem sempre juntas, nunca uma no lugar da outra.
function Get-Or-NovoGrupo($gruposDict, $l, $nomeSup) {
    $key = "$($l.NumPedWinthor)|$($l.CodCliente)"
    if (-not $gruposDict.ContainsKey($key)) {
        $gruposDict[$key] = [ordered]@{
            data = $l.Data.ToString("yyyy-MM-dd")
            codCliente = $l.CodCliente
            cnpj = $l.CNPJ
            razaoSocial = $l.NomeCliente
            numeroPedidoWinthor = $l.NumPedWinthor
            numeroPedidoClube = $l.NumPedRCA
            codigoRCA = $l.CodVendedor
            vendedor = $l.NomeVendedor
            supervisor = $nomeSup
            segmento = $l.SegTeam
            valorTotal = 0.0
            produtos = New-Object System.Collections.Generic.List[object]
        }
    }
    return $gruposDict[$key]
}

Write-Host "Montando devolucoes..."
$devLinhas = $linhas | Where-Object { $_.TipoVenda -eq "DEVOLUCAO" -or $_.TipoVenda -eq "DEVOLUÇÃO" }
$devPorSupGrupos = @{}
$devGeralGrupos = @{}
foreach ($l in $devLinhas) {
    $cs = $l.CodSupervisor
    $nomeSup = if ($supervisoresInfo.Contains($cs)) { $supervisoresInfo[$cs].nome } else { "Outros/Interno" }
    $valorAbs = [Math]::Round([Math]::Abs($l.Valor),2)
    if (-not $devPorSupGrupos.ContainsKey($cs)) { $devPorSupGrupos[$cs] = @{} }

    $gSup = Get-Or-NovoGrupo $devPorSupGrupos[$cs] $l $nomeSup
    $gGeral = Get-Or-NovoGrupo $devGeralGrupos $l $nomeSup
    foreach ($g in @($gSup, $gGeral)) {
        $g.valorTotal = [Math]::Round($g.valorTotal + $valorAbs, 2)
        $g.produtos.Add([ordered]@{ codProduto = $l.CodProd; produto = $l.DescProduto; fornecedor = $l.Fornecedor; valor = $valorAbs })
    }
}

# ---------------------------------------------------------------------------
# 8. Cortes - relatorio real "1454 - Consultar Corte de Mercadorias" do Winthor
# (substituiu o catalogo estatico em 2026-07-22). E um relatorio IMPRESSO, nao uma
# tabela: repete um cabecalho de filtros a cada "pagina" e reimprime o contexto do
# supervisor/RCA correntes, entao o mesmo supervisor pode aparecer em varios trechos
# nao contiguos do arquivo. O parser abaixo rastreia supervisor/RCA "correntes" linha
# a linha (atualizando sempre que encontra uma linha de cabecalho) e acumula qualquer
# linha de dado sob o contexto vigente no momento - testado e validado: a soma de
# TODAS as linhas de dado bate exatamente com o ultimo "Total do Supervisor:" impresso
# no arquivo (que na verdade e cumulativo do relatorio inteiro, nao por supervisor,
# apesar do nome - servico so como checksum de que nenhuma linha foi perdida).
# Colunas (fixas, confirmadas por inspecao - nao ha cabecalho por coluna no topo do
# arquivo, so nos blocos repetidos "Data|Pedido|Cod.|Descricao|...|Embalagem|...|Un.|
# Qt. Corte|Preco Unit.|_|Vlr. Total|Comprador|Depto."):
#   1=Data  2=Pedido(Winthor)  3=CodProd  4=Descricao  9=Embalagem  11=Unidade
#   12=Qt.Corte  13=Preco Unit.  15=Vlr.Total  16=Comprador  17=Depto.
# Nao tem CNPJ/cod.cliente/razao social - enriquecido via $pedidoWinthorToCliente.
Write-Host "Lendo relatorio real de cortes (Cortes Geral.xls)..."
$excel2 = New-Object -ComObject Excel.Application
$excel2.Visible = $false
$excel2.DisplayAlerts = $false
$wbCorteFile = $excel2.Workbooks.Open((Resolve-BaseFile "Cortes Geral.xls"), $null, $true)
$wsCorte = $wbCorteFile.Worksheets.Item(1)
$rowsCorte = $wsCorte.UsedRange.Rows.Count
$arrCorte = $wsCorte.UsedRange.Value2

$cortePorSupGrupos = @{}
$corteGeralGrupos = @{}
$corteConsumoProduto = @{}
$curCorteSupCod = "OUTROS"; $curCorteSupNome = "Outros/Interno"
$curCorteRcaCod = $null; $curCorteRcaNome = $null
$corteLinhasCount = 0

for ($r = 1; $r -le $rowsCorte; $r++) {
    $c1 = $arrCorte[$r,1]; $c2 = $arrCorte[$r,2]; $c3 = $arrCorte[$r,3]; $c4 = $arrCorte[$r,4]; $c5 = $arrCorte[$r,5]

    if ("$c2" -eq "SUPERVISOR :" -or "$c1" -eq "SUPERVISOR :") {
        $codRaw = "$([int]$c3)"
        $curCorteSupCod = if ($foldSupervisor.ContainsKey($codRaw)) { $foldSupervisor[$codRaw] } else { $codRaw }
        if (-not $supervisoresInfo.Contains($curCorteSupCod)) { $curCorteSupCod = "OUTROS" }
        $curCorteSupNome = if ($supervisoresInfo.Contains($curCorteSupCod)) { $supervisoresInfo[$curCorteSupCod].nome } else { "Outros/Interno" }
        continue
    }
    if ("$c2" -eq "RCA :") {
        $curCorteRcaCod = "$([int]$c3)"
        $curCorteRcaNome = "$c5"
        continue
    }
    # linha de dado: Data (double/serial), Pedido (double) e CodProd (double) juntos
    if ($c1 -is [double] -and $c2 -is [double] -and $c3 -is [double]) {
        $corteLinhasCount++
        $numPedWinthorCorte = "$([int64]$c2)"
        $codProdCorte = "$([int]$c3)"
        $qtCorte = Parse-ValorBR $arrCorte[$r,12]
        $valorCorte = Parse-ValorBR $arrCorte[$r,15]
        $cliente = $pedidoWinthorToCliente[$numPedWinthorCorte]

        $linhaCorte = [pscustomobject]@{
            Data = [datetime]::FromOADate($c1)
            CodCliente = if ($cliente) { $cliente.CodCliente } else { "" }
            NomeCliente = if ($cliente) { $cliente.NomeCliente } else { "Cliente não identificado" }
            CNPJ = if ($cliente) { $cliente.CNPJ } else { "" }
            SegTeam = $null
            NumPedRCA = ""
            NumPedWinthor = $numPedWinthorCorte
            CodVendedor = $curCorteRcaCod
            NomeVendedor = $curCorteRcaNome
            CodSupervisor = $curCorteSupCod
        }
        $forn = $null
        if ($codprodToForn.ContainsKey($codProdCorte)) { $forn = $codprodToForn[$codProdCorte] }

        if (-not $cortePorSupGrupos.ContainsKey($curCorteSupCod)) { $cortePorSupGrupos[$curCorteSupCod] = @{} }
        $gSup = Get-Or-NovoGrupo $cortePorSupGrupos[$curCorteSupCod] $linhaCorte $curCorteSupNome
        $gGeral = Get-Or-NovoGrupo $corteGeralGrupos $linhaCorte $curCorteSupNome
        foreach ($g in @($gSup, $gGeral)) {
            $g.valorTotal = [Math]::Round($g.valorTotal + $valorCorte, 2)
            $g.produtos.Add([ordered]@{
                codProduto = $codProdCorte
                produto = "$($arrCorte[$r,4])".Trim()
                fornecedor = $forn
                embalagem = "$($arrCorte[$r,9])".Trim()
                qtCorte = $qtCorte
                precoUnit = Parse-ValorBR $arrCorte[$r,13]
                valor = [Math]::Round($valorCorte,2)
            })
        }
        if (-not $corteConsumoProduto.ContainsKey($codProdCorte)) {
            $corteConsumoProduto[$codProdCorte] = [ordered]@{
                produto = "$($arrCorte[$r,4])".Trim()
                fornecedor = $forn
                qtCorte = 0.0
                valor = 0.0
            }
        }
        $corteConsumoProduto[$codProdCorte].qtCorte += $qtCorte
        $corteConsumoProduto[$codProdCorte].valor = [Math]::Round($corteConsumoProduto[$codProdCorte].valor + $valorCorte, 2)
    }
}
$wbCorteFile.Close($false)
$excel2.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel2) | Out-Null
Write-Host "  $corteLinhasCount linhas de corte lidas"

# ---------------------------------------------------------------------------
# 8b. Consulta de Pedidos (status operacional - novo na 8022, para vendedores
# consultarem rapidamente o andamento de qualquer pedido da empresa toda).
# Mesmo agrupamento cliente+pedido, mas cobre TODAS as vendas (nao so fornecedor
# API ou itens de corte) e carrega STATUS BLOQUEIO por produto - a informacao
# principal que o vendedor quer ver (PENDENTE/LIBERADO/MONTADO/FATURADO/
# BLOQUEADO/DEVOLVIDO), alem do STATUS PEDIDO (Faturado/A Faturar) ja usado
# no resto do site.
# ---------------------------------------------------------------------------
Write-Host "Montando consulta de pedidos..."
# Consulta cobre o mes vigente + qualquer base historica extra (secao 5b) - diferente do
# dashboard/devolucoes/cortes, que ficam restritos so ao mes vigente.
$consultaLinhas = @($linhas | Where-Object { $_.TipoVenda -eq "VENDA" }) + @($linhasHistorico | Where-Object { $_.TipoVenda -eq "VENDA" })
$consultaGrupos = @{}
foreach ($l in $consultaLinhas) {
    $cs = $l.CodSupervisor
    $nomeSupQ = if ($supervisoresInfo.Contains($cs)) { $supervisoresInfo[$cs].nome } else { "Outros/Interno" }
    $g = Get-Or-NovoGrupo $consultaGrupos $l $nomeSupQ
    $g.valorTotal = [Math]::Round($g.valorTotal + $l.Valor, 2)
    $g.produtos.Add([ordered]@{
        codProduto = $l.CodProd
        produto = $l.DescProduto
        fornecedor = $l.Fornecedor
        statusPedido = $l.StatusFat
        statusBloqueio = $l.StatusBloqueio
        valor = [Math]::Round($l.Valor,2)
    })
}
Write-Host "  $($consultaGrupos.Count) pedidos distintos (mes vigente + historico)"

# Anexa produtos cortados aos pedidos da Consulta - usa a MESMA chave "pedido|cliente"
# de Get-Or-NovoGrupo, entao um pedido que teve corte aparece automaticamente com os
# produtos cortados junto do resto (o corte e' sempre dentro de um pedido de VENDA real,
# diferente da devolucao abaixo, que a base nao vincula a nenhum numero de pedido).
$pedidosComCorte = 0
foreach ($key in $corteGeralGrupos.Keys) {
    if ($consultaGrupos.ContainsKey($key)) {
        $consultaGrupos[$key].produtosCortados = $corteGeralGrupos[$key].produtos
        $pedidosComCorte++
    }
}
Write-Host "  $pedidosComCorte pedidos com produtos cortados"

# Devolucao por cliente, para a Consulta de Pedidos (busca por CNPJ/RCA/supervisor).
# A base 8022 NAO registra o numero do pedido original da devolucao (NUMERO PED. WINTHOR/
# RCA vem com valores fixos "1"/"2" em toda linha de devolucao, nao um numero real - achado
# ao investigar isso) - por isso nao da pra anexar a devolucao a um pedido especifico como
# se faz com o corte acima. Em vez disso, agrupa por CNPJ (cobrindo mes vigente + historico,
# igual a consulta de pedidos) e mostra como um bloco proprio "Devolucao no periodo".
Write-Host "Montando devolucoes para Consulta de Pedidos..."
$devLinhasConsulta = @($linhas | Where-Object { $_.TipoVenda -eq "DEVOLUCAO" -or $_.TipoVenda -eq "DEVOLUÇÃO" }) + @($linhasHistorico | Where-Object { $_.TipoVenda -eq "DEVOLUCAO" -or $_.TipoVenda -eq "DEVOLUÇÃO" })
$devolucaoPorClienteDict = @{}
foreach ($l in $devLinhasConsulta) {
    $key = $l.CNPJ
    if (-not $devolucaoPorClienteDict.ContainsKey($key)) {
        $nomeSupDev = if ($supervisoresInfo.Contains($l.CodSupervisor)) { $supervisoresInfo[$l.CodSupervisor].nome } else { "Outros/Interno" }
        $devolucaoPorClienteDict[$key] = [ordered]@{
            cnpj = $l.CNPJ
            codCliente = $l.CodCliente
            razaoSocial = $l.NomeCliente
            codigoRCA = $l.CodVendedor
            vendedor = $l.NomeVendedor
            supervisor = $nomeSupDev
            valorTotal = 0.0
            produtos = New-Object System.Collections.Generic.List[object]
        }
    }
    $gDev = $devolucaoPorClienteDict[$key]
    $valorAbsDev = [Math]::Round([Math]::Abs($l.Valor),2)
    $gDev.valorTotal = [Math]::Round($gDev.valorTotal + $valorAbsDev, 2)
    $gDev.produtos.Add([ordered]@{
        data = $l.Data.ToString("yyyy-MM-dd")
        codProduto = $l.CodProd
        produto = $l.DescProduto
        fornecedor = $l.Fornecedor
        valor = $valorAbsDev
    })
}
$consultaDevolucoes = New-Object System.Collections.Generic.List[object]
foreach ($gDev in ($devolucaoPorClienteDict.Values | Sort-Object { $_.valorTotal } -Descending)) {
    $gDev.produtos = @($gDev.produtos | Sort-Object { $_.data } -Descending)
    $consultaDevolucoes.Add($gDev)
}
Write-Host "  $($consultaDevolucoes.Count) clientes com devolucao (mes vigente + historico)"

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
        $pedidosArr = @($vAgg.pedidos.Values | Sort-Object { $_.data } -Descending)
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
        vendedores = $vendOut
    })
}

$fornecedoresOut = New-Object System.Collections.Generic.List[object]
foreach ($fid in $fornecedores.Keys) {
    $fornecedoresOut.Add([ordered]@{ id = $fid; nome = $fornecedores[$fid].nome; cor = $fornecedores[$fid].cor })
}

function Sum-GrupoValor($grupos) {
    $soma = 0.0
    foreach ($g in $grupos) { $soma += $g.valorTotal }
    return $soma
}
function Grupos-ParaSaida($gruposDict) {
    return @($gruposDict.Values | ForEach-Object {
        $_.produtos = @($_.produtos | Sort-Object { $_.produto })
        if ($_.Contains("produtosCortados")) { $_.produtosCortados = @($_.produtosCortados | Sort-Object { $_.produto }) }
        $_
    } | Sort-Object { $_.data } -Descending)
}

$devTotalAbs = [Math]::Abs([Math]::Round(($devLinhas | Measure-Object -Property Valor -Sum).Sum,2))
$devolucoesOut = [ordered]@{
    valorTotal = $devTotalAbs
    geral = [ordered]@{
        valorTotal = $devTotalAbs
        grupos = Grupos-ParaSaida $devGeralGrupos
    }
    porSupervisor = New-Object System.Collections.Generic.List[object]
}
foreach ($cs in $devPorSupGrupos.Keys) {
    $nomeSup = if ($supervisoresInfo.Contains($cs)) { $supervisoresInfo[$cs].nome } else { "Outros/Interno" }
    $devolucoesOut.porSupervisor.Add([ordered]@{
        codigo = $cs
        nome = $nomeSup
        valorTotal = [Math]::Round((Sum-GrupoValor $devPorSupGrupos[$cs].Values),2)
        grupos = Grupos-ParaSaida $devPorSupGrupos[$cs]
    })
}

$corteTotal = [Math]::Round((Sum-GrupoValor $corteGeralGrupos.Values),2)
$cortesOut = [ordered]@{
    valorTotal = $corteTotal
    produtos = @($corteConsumoProduto.Values | Sort-Object { $_.valor } -Descending)
    geral = [ordered]@{
        valorTotal = $corteTotal
        grupos = Grupos-ParaSaida $corteGeralGrupos
    }
    porSupervisor = New-Object System.Collections.Generic.List[object]
}
foreach ($cs in $cortePorSupGrupos.Keys) {
    $nomeSup = if ($supervisoresInfo.Contains($cs)) { $supervisoresInfo[$cs].nome } else { "Outros/Interno" }
    $cortesOut.porSupervisor.Add([ordered]@{
        codigo = $cs
        nome = $nomeSup
        valorTotal = [Math]::Round((Sum-GrupoValor $cortePorSupGrupos[$cs].Values),2)
        grupos = Grupos-ParaSaida $cortePorSupGrupos[$cs]
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
    supervisores = $supervisoresOut
    devolucoes = $devolucoesOut
    cortes = $cortesOut
    consultaPedidos = Grupos-ParaSaida $consultaGrupos
    consultaDevolucoes = $consultaDevolucoes
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

# Cache-busting: o GitHub Pages serve estes arquivos com Cache-Control max-age=600, entao o
# navegador pode mostrar codigo/dados de ate 10 minutos atras mesmo depois de publicar uma
# atualizacao. Atualiza a query string "?v=" de TODOS os arquivos locais versionados
# (data.js, script.js, style.css) em cada index.html - nao so data.js, senao uma mudanca no
# script/CSS fica presa em cache do navegador enquanto os dados atualizam normalmente.
$buildVersion = Get-Date -Format "yyyyMMddHHmmss"
foreach ($idxFile in @((Join-Path $sitePath "index.html"), (Join-Path $ritmoPath "index.html"), (Join-Path $consultaPath "index.html"), (Join-Path $root "index.html"))) {
    $content = Get-Content -Path $idxFile -Raw -Encoding UTF8
    $content = $content -replace '((?:data|script|style)\.(?:js|css))\?v=\d+', "`$1?v=$buildVersion"
    Set-Content -Path $idxFile -Value $content -Encoding UTF8 -NoNewline
}
Write-Host "OK: cache-buster atualizado para v=$buildVersion (data.js, script.js, style.css)"

Write-Host "Concluido."
