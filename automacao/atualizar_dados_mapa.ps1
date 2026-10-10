<#
.SYNOPSIS
  Gera dados\mapa_dados.js para o Mapa de Transportadoras da TV.

.DESCRIPTION
  O mapa (mapa_tv.html) fica numa pasta do SharePoint/OneDrive sincronizada
  no computador da TV. Ele lê dados\mapa_dados.js ao abrir e confere a cada
  5 minutos se há versão nova. Este script pega o relatório "Lista de
  Parceiros de negócios" mais recente da pasta relatorios\ (do jeito que sai
  do SAP, sem filtro nem PROCX) e grava esse arquivo.

  As regras (Ativo = Y, código do PN começando com C e transportadoras
  autorizadas) são aplicadas pela própria página, como no upload manual.

  Também mantém o histórico de mudanças (dados\historico.js): a cada relatório
  novo, compara os clientes com o relatório anterior e registra quem entrou,
  saiu, trocou de transportadora ou de cidade.

  Só regrava o arquivo de dados quando o relatório mudou. Compatível com o
  Windows PowerShell 5.1 que já vem no Windows.

.EXAMPLE
  powershell -NoProfile -ExecutionPolicy Bypass -File atualizar_dados_mapa.ps1
#>
[CmdletBinding()]
param(
  # Pasta do mapa (onde fica mapa_tv.html). Padrão: a pasta acima de "automacao".
  [string]$PastaPainel = "",
  # Onde o relatório .xlsx é salvo. Padrão: <PastaPainel>\relatorios
  [string]$PastaRelatorios = "",
  # Nome do relatório (curingas permitidos)
  [string]$PadraoRelatorio = "*Parceiros*.xlsx",
  # Regrava mesmo sem mudança
  [switch]$Forcar
)

$ErrorActionPreference = "Stop"
$utf8 = New-Object System.Text.UTF8Encoding($false)

# Pasta deste script. Em alguns Windows PowerShell 5.1 o $PSScriptRoot vem vazio: usa o caminho do
# próprio script e, em último caso, a pasta atual (de onde o comando foi rodado).
$pastaScript = $PSScriptRoot
if (-not $pastaScript -and $MyInvocation.MyCommand.Path) { $pastaScript = Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $pastaScript) { $pastaScript = (Get-Location).ProviderPath }
if (-not $PastaPainel) { $PastaPainel = Split-Path -Parent $pastaScript }
$PastaPainel = (Resolve-Path -LiteralPath $PastaPainel).ProviderPath

if (-not $PastaRelatorios) { $PastaRelatorios = Join-Path $PastaPainel "relatorios" }
$pastaDados = Join-Path $PastaPainel "dados"
$arquivoDados = Join-Path $pastaDados "mapa_dados.js"
$arquivoVersao = Join-Path $pastaDados "versao.js"
$arquivoAssinatura = Join-Path $pastaDados ".assinatura"
$arquivoLog = Join-Path $pastaDados "atualizacao.log"
if (-not (Test-Path $pastaDados)) { New-Item -ItemType Directory -Path $pastaDados | Out-Null }

function Registrar([string]$msg) {
  $linha = "{0:yyyy-MM-dd HH:mm:ss}  {1}" -f (Get-Date), $msg
  Write-Host $linha
  $anteriores = @()
  if (Test-Path $arquivoLog) { $anteriores = @(Get-Content -Path $arquivoLog -Encoding UTF8 | Select-Object -Last 499) }
  [System.IO.File]::WriteAllLines($arquivoLog, [string[]]($anteriores + $linha), $utf8)
}

# Lê o arquivo mesmo que esteja aberto no Excel ou sincronizando
function Ler-Bytes([string]$caminho) {
  $fs = [System.IO.File]::Open($caminho, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
  try {
    $ms = New-Object System.IO.MemoryStream
    $fs.CopyTo($ms)
    return $ms.ToArray()
  } finally { $fs.Dispose() }
}

function Texto-Json([string]$s) {
  if ($null -eq $s) { return "null" }
  return (ConvertTo-Json -InputObject $s -Compress)
}

# ---------- Histórico de mudanças ----------
# Lê o .xlsx direto (sem Excel), guarda uma foto dos clientes "C" em dados\clientes_base.tsv e,
# a cada relatório novo, grava em dados\historico.js o que mudou: cliente novo ou removido,
# ativo/inativo, código da transportadora e cidade/UF. O mapa aplica as regras ao exibir.

function Normalizar([string]$s) {
  if (-not $s) { return "" }
  $d = $s.Normalize([System.Text.NormalizationForm]::FormD)
  $sb = New-Object System.Text.StringBuilder
  foreach ($ch in $d.ToCharArray()) {
    if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($ch) }
  }
  return (($sb.ToString().ToUpperInvariant()) -replace '[^A-Z0-9]+', ' ').Trim()
}

function Indice-Coluna([string]$ref) {
  $n = 0
  foreach ($ch in $ref.ToCharArray()) {
    $v = [int][char]::ToUpperInvariant($ch)
    if ($v -ge 65 -and $v -le 90) { $n = $n * 26 + ($v - 64) } else { break }
  }
  return $n - 1
}

# Primeira aba da planilha -> lista de linhas (cada linha = string[] com os valores das colunas)
function Ler-Planilha([byte[]]$bytes, [scriptblock]$Util = { $true }) {
  Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
  $ms = New-Object System.IO.MemoryStream(, $bytes)
  $zip = New-Object System.IO.Compression.ZipArchive($ms, [System.IO.Compression.ZipArchiveMode]::Read)
  try {
    # textos compartilhados
    $textos = New-Object 'System.Collections.Generic.List[string]'
    $ent = $zip.GetEntry("xl/sharedStrings.xml")
    if ($ent) {
      $r = [System.Xml.XmlReader]::Create($ent.Open())
      try {
        $sb = New-Object System.Text.StringBuilder; $emT = $false; $emRph = $false
        while ($r.Read()) {
          $tipo = [int]$r.NodeType
          if ($tipo -eq 1) {                                 # elemento
            $nome = $r.LocalName
            if ($nome -eq "t") { $emT = -not $r.IsEmptyElement }
            elseif ($nome -eq "si") { [void]$sb.Clear(); if ($r.IsEmptyElement) { $textos.Add("") } }
            elseif ($nome -eq "rPh") { $emRph = -not $r.IsEmptyElement }
          }
          elseif ($tipo -eq 15) {                            # fim de elemento
            $nome = $r.LocalName
            if ($nome -eq "t") { $emT = $false }
            elseif ($nome -eq "si") { $textos.Add($sb.ToString()) }
            elseif ($nome -eq "rPh") { $emRph = $false }
          }
          elseif ($emT -and -not $emRph -and ($tipo -eq 3 -or $tipo -eq 13 -or $tipo -eq 14)) { [void]$sb.Append($r.Value) }   # texto
        }
      } finally { $r.Dispose() }
    }

    # caminho da primeira aba
    $caminhoAba = "xl/worksheets/sheet1.xml"
    try {
      $wb = New-Object System.Xml.XmlDocument; $wb.Load($zip.GetEntry("xl/workbook.xml").Open())
      $aba = $wb.GetElementsByTagName("sheet") | Select-Object -First 1
      $rid = $aba.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
      $rels = New-Object System.Xml.XmlDocument; $rels.Load($zip.GetEntry("xl/_rels/workbook.xml.rels").Open())
      foreach ($rel in $rels.GetElementsByTagName("Relationship")) {
        if ($rel.GetAttribute("Id") -eq $rid) {
          $alvo = $rel.GetAttribute("Target").TrimStart("/")
          if ($alvo -notlike "xl/*") { $alvo = "xl/" + $alvo }
          $caminhoAba = $alvo
        }
      }
    } catch { }

    # depois do cabeçalho, só lê as colunas úteis ($Util); as outras são puladas sem ler o conteúdo
    $linhas = New-Object 'System.Collections.Generic.List[object]'
    $usar = $null
    $r = [System.Xml.XmlReader]::Create($zip.GetEntry($caminhoAba).Open())
    try {
      $linha = $null; $col = -1; $tipoCel = ""; $sb = New-Object System.Text.StringBuilder; $emValor = $false
      $segue = $r.Read()
      while ($segue) {
        $tipo = [int]$r.NodeType
        if ($tipo -eq 1) {                                   # elemento
          $nome = $r.LocalName
          if ($nome -eq "c") {
            $ref = $r.GetAttribute("r")
            if ($ref) {
              $col = 0
              foreach ($ch in $ref.ToCharArray()) { $v = [int]$ch; if ($v -ge 65 -and $v -le 90) { $col = $col * 26 + $v - 64 } else { break } }
              $col--
            } else { $col++ }
            if ($null -ne $usar -and ($col -ge $usar.Length -or -not $usar[$col])) { $r.Skip(); $segue = -not $r.EOF; continue }
            $tipoCel = $r.GetAttribute("t"); [void]$sb.Clear()
          }
          elseif ($nome -eq "v" -or $nome -eq "t") { $emValor = -not $r.IsEmptyElement }
          elseif ($nome -eq "row") { $linha = New-Object 'System.Collections.Generic.List[string]'; $col = -1 }
        }
        elseif ($tipo -eq 15) {                              # fim de elemento
          $nome = $r.LocalName
          if ($nome -eq "v" -or $nome -eq "t") { $emValor = $false }
          elseif ($nome -eq "c") {
            $valor = $sb.ToString()
            if ($tipoCel -eq "s" -and $valor -ne "") { $valor = $textos[[int]$valor] }
            while ($linha.Count -le $col) { $linha.Add("") }
            $linha[$col] = $valor
          }
          elseif ($nome -eq "row") {
            $linhas.Add($linha.ToArray())
            if ($null -eq $usar) {
              $usar = New-Object 'bool[]' ($linha.Count)
              for ($k = 0; $k -lt $linha.Count; $k++) { $usar[$k] = [bool](& $Util (Normalizar $linha[$k])) }
            }
            $linha = $null
          }
        }
        elseif ($emValor -and ($tipo -eq 3 -or $tipo -eq 13 -or $tipo -eq 14)) { [void]$sb.Append($r.Value) }   # texto
        $segue = $r.Read()
      }
    } finally { $r.Dispose() }
    return , $linhas
  } finally { $zip.Dispose(); $ms.Dispose() }
}

function Celula($linha, [int]$i) {
  if ($i -lt 0 -or $i -ge $linha.Length) { return "" }
  return (($linha[$i] -replace '[\t\r\n]+', ' ').Trim())
}

# Clientes "C" do relatório -> dicionário código -> @(nome, cidade, uf, ativo, transportadora)
function Ler-Clientes([byte[]]$bytes) {
  $linhas = Ler-Planilha $bytes { param($c) $c -match 'CODIGO DO PN|NOME DO PN|CIDADE|MUNICIPIO|^UF$|ESTADO|^ATIVO$|TRANSPORT' }
  if ($linhas.Count -lt 2) { throw "planilha vazia" }
  $cab = @($linhas[0] | ForEach-Object { Normalizar $_ })
  $achar = {
    param([scriptblock[]]$testes)
    foreach ($t in $testes) { for ($i = 0; $i -lt $cab.Count; $i++) { if (& $t $cab[$i]) { return $i } } }
    return -1
  }
  $iCod = & $achar @({ param($c) $c -eq "CODIGO DO PN" })
  $iNome = & $achar @({ param($c) $c -eq "NOME DO PN" })
  $iCid = & $achar @({ param($c) $c -eq "ENDERECO DE COBRANCA CIDADE" }, { param($c) $c -like "*CIDADE*" }, { param($c) $c -like "*MUNICIPIO*" })
  $iUf = & $achar @({ param($c) $c -eq "UF" }, { param($c) $c -eq "ESTADO DO DESTINATARIO" }, { param($c) $c -like "*ESTADO*" })
  $iAtivo = & $achar @({ param($c) $c -eq "ATIVO" })
  if ($iCod -lt 0 -or $iAtivo -lt 0) { throw "colunas 'Código do PN' e 'Ativo' não encontradas" }
  # código da transportadora: pelo título ou, se vier como "Transportadora", pelo conteúdo (F00141...)
  $iTransp = & $achar @({ param($c) ($c -replace ' ', '') -eq "CODIGOTRANSPORTADORA" })
  if ($iTransp -lt 0) {
    for ($i = 0; $i -lt $cab.Count -and $iTransp -lt 0; $i++) {
      if ($cab[$i] -notlike "*TRANSPORT*") { continue }
      $tot = 0; $cod = 0
      for ($k = 1; $k -lt $linhas.Count -and $tot -lt 3000; $k++) {
        $v = Celula $linhas[$k] $i
        if ($v) { $tot++; if ($v -match '^[A-Za-z]\d{4,}$') { $cod++ } }
      }
      if ($tot -gt 0 -and $cod -ge $tot * 0.8) { $iTransp = $i }
    }
  }
  if ($iTransp -lt 0) { throw "coluna com o código da transportadora não encontrada" }

  $clientes = New-Object 'System.Collections.Generic.Dictionary[string,string[]]'
  for ($k = 1; $k -lt $linhas.Count; $k++) {
    $l = $linhas[$k]
    $cod = (Celula $l $iCod).ToUpperInvariant()
    if (-not $cod.StartsWith("C")) { continue }
    $clientes[$cod] = @((Celula $l $iNome), (Celula $l $iCid), (Celula $l $iUf).ToUpperInvariant(), (Celula $l $iAtivo).ToUpperInvariant(), (Celula $l $iTransp).ToUpperInvariant())
  }
  return , $clientes
}

function Ler-Base([string]$arquivo) {
  $base = New-Object 'System.Collections.Generic.Dictionary[string,string[]]'
  foreach ($l in [System.IO.File]::ReadAllLines($arquivo, $utf8)) {
    if (-not $l -or $l.StartsWith("#")) { continue }
    $p = $l.Split("`t")
    if ($p.Length -ge 6) { $base[$p[0]] = @($p[1], $p[2], $p[3], $p[4], $p[5]) }
  }
  return , $base
}

function Gravar-Base([string]$arquivo, $clientes, [string]$relatorioEm, [string]$nomeArquivo) {
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine("#relatorio`t$relatorioEm`t$nomeArquivo")
  foreach ($k in $clientes.Keys) { [void]$sb.AppendLine($k + "`t" + ($clientes[$k] -join "`t")) }
  $tmp = "$arquivo.tmp"
  [System.IO.File]::WriteAllText($tmp, $sb.ToString(), $utf8)
  Move-Item -Path $tmp -Destination $arquivo -Force
}

function Json-Texto([string]$s) {
  if ($null -eq $s) { return "null" }
  $s = $s.Replace('\', '\\').Replace('"', '\"') -replace '[\x00-\x1f]', ' '
  return '"' + $s + '"'
}
function Json-Estado($v) {
  if ($null -eq $v) { return "null" }
  return "[" + ((@($v[3], $v[4], $v[1], $v[2]) | ForEach-Object { Json-Texto $_ }) -join ",") + "]"   # ativo, transportadora, cidade, uf
}

# Compara o relatório com a base e acrescenta um evento ao histórico. Devolve um resumo para o log.
function Atualizar-Historico([byte[]]$bytes, $rel, [string]$agora) {
  $arquivoBase = Join-Path $pastaDados "clientes_base.tsv"
  $arquivoEventos = Join-Path $pastaDados "historico.jsl"
  $arquivoHistJs = Join-Path $pastaDados "historico.js"
  $relatorioEm = $rel.LastWriteTimeUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
  $novos = Ler-Clientes $bytes
  if ($novos.Count -eq 0) { throw "nenhum cliente 'C' no relatório" }

  $mudancas = New-Object 'System.Collections.Generic.List[string]'
  $ehBase = -not (Test-Path $arquivoBase)
  if (-not $ehBase) {
    $base = Ler-Base $arquivoBase
    if ($novos.Count -lt $base.Count * 0.5) {
      throw "relatório com $($novos.Count) clientes 'C' e a base anterior tem $($base.Count); parece incompleto, histórico não alterado"
    }
    $chaves = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($k in $base.Keys) { [void]$chaves.Add($k) }
    foreach ($k in $novos.Keys) { [void]$chaves.Add($k) }
    foreach ($k in $chaves) {
      $a = $null; $d = $null
      [void]$base.TryGetValue($k, [ref]$a)
      [void]$novos.TryGetValue($k, [ref]$d)
      $mudou = $false
      if ($null -eq $a -or $null -eq $d) { $mudou = $true }
      elseif ($a[3] -ne $d[3] -or $a[4] -ne $d[4] -or (Normalizar $a[1]) -ne (Normalizar $d[1]) -or $a[2] -ne $d[2]) { $mudou = $true }
      if (-not $mudou) { continue }
      $nome = if ($null -ne $d) { $d[0] } else { $a[0] }
      $mudancas.Add('{"c":' + (Json-Texto $k) + ',"n":' + (Json-Texto $nome) + ',"a":' + (Json-Estado $a) + ',"d":' + (Json-Estado $d) + '}')
    }
  }

  if ($ehBase -or $mudancas.Count -gt 0) {
    $evento = '{"relatorioEm":' + (Json-Texto $relatorioEm) + ',"processadoEm":' + (Json-Texto $agora) + ',"arquivo":' + (Json-Texto $rel.Name) +
      ',"base":' + $(if ($ehBase) { "true" } else { "false" }) + ',"clientes":' + $novos.Count + ',"mudancas":[' + ($mudancas -join ",") + ']}'
    $eventos = @()
    if (Test-Path $arquivoEventos) { $eventos = @([System.IO.File]::ReadAllLines($arquivoEventos, $utf8) | Where-Object { $_ }) }
    $eventos = @($eventos + $evento | Select-Object -Last 400)
    [System.IO.File]::WriteAllLines($arquivoEventos, [string[]]$eventos, $utf8)
    $tmp = "$arquivoHistJs.tmp"
    [System.IO.File]::WriteAllText($tmp, "window.MAPA_HISTORICO = [" + ($eventos -join ",`n") + "];", $utf8)
    Move-Item -Path $tmp -Destination $arquivoHistJs -Force
  }
  Gravar-Base $arquivoBase $novos $relatorioEm $rel.Name
  if ($ehBase) { return "histórico iniciado com $($novos.Count) clientes 'C'" }
  return "histórico: $($mudancas.Count) mudança(s)"
}

try {
  $agora = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

  if (-not (Test-Path $PastaRelatorios)) { throw "Pasta de relatórios '$PastaRelatorios' não existe." }
  $rel = Get-ChildItem -Path $PastaRelatorios -Filter $PadraoRelatorio -File |
    Where-Object { $_.Name -notlike '~$*' } |
    Sort-Object LastWriteTimeUtc -Descending |
    Select-Object -First 1
  if (-not $rel) { throw "Nenhum relatório '$PadraoRelatorio' encontrado em '$PastaRelatorios'." }

  $assinatura = "xlsx:{0}|{1}|{2}" -f $rel.Name, $rel.LastWriteTimeUtc.Ticks, $rel.Length
  $temBase = Test-Path (Join-Path $pastaDados "clientes_base.tsv")
  if (-not $Forcar -and $temBase -and (Test-Path $arquivoAssinatura) -and ([System.IO.File]::ReadAllText($arquivoAssinatura, $utf8) -eq $assinatura)) {
    Registrar "Sem mudanças ($($rel.Name))."
    exit 0
  }

  # espera o arquivo parar de crescer (SAP salvando ou OneDrive sincronizando)
  $tam = -1
  for ($i = 0; $i -lt 10 -and $tam -ne (Get-Item $rel.FullName).Length; $i++) {
    $tam = (Get-Item $rel.FullName).Length
    Start-Sleep -Seconds 3
  }
  $rel = Get-Item $rel.FullName

  $bytes = Ler-Bytes $rel.FullName
  $b64 = [System.Convert]::ToBase64String($bytes)
  $relatorioEm = $rel.LastWriteTimeUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
  $js = "window.MAPA_DADOS = {""versao"":1,""formato"":""xlsx-base64"",""geradoEm"":$(Texto-Json $agora),""relatorioEm"":$(Texto-Json $relatorioEm),""arquivo"":$(Texto-Json $rel.Name),""conteudo"":""$b64""};"

  # grava num temporário e troca de uma vez, para a TV nunca ler um arquivo pela metade
  $tmp = "$arquivoDados.tmp"
  [System.IO.File]::WriteAllText($tmp, $js, $utf8)
  Move-Item -Path $tmp -Destination $arquivoDados -Force
  # histórico de mudanças: se falhar, o mapa é atualizado do mesmo jeito
  $msgHist = ""
  try { $msgHist = " " + (Atualizar-Historico $bytes $rel $agora) + "." }
  catch { $msgHist = " AVISO: histórico não atualizado ($($_.Exception.Message))." }
  # versão por último: a TV só baixa os dados quando ela muda
  [System.IO.File]::WriteAllText($arquivoVersao, "window.MAPA_VERSAO = $(Texto-Json $agora);", $utf8)
  [System.IO.File]::WriteAllText($arquivoAssinatura, ("xlsx:{0}|{1}|{2}" -f $rel.Name, $rel.LastWriteTimeUtc.Ticks, $rel.Length), $utf8)
  Registrar "Dados gravados: $($rel.Name).$msgHist"
  exit 0
}
catch {
  Registrar "ERRO: $($_.Exception.Message)"
  exit 1
}
