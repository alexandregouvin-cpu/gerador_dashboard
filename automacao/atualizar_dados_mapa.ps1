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

try {
  $agora = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

  if (-not (Test-Path $PastaRelatorios)) { throw "Pasta de relatórios '$PastaRelatorios' não existe." }
  $rel = Get-ChildItem -Path $PastaRelatorios -Filter $PadraoRelatorio -File |
    Where-Object { $_.Name -notlike '~$*' } |
    Sort-Object LastWriteTimeUtc -Descending |
    Select-Object -First 1
  if (-not $rel) { throw "Nenhum relatório '$PadraoRelatorio' encontrado em '$PastaRelatorios'." }

  $assinatura = "xlsx:{0}|{1}|{2}" -f $rel.Name, $rel.LastWriteTimeUtc.Ticks, $rel.Length
  if (-not $Forcar -and (Test-Path $arquivoAssinatura) -and ([System.IO.File]::ReadAllText($arquivoAssinatura, $utf8) -eq $assinatura)) {
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

  $b64 = [System.Convert]::ToBase64String((Ler-Bytes $rel.FullName))
  $relatorioEm = $rel.LastWriteTimeUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
  $js = "window.MAPA_DADOS = {""versao"":1,""formato"":""xlsx-base64"",""geradoEm"":$(Texto-Json $agora),""relatorioEm"":$(Texto-Json $relatorioEm),""arquivo"":$(Texto-Json $rel.Name),""conteudo"":""$b64""};"

  # grava num temporário e troca de uma vez, para a TV nunca ler um arquivo pela metade
  $tmp = "$arquivoDados.tmp"
  [System.IO.File]::WriteAllText($tmp, $js, $utf8)
  Move-Item -Path $tmp -Destination $arquivoDados -Force
  # versão por último: a TV só baixa os dados quando ela muda
  [System.IO.File]::WriteAllText($arquivoVersao, "window.MAPA_VERSAO = $(Texto-Json $agora);", $utf8)
  [System.IO.File]::WriteAllText($arquivoAssinatura, ("xlsx:{0}|{1}|{2}" -f $rel.Name, $rel.LastWriteTimeUtc.Ticks, $rel.Length), $utf8)
  Registrar "Dados gravados: $($rel.Name)."
  exit 0
}
catch {
  Registrar "ERRO: $($_.Exception.Message)"
  exit 1
}
