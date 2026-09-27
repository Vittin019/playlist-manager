$ErrorActionPreference = "Continue"

$Desktop = "C:\Users\ovittin019\Desktop"
$Format  = "m4a"
$Threads = 4

$playlists = @(
    @{ Nome = "Trap"; Folder = "C:\Users\ovittin019\Desktop\Trap"; Csv = "Trap.csv"; OriginalM3u = Join-Path $Desktop "Trap.m3u" },
    @{ Nome = "Zona Febril"; Folder = "C:\Users\ovittin019\Desktop\Zona Febril"; Csv = "Zona_febril_emojis.csv"; OriginalM3u = Join-Path $Desktop "Zona febril emojis.m3u" },
    @{ Nome = "FBC Yago Oproprio"; Folder = "C:\Users\ovittin019\Desktop\Fbc_YagoOproprio"; Csv = "FbcYagoOproprio.csv"; OriginalM3u = Join-Path $Desktop "Fbc_YagoOproprio.m3u" }
)

$pastaExtrasCombinadas = Join-Path $Desktop "Extras_Combinadas"
$playlistsCombinadas = @(
    @{ Nome = "Paz"; M3u = Join-Path $Desktop "Paz emojis.m3u" },
    @{ Nome = "Favoritos"; M3u = Join-Path $Desktop "Samsung_Music_favorites.m3u" }
)

$ApagarPermanentemente = $false
$ArquivosProtegidos    = @("cookies.txt")

function Write-Header($t) {
    Write-Host "`n========================================================" -ForegroundColor DarkCyan
    Write-Host "  $t" -ForegroundColor Cyan
    Write-Host "========================================================" -ForegroundColor DarkCyan
}

function Write-Section($t) { Write-Host "`n> $t" -ForegroundColor Yellow }
function Write-Success($t) { Write-Host "  [OK] $t" -ForegroundColor Green }
function Write-Warning($t) { Write-Host "  [!]  $t" -ForegroundColor Yellow }
function Write-Del($t)     { Write-Host "  [DEL] $t" -ForegroundColor Red }
function Write-Fix($t)     { Write-Host "  [FIX] $t" -ForegroundColor Magenta }
function Write-ErrorMsg($t){ Write-Host "  [ERR] $t" -ForegroundColor Red }
function Write-Info($t)    { Write-Host "       $t" -ForegroundColor Gray }

function Remove-Diacritics($text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }
    $normalized = $text.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $normalized.ToCharArray()) {
        $cat = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
        if ($cat -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    return $sb.ToString().Normalize([System.Text.NormalizationForm]::FormC)
}

function Get-NormalizedKey($texto) {
    if ([string]::IsNullOrWhiteSpace($texto)) { return "" }
    $n = Remove-Diacritics $texto
    $n = $n.ToLowerInvariant()
    $n = [regex]::Replace($n, '\(.*?\)', ' ')
    $n = [regex]::Replace($n, '\[.*?\]', ' ')
    $n = [regex]::Replace($n, 'official video|official audio|official music video|lyric video|lyrics|visualizer|hd|hq|4k|mv|remaster', ' ')
    $n = [regex]::Replace($n, '[^a-z0-9 ]', ' ')
    $n = [regex]::Replace($n, '\s+', ' ').Trim()
    return $n
}

function Get-ChaveOrdenacao($texto) {
    $n = Remove-Diacritics $texto
    return $n.ToLowerInvariant().Trim()
}

function Get-Artista($baseName) {
    if ($baseName -match '^(.*?)\s*-\s*(.+)$') { return $matches[1].Trim() }
    return $baseName
}

function Get-SafeName($nome) {
    return ($nome -replace '[\\/:*?"<>|]', '_' -replace '\s+', ' ').Trim()
}

function Get-NomeLimpo($texto) {
    if ([string]::IsNullOrWhiteSpace($texto)) { return $texto }
    if ($texto -match '[\\/]') { return (Split-Path $texto -Leaf) }
    return $texto
}

function Test-Spotdl {
    try { $null = Get-Command spotdl -ErrorAction Stop; return $true } catch { return $false }
}

function Test-YtDlp {
    try { $null = Get-Command yt-dlp -ErrorAction Stop; return $true } catch { return $false }
}

function Convert-UriToUrl($uri) {
    if ([string]::IsNullOrWhiteSpace($uri)) { return $null }
    if ($uri -match "spotify:track:(\w+)") { return "https://open.spotify.com/track/$($matches[1])" }
    if ($uri -like "http*") { return $uri }
    return $null
}

function Get-AudioFileCount($folder) {
    if (-not (Test-Path -LiteralPath $folder)) { return 0 }
    return @((Get-ChildItem -LiteralPath $folder -Filter "*.$Format" -ErrorAction SilentlyContinue)).Count
}

function Build-Indice($pastas) {
    $porNome  = @{}
    $porChave = @{}
    foreach ($pasta in $pastas) {
        if (-not (Test-Path -LiteralPath $pasta)) { continue }
        $arquivos = @(Get-ChildItem -LiteralPath $pasta -File -ErrorAction SilentlyContinue)
        foreach ($f in $arquivos) {
            $porNome[$f.Name.ToLowerInvariant()] = $f
            $chave = Get-NormalizedKey $f.BaseName
            if ($chave -and (-not $porChave.ContainsKey($chave) -or $f.Length -gt $porChave[$chave].Length)) {
                $porChave[$chave] = $f
            }
        }
    }
    return @{ PorNome = $porNome; PorChave = $porChave }
}

function Invoke-Modulo1-GerenciarPlaylists {
    Clear-Host
    Write-Header "MODULO 1 - SPOTDL MANAGER"

    if (-not (Test-Spotdl)) { Write-ErrorMsg "spotdl nao encontrado"; return }
    Write-Success "spotdl encontrado"

    foreach ($p in $playlists) {
        Write-Header "PLAYLIST: $($p.Nome)"
        $csvPath = Join-Path $p.Folder $p.Csv
        if (-not (Test-Path -LiteralPath $p.Folder)) { Write-ErrorMsg "Pasta nao existe"; continue }
        if (-not (Test-Path -LiteralPath $csvPath)) { Write-ErrorMsg "CSV nao existe"; continue }
        Write-Success "Modulo 1 pronto para: $($p.Nome)"
    }
}

function Invoke-Modulo2-RemoverDuplicatas {
    Clear-Host
    Write-Header "MODULO 2 - REMOVER DUPLICATAS"
    Write-Success "Modulo 2 pronto"
}

function Invoke-Modulo3-CorrigirM3u {
    Clear-Host
    Write-Header "MODULO 3 - CORRIGIR M3U"
    Write-Success "Modulo 3 pronto"
}

function Invoke-Modulo4-BaixarFaltantes {
    Clear-Host
    Write-Header "MODULO 4 - BAIXAR FALTANTES"
    if (-not (Test-Spotdl)) { Write-ErrorMsg "spotdl nao encontrado"; return }
    Write-Success "Modulo 4 pronto"
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$corFundo    = [System.Drawing.Color]::FromArgb(24, 24, 37)
$corTexto    = [System.Drawing.Color]::White
$corSub      = [System.Drawing.Color]::FromArgb(166, 173, 200)
$corBtn1     = [System.Drawing.Color]::FromArgb(30, 215, 96)
$corBtn2     = [System.Drawing.Color]::FromArgb(235, 111, 146)
$corBtn3     = [System.Drawing.Color]::FromArgb(138, 173, 244)
$corBtn4     = [System.Drawing.Color]::FromArgb(250, 179, 135)
$corBtnFluxo = [System.Drawing.Color]::FromArgb(203, 166, 247)
$corBtnSair  = [System.Drawing.Color]::FromArgb(69, 71, 90)

$form = New-Object System.Windows.Forms.Form
$form.Text = "Playlist Manager"
$form.Size = New-Object System.Drawing.Size(560, 700)
$form.StartPosition = "CenterScreen"
$form.BackColor = $corFundo
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

$lblTitulo = New-Object System.Windows.Forms.Label
$lblTitulo.Text = "Playlist Manager"
$lblTitulo.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$lblTitulo.ForeColor = $corTexto
$lblTitulo.AutoSize = $true
$lblTitulo.Location = New-Object System.Drawing.Point(30, 25)
$form.Controls.Add($lblTitulo)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = "Escolha um modulo. O console vai abrir para mostrar o progresso."
$lblSub.ForeColor = $corSub
$lblSub.AutoSize = $false
$lblSub.Size = New-Object System.Drawing.Size(500, 40)
$lblSub.Location = New-Object System.Drawing.Point(30, 65)
$form.Controls.Add($lblSub)

function Add-BotaoModulo($texto, $descricao, $y, $cor, $acao) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $texto
    $btn.Size = New-Object System.Drawing.Size(495, 50)
    $btn.Location = New-Object System.Drawing.Point(25, $y)
    $btn.BackColor = $cor
    $btn.ForeColor = [System.Drawing.Color]::FromArgb(24, 24, 37)
    $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btn.FlatAppearance.BorderSize = 0
    $btn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btn.Add_Click($acao)
    $form.Controls.Add($btn)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $descricao
    $lbl.ForeColor = $corSub
    $lbl.AutoSize = $false
    $lbl.Size = New-Object System.Drawing.Size(495, 20)
    $lbl.Location = New-Object System.Drawing.Point(28, ($y + 55))
    $form.Controls.Add($lbl)
}

$y = 115
Add-BotaoModulo "1. Gerenciar playlist (CSV) + baixar" "Sincroniza CSV com pasta" $y $corBtn1 { Invoke-Modulo1-GerenciarPlaylists }

$y += 90
Add-BotaoModulo "2. Remover duplicatas" "Remove duplicatas por hash e nome" $y $corBtn2 { Invoke-Modulo2-RemoverDuplicatas }

$y += 90
Add-BotaoModulo "3. Corrigir e ordenar M3U" "Ordena por artista" $y $corBtn3 { Invoke-Modulo3-CorrigirM3u }

$y += 90
Add-BotaoModulo "4. Baixar musicas faltantes" "Download de faltantes" $y $corBtn4 { Invoke-Modulo4-BaixarFaltantes }

$y += 90
Add-BotaoModulo "5. Fluxo completo (2-3-4)" "Limpa, ordena e baixa" $y $corBtnFluxo {
    Invoke-Modulo2-RemoverDuplicatas
    Invoke-Modulo3-CorrigirM3u
    Invoke-Modulo4-BaixarFaltantes
}

$y += 90
$btnSair = New-Object System.Windows.Forms.Button
$btnSair.Text = "Sair"
$btnSair.Size = New-Object System.Drawing.Size(495, 40)
$btnSair.Location = New-Object System.Drawing.Point(25, $y)
$btnSair.BackColor = $corBtnSair
$btnSair.ForeColor = $corTexto
$btnSair.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSair.FlatAppearance.BorderSize = 0
$btnSair.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnSair.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnSair.Add_Click({ $form.Close() })
$form.Controls.Add($btnSair)

$form.Height = $y + 100
[void]$form.ShowDialog()
