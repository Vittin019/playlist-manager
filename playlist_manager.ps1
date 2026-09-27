$ErrorActionPreference = "Continue"

$Desktop = [Environment]::GetFolderPath("Desktop")
$Format  = "m4a"
$Threads = 4

$playlists = @(
    @{ Nome = "Trap"; Folder = Join-Path $Desktop "Trap"; Csv = "Trap.csv"; OriginalM3u = Join-Path $Desktop "Trap.m3u" },
    @{ Nome = "Zona Febril"; Folder = Join-Path $Desktop "Zona Febril"; Csv = "Zona_febril_emojis.csv"; OriginalM3u = Join-Path $Desktop "Zona febril emojis.m3u" },
    @{ Nome = "FBC Yago Oproprio"; Folder = Join-Path $Desktop "Fbc_YagoOproprio"; Csv = "FbcYagoOproprio.csv"; OriginalM3u = Join-Path $Desktop "Fbc_YagoOproprio.m3u" }
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
    if ([string]::IsNullOrWhiteSpace($texto)) { return "" }
    $n = Remove-Diacritics $texto
    return $n.ToLowerInvariant().Trim()
}

function Get-Artista($baseName) {
    if ($baseName -match '^(.*?)\s*-\s*(.+)$') { return $matches[1].Trim() }
    return $baseName
}

function Get-SafeName($nome) {
    if ([string]::IsNullOrWhiteSpace($nome)) { return "" }
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

function Get-TrackUrlFromValue($value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return $null }
    $text = $value.Trim()
    if ($text -match 'spotify:track:([A-Za-z0-9]+)') { return "https://open.spotify.com/track/$($matches[1])" }
    if ($text -match 'https?://') { return $text }
    return $null
}

function Get-AudioFileCount($folder) {
    if (-not (Test-Path -LiteralPath $folder)) { return 0 }
    return @((Get-ChildItem -LiteralPath $folder -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -ieq ".m4a" })).Count
}

function Ensure-Folder($folder) {
    if (-not [string]::IsNullOrWhiteSpace($folder) -and -not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
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

function Read-PlaylistCsv($csvPath) {
    if (-not (Test-Path -LiteralPath $csvPath)) { return @() }

    $items = @()
    try {
        $rows = Import-Csv -LiteralPath $csvPath -Encoding UTF8
    }
    catch {
        try {
            $rows = Import-Csv -LiteralPath $csvPath -Encoding Default
        }
        catch {
            Write-Warning "Nao foi possivel ler CSV: $csvPath"
            return @()
        }
    }

    foreach ($row in $rows) {
        if ($null -eq $row) { continue }

        $title = $null
        foreach ($candidate in @('Title', 'title', 'Name', 'name', 'Track', 'track', 'Song', 'song', 'Musica', 'Música', 'Music', 'music')) {
            if ($row.PSObject.Properties.Name -contains $candidate -and -not [string]::IsNullOrWhiteSpace($row.$candidate)) {
                $title = $row.$candidate
                break
            }
        }

        $uri = $null
        foreach ($candidate in @('SpotifyUrl', 'spotify_url', 'spotifyUrl', 'URL', 'Url', 'url', 'TrackUrl', 'track_url', 'uri', 'URI', 'spotify_uri', 'Spotify URI')) {
            if ($row.PSObject.Properties.Name -contains $candidate -and -not [string]::IsNullOrWhiteSpace($row.$candidate)) {
                $uri = $row.$candidate
                break
            }
        }

        if (-not $title -and $uri) { $title = $uri }
        if (-not $title -and -not $uri) { continue }

        $items += [pscustomobject]@{
            Title = $title
            Url   = Get-TrackUrlFromValue $uri
            Raw   = $row
        }
    }

    return $items
}

function Get-FileHashKey($filePath) {
    if (-not (Test-Path -LiteralPath $filePath)) { return $null }
    try {
        return (Get-FileHash -LiteralPath $filePath -Algorithm SHA256).Hash
    }
    catch {
        return $null
    }
}

function Get-PlaylistMissingEntries($playlist) {
    $csvPath = Join-Path $playlist.Folder $playlist.Csv
    if (-not (Test-Path -LiteralPath $csvPath)) { return @() }

    $entries = Read-PlaylistCsv $csvPath
    $files = @(Get-ChildItem -LiteralPath $playlist.Folder -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -ieq ".m4a" })
    $index = @{}
    foreach ($file in $files) {
        $key = Get-NormalizedKey $file.BaseName
        if (-not [string]::IsNullOrWhiteSpace($key)) { $index[$key] = $file }
    }

    $faltantes = @()
    foreach ($entry in $entries) {
        $key = Get-NormalizedKey $entry.Title
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        if (-not $index.ContainsKey($key)) {
            $faltantes += $entry
        }
    }

    return $faltantes
}

function Invoke-SpotifyDownload($url, $folder) {
    if ([string]::IsNullOrWhiteSpace($url)) { return $false }
    Ensure-Folder $folder

    $args = @(
        '--format', $Format,
        '--output', $folder,
        '--threads', [string]$Threads,
        '--overwrite', 'skip',
        $url
    )

    $output = & spotdl @args 2>&1
    $exitCode = $LASTEXITCODE
    if ($output) {
        $output | ForEach-Object { Write-Info $_ }
    }

    return ($exitCode -eq 0)
}

function Get-M3uEntries($m3uPath) {
    if (-not (Test-Path -LiteralPath $m3uPath)) { return @() }
    $lines = Get-Content -LiteralPath $m3uPath -ErrorAction SilentlyContinue
    $entries = @()
    foreach ($line in $lines) {
        $trimmed = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
        if ($trimmed.StartsWith('#')) { continue }
        $entries += $trimmed
    }
    return $entries
}

function Set-M3uEntries($m3uPath, $entries) {
    $content = @()
    foreach ($entry in $entries) {
        if ([string]::IsNullOrWhiteSpace($entry)) { continue }
        $content += $entry
    }
    Set-Content -LiteralPath $m3uPath -Value $content -Encoding UTF8
}

function Invoke-Modulo1-GerenciarPlaylists {
    Clear-Host
    Write-Header "MODULO 1 - SPOTDL MANAGER"

    if (-not (Test-Spotdl)) {
        Write-ErrorMsg "spotdl nao encontrado. Instale com: pip install spotdl"
        return
    }
    Write-Success "spotdl encontrado"

    foreach ($p in $playlists) {
        Write-Header "PLAYLIST: $($p.Nome)"
        $folder = $p.Folder
        $csvPath = Join-Path $folder $p.Csv

        if (-not (Test-Path -LiteralPath $folder)) {
            Write-ErrorMsg "Pasta nao existe: $folder"
            continue
        }

        if (-not (Test-Path -LiteralPath $csvPath)) {
            Write-ErrorMsg "CSV nao existe: $csvPath"
            continue
        }

        $entries = Read-PlaylistCsv $csvPath
        $faltantes = Get-PlaylistMissingEntries $p

        Write-Success "CSV carregado: $($entries.Count) entradas"
        Write-Success "Pasta validada: $folder"

        if ($faltantes.Count -gt 0) {
            Write-Warning "Existem $($faltantes.Count) itens faltando na pasta"
            foreach ($item in $faltantes) {
                $title = if ($item.Title) { $item.Title } else { "sem nome" }
                Write-Info "- faltando: $title"
            }
        }
        else {
            Write-Success "Nenhuma musica faltando"
        }

        if ($p.OriginalM3u -and (Test-Path -LiteralPath $p.OriginalM3u)) {
            Write-Success "M3U original encontrado: $($p.OriginalM3u)"
        }
    }
}

function Invoke-Modulo2-RemoverDuplicatas {
    Clear-Host
    Write-Header "MODULO 2 - REMOVER DUPLICATAS"

    $pastas = @()
    foreach ($p in $playlists) {
        if ($p.Folder) { $pastas += $p.Folder }
    }

    foreach ($pasta in $pastas) {
        if (-not (Test-Path -LiteralPath $pasta)) { continue }

        $arquivos = @(Get-ChildItem -LiteralPath $pasta -File -ErrorAction SilentlyContinue)
        if ($arquivos.Count -le 1) { continue }

        $processados = @{}
        $remover = @()

        foreach ($arquivo in $arquivos) {
            if ($processados.ContainsKey($arquivo.FullName)) { continue }
            $processados[$arquivo.FullName] = $true

            $hash = Get-FileHashKey $arquivo.FullName
            $key = Get-NormalizedKey $arquivo.BaseName

            foreach ($outro in $arquivos) {
                if ($outro.FullName -eq $arquivo.FullName) { continue }
                if ($processados.ContainsKey($outro.FullName)) { continue }

                $sameHash = $hash -and (Get-FileHashKey $outro.FullName) -and ($hash -eq (Get-FileHashKey $outro.FullName))
                $sameKey = ($key -and (Get-NormalizedKey $outro.BaseName) -and ($key -eq (Get-NormalizedKey $outro.BaseName)))

                if ($sameHash -or $sameKey) {
                    $candidato = $outro
                    if ($arquivo.Length -ge $candidato.Length) {
                        $remover += $candidato.FullName
                        $processados[$outro.FullName] = $true
                    }
                    else {
                        $remover += $arquivo.FullName
                        $processados[$arquivo.FullName] = $true
                        break
                    }
                }
            }
        }

        if ($remover.Count -eq 0) {
            Write-Success "Nenhuma duplicata encontrada em $pasta"
            continue
        }

        $remover = $remover | Sort-Object -Unique

        foreach ($fileToDelete in $remover) {
            $nome = Split-Path $fileToDelete -Leaf
            if ($ArquivosProtegidos -contains $nome.ToLowerInvariant()) {
                Write-Warning "Arquivo protegido, ignorado: $nome"
                continue
            }

            if ($ApagarPermanentemente) {
                Remove-Item -LiteralPath $fileToDelete -Force -ErrorAction SilentlyContinue
                Write-Del "Removido: $nome"
            }
            else {
                Write-Warning "Duplicata detectada (nao removida por seguranca): $nome"
            }
        }
    }
}

function Invoke-Modulo3-CorrigirM3u {
    Clear-Host
    Write-Header "MODULO 3 - CORRIGIR M3U"

    $m3uList = @($playlists | ForEach-Object { $_.OriginalM3u }) + @($playlistsCombinadas | ForEach-Object { $_.M3u })

    foreach ($m3u in $m3uList) {
        if (-not $m3u) { continue }
        if (-not (Test-Path -LiteralPath $m3u)) {
            Write-Warning "M3U nao encontrado: $m3u"
            continue
        }

        $entries = Get-M3uEntries $m3u
        $sorted = @()

        foreach ($entry in $entries) {
            $item = $entry.Trim()
            if ([string]::IsNullOrWhiteSpace($item)) { continue }
            $sorted += $item
        }

        $sorted = $sorted | Sort-Object {
            $leaf = Split-Path $_ -Leaf
            $base = [System.IO.Path]::GetFileNameWithoutExtension($leaf)
            Get-ChaveOrdenacao $base
        }

        Set-M3uEntries $m3u $sorted
        Write-Success "M3U corrigido e ordenado: $m3u"
    }
}

function Invoke-Modulo4-BaixarFaltantes {
    Clear-Host
    Write-Header "MODULO 4 - BAIXAR FALTANTES"

    if (-not (Test-Spotdl)) {
        Write-ErrorMsg "spotdl nao encontrado"
        return
    }

    foreach ($p in $playlists) {
        if (-not (Test-Path -LiteralPath $p.Folder)) {
            Write-ErrorMsg "Pasta nao existe: $($p.Folder)"
            continue
        }

        $faltantes = Get-PlaylistMissingEntries $p
        if ($faltantes.Count -eq 0) {
            Write-Success "Nada para baixar em: $($p.Nome)"
            continue
        }

        Write-Section "Baixando faltantes de $($p.Nome)"
        foreach ($item in $faltantes) {
            $title = if ($item.Title) { $item.Title } else { "titulo indefinido" }
            $url = $item.Url
            if (-not $url) {
                $url = Convert-UriToUrl $item.Raw.Uri
            }

            if (-not $url) {
                Write-Warning "Sem URL valida para: $title"
                continue
            }

            Write-Info "Baixando: $title"
            $ok = Invoke-SpotifyDownload $url $p.Folder
            if ($ok) {
                Write-Success "Download concluido: $title"
            }
            else {
                Write-ErrorMsg "Erro ao baixar: $title"
            }
        }
    }
}

function Invoke-FluxoCompleto {
    Invoke-Modulo2-RemoverDuplicatas
    Invoke-Modulo3-CorrigirM3u
    Invoke-Modulo4-BaixarFaltantes
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
    Invoke-FluxoCompleto
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
