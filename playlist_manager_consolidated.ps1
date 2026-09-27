# ============================================================
#  PLAYLIST MANAGER CONSOLIDADO — TODOS OS MÓDULOS EM UM SÓ ARQUIVO
#  Interface simples no console
# ============================================================

$ErrorActionPreference = "Continue"

# ============================================================
#  CONFIGURAÇÃO GERAL
# ============================================================
$Desktop = "C:\Users\ovittin019\Desktop"
$Format  = "m4a"
$Threads = 4

$playlists = @(
    @{ 
        Nome = "Trap"
        Folder = "C:\Users\ovittin019\Desktop\Trap"
        Csv = "Trap.csv"
        OriginalM3u = Join-Path $Desktop "Trap.m3u"
    },
    @{ 
        Nome = "Zona Febril"
        Folder = "C:\Users\ovittin019\Desktop\Zona Febril"
        Csv = "Zona_febril_🤒.csv"
        OriginalM3u = Join-Path $Desktop "Zona febril 🤒.m3u"
    },
    @{ 
        Nome = "FBC / Yago Opróprio"
        Folder = "C:\Users\ovittin019\Desktop\Fbc_YagoOproprio"
        Csv = "FbcYagoOproprio.csv"
        OriginalM3u = Join-Path $Desktop "Fbc_YagoOproprio.m3u"
    }
)

$pastaExtrasCombinadas = Join-Path $Desktop "Extras_Combinadas"

$playlistsCombinadas = @(
    @{ Nome = "Paz"; M3u = Join-Path $Desktop "Paz 💭.m3u" },
    @{ Nome = "Favoritos"; M3u = Join-Path $Desktop "!#SamsungMusic_favorites#!.m3u" }
)

$ApagarPermanentemente = $false
$ArquivosProtegidos = @("cookies.txt")

# ============================================================
#  FUNÇÕES DE UI
# ============================================================
function Write-Header($t) {
    Write-Host "`n========================================" -ForegroundColor DarkCyan
    Write-Host "  $t" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor DarkCyan
}

function Write-Section($t) { Write-Host "`n> $t" -ForegroundColor Yellow }
function Write-Success($t) { Write-Host "  [OK] $t" -ForegroundColor Green }
function Write-Warning($t) { Write-Host "  [!]  $t" -ForegroundColor Yellow }
function Write-Del($t)     { Write-Host "  [DEL] $t" -ForegroundColor Red }
function Write-Fix($t)     { Write-Host "  [FIX] $t" -ForegroundColor Magenta }
function Write-ErrorMsg($t){ Write-Host "  [ERR] $t" -ForegroundColor Red }
function Write-Info($t)    { Write-Host "       $t" -ForegroundColor Gray }

# ============================================================
#  FUNÇÕES DE TEXTO
# ============================================================
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
    $noise = @('official video','official audio','official music video','lyrics','video','audio','feat\.?','ft\.?')
    foreach ($w in $noise) { $n = [regex]::Replace($n, "\b$w\b", ' ') }
    $n = [regex]::Replace($n, '[^a-z0-9\s]', ' ')
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

# ============================================================
#  UTILITÁRIOS
# ============================================================
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

function Get-FileHashKey($filePath) {
    if (-not (Test-Path -LiteralPath $filePath)) { return $null }
    try {
        return (Get-FileHash -LiteralPath $filePath -Algorithm SHA256).Hash
    }
    catch {
        return $null
    }
}

function Get-ArquivosPermitidos($p) {
    $safeName = Get-SafeName $p.Nome
    $permitidos = New-Object System.Collections.Generic.HashSet[string]
    [void]$permitidos.Add("$safeName.m3u")
    if ($p.Csv) { [void]$permitidos.Add($p.Csv) }
    foreach ($prot in $ArquivosProtegidos) { [void]$permitidos.Add($prot) }
    return $permitidos
}

# ============================================================
#  LEITURA E ESCRITA M3U
# ============================================================
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

function Read-M3UFull($path) {
    if (-not (Test-Path -LiteralPath $path)) { return @() }
    $linhas = Get-Content -Path $path -Encoding UTF8 -ErrorAction SilentlyContinue
    $entries = @()
    $infPendente = $null
    foreach ($linhaCrua in $linhas) {
        $l = $linhaCrua.Trim()
        if ([string]::IsNullOrWhiteSpace($l)) { continue }
        if ($l -eq "#EXTM3U") { continue }
        if ($l.StartsWith("#EXTINF")) {
            if ($l -match '#EXTINF:-?\d+,(.*)') {
                $infPendente = $matches[1].Trim()
                $infPendente = [regex]::Replace($infPendente, '^(\s*\[NAO ENCONTRADO\]\s*)+', '', 'IgnoreCase').Trim()
            } else { $infPendente = $null }
            continue
        }
        if ($l.StartsWith("#")) { continue }
        $entries += [PSCustomObject]@{ Titulo = $infPendente; PathOriginal = $l }
        $infPendente = $null
    }
    return $entries
}

function Read-NaoEncontrados($m3uPath) {
    if (-not (Test-Path -LiteralPath $m3uPath)) { return @() }
    $linhas = Get-Content -Path $m3uPath -Encoding UTF8 -ErrorAction SilentlyContinue
    $resultado = @()
    $dentroDaSecao = $false
    $tituloPendente = $null
    foreach ($linhaCrua in $linhas) {
        $l = $linhaCrua.Trim()
        if ($l -like "*NÃO ENCONTRADOS*") { $dentroDaSecao = $true; continue }
        if (-not $dentroDaSecao) { continue }
        if ([string]::IsNullOrWhiteSpace($l)) { continue }
        if ($l.StartsWith("#EXTINF")) {
            if ($l -match '#EXTINF:-?\d+,\[NAO ENCONTRADO\]\s*(.*)') { $tituloPendente = $matches[1].Trim() }
            elseif ($l -match '#EXTINF:-?\d+,(.*)') { $tituloPendente = $matches[1].Trim() }
            if ($tituloPendente) {
                $tituloPendente = [regex]::Replace($tituloPendente, '^(\s*\[NAO ENCONTRADO\]\s*)+', '', 'IgnoreCase').Trim()
            }
            continue
        }
        if ($l.StartsWith("#")) { continue }
        $resultado += [PSCustomObject]@{ Titulo = $tituloPendente; PathOriginal = $l }
        $tituloPendente = $null
    }
    return $resultado
}

function Resolve-Entrada($entrada, $m3uDir, $indice) {
    try {
        $original = $entrada.PathOriginal
        if (Test-Path -LiteralPath $original -PathType Leaf -ErrorAction Stop) {
            return @{ Arquivo = (Get-Item -LiteralPath $original); Metodo = "caminho original" }
        }
    } catch { }
    try {
        $tentativaRelativa = Join-Path $m3uDir $original
        if (Test-Path -LiteralPath $tentativaRelativa -PathType Leaf -ErrorAction Stop) {
            return @{ Arquivo = (Get-Item -LiteralPath $tentativaRelativa); Metodo = "caminho relativo" }
        }
    } catch { }
    $nomeArquivo = $null
    try { $nomeArquivo = Split-Path $original -Leaf } catch { }
    if ($nomeArquivo -and $indice.PorNome.ContainsKey($nomeArquivo.ToLowerInvariant())) {
        return @{ Arquivo = $indice.PorNome[$nomeArquivo.ToLowerInvariant()]; Metodo = "nome exato" }
    }
    if ($nomeArquivo) {
        $baseSemExt = [System.IO.Path]::GetFileNameWithoutExtension($nomeArquivo)
        $chave = Get-NormalizedKey $baseSemExt
        if ($chave -and $indice.PorChave.ContainsKey($chave)) {
            return @{ Arquivo = $indice.PorChave[$chave]; Metodo = "nome normalizado" }
        }
    }
    if ($entrada.Titulo) {
        $chaveTitulo = Get-NormalizedKey $entrada.Titulo
        if ($chaveTitulo -and $indice.PorChave.ContainsKey($chaveTitulo)) {
            return @{ Arquivo = $indice.PorChave[$chaveTitulo]; Metodo = "título do EXTINF" }
        }
    }
    return $null
}

# ============================================================
#  MÓDULO 1 — GERENCIAR PLAYLIST
# ============================================================
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

        try {
            $entries = Import-Csv -LiteralPath $csvPath -Encoding UTF8 -ErrorAction Stop
        }
        catch {
            try {
                $entries = Import-Csv -LiteralPath $csvPath -Encoding Default
            }
            catch {
                Write-Warning "Nao foi possivel ler CSV: $csvPath"
                continue
            }
        }

        Write-Success "CSV carregado: $($entries.Count) entradas"
        Write-Success "Pasta validada: $folder"
        Write-Info "Próximo passo: execute o Módulo 4 para baixar as músicas faltantes"
    }
}

# ============================================================
#  MÓDULO 2 — REMOVER DUPLICATAS
# ============================================================
function Invoke-Modulo2-RemoverDuplicatas {
    Clear-Host
    Write-Header "MODULO 2 - REMOVER DUPLICATAS"

    $pastas = @()
    foreach ($p in $playlists) {
        if ($p.Folder) { $pastas += $p.Folder }
    }

    $totalDuplicatas = 0
    $totalEspacoLiberado = 0

    foreach ($pasta in $pastas) {
        if (-not (Test-Path -LiteralPath $pasta)) { continue }

        Write-Section "Processando: $pasta"
        $arquivos = @(Get-ChildItem -LiteralPath $pasta -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -ieq ".m4a" })
        if ($arquivos.Count -le 1) { 
            Write-Success "0 ou 1 arquivo, nada a fazer"
            continue 
        }

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
                        $totalEspacoLiberado += $candidato.Length
                    }
                    else {
                        $remover += $arquivo.FullName
                        $processados[$arquivo.FullName] = $true
                        $totalEspacoLiberado += $arquivo.Length
                        break
                    }
                }
            }
        }

        if ($remover.Count -eq 0) {
            Write-Success "Nenhuma duplicata encontrada"
            continue
        }

        $remover = $remover | Sort-Object -Unique
        $totalDuplicatas += $remover.Count

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

    Write-Header "MODULO 2 CONCLUIDO"
    Write-Success "Total de duplicatas: $totalDuplicatas"
    $mb = [math]::Round($totalEspacoLiberado / 1MB, 2)
    Write-Success "Espaço que poderia ser liberado: $mb MB"
}

# ============================================================
#  MÓDULO 3 — CORRIGIR M3U
# ============================================================
function Invoke-Corrigir($nomePlaylist, $originalPath, $indice, $destinoPath) {
    Write-Header "PLAYLIST: $nomePlaylist"

    if (-not (Test-Path -LiteralPath $originalPath)) {
        Write-ErrorMsg "M3U original não encontrado: $originalPath"
        return
    }

    $m3uDir = Split-Path $originalPath -Parent
    $entradas = Read-M3UFull $originalPath
    Write-Info "Entradas no ORIGINAL: $($entradas.Count)"

    $resolvidos     = @()
    $naoEncontrados = @()
    $corrigidos     = 0
    $jaCorretos     = 0

    foreach ($entrada in $entradas) {
        $r = Resolve-Entrada $entrada $m3uDir $indice

        if (-not $r) {
            $naoEncontrados += $entrada
            $sufixoTitulo = if ($entrada.Titulo) { "  (título: $($entrada.Titulo))" } else { "" }
            Write-ErrorMsg "Não encontrado: $($entrada.PathOriginal)$sufixoTitulo"
            continue
        }

        $arquivo = $r.Arquivo

        if ($r.Metodo -eq "caminho original") {
            $jaCorretos++
        } else {
            $corrigidos++
            Write-Fix "Corrigido ($($r.Metodo)): `"$($entrada.PathOriginal)`""
        }

        $resolvidos += $arquivo
    }

    $resolvidosUnicos = $resolvidos | Sort-Object FullName -Unique

    $resolvidosOrdenados = $resolvidosUnicos | Sort-Object `
        @{ Expression = { Get-ChaveOrdenacao (Get-Artista $_.BaseName) } }, `
        @{ Expression = { Get-ChaveOrdenacao $_.BaseName } }

    if (Test-Path -LiteralPath $destinoPath) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $backupPath = "$destinoPath.bak_$timestamp"
        Copy-Item -LiteralPath $destinoPath -Destination $backupPath -Force
        Write-Info "Backup criado: $backupPath"
    }

    $linhasFinais = @("#EXTM3U")
    foreach ($f in $resolvidosOrdenados) {
        $linhasFinais += "#EXTINF:-1,$($f.BaseName)"
        $linhasFinais += $f.FullName
    }
    if ($naoEncontrados.Count -gt 0) {
        $linhasFinais += "# ---- NÃO ENCONTRADOS NO ORIGINAL (verifique manualmente / baixe) ----"
        foreach ($n in $naoEncontrados) {
            $rotulo = if ($n.Titulo) { $n.Titulo } else { $n.PathOriginal }
            $linhasFinais += "#EXTINF:-1,[NAO ENCONTRADO] $rotulo"
            $linhasFinais += $n.PathOriginal
        }
    }
    $linhasFinais | Set-Content -Path $destinoPath -Encoding UTF8

    Write-Section "Resumo"
    Write-Success "Já corretos       : $jaCorretos"
    Write-Fix     "Corrigidos        : $corrigidos"
    if ($naoEncontrados.Count -gt 0) {
        Write-ErrorMsg "Não encontrados   : $($naoEncontrados.Count) (marcados no final)"
    } else {
        Write-Success "Não encontrados   : 0"
    }
    Write-Success "Escrito → $destinoPath"
}

function Invoke-Modulo3-CorrigirM3u {
    Clear-Host
    Write-Header "MODULO 3 - CORRIGIR M3U"

    foreach ($pl in $playlists) {
        if (-not (Test-Path $pl.Folder)) {
            Write-ErrorMsg "Pasta da playlist não encontrada: $($pl.Folder)"
            continue
        }
        $indice = Build-Indice @($pl.Folder)
        Write-Info "[$($pl.Nome)] arquivos indexados: $($indice.PorNome.Count)"

        $safeName = Get-SafeName $pl.Nome
        $destino = Join-Path $pl.Folder "$safeName.m3u"

        Invoke-Corrigir $pl.Nome $pl.OriginalM3u $indice $destino
    }

    Ensure-Folder $pastaExtrasCombinadas
    $pastasCombinadas = ($playlists | ForEach-Object { $_.Folder }) + $pastaExtrasCombinadas
    $indiceCombinado = Build-Indice $pastasCombinadas
    Write-Info "Índice combinado: $($indiceCombinado.PorNome.Count) arquivos"

    foreach ($pl in $playlistsCombinadas) {
        Invoke-Corrigir $pl.Nome $pl.M3u $indiceCombinado $pl.M3u
    }

    Write-Header "MODULO 3 CONCLUIDO"
}

# ============================================================
#  MÓDULO 4 — BAIXAR FALTANTES
# ============================================================
function Invoke-Modulo4-BaixarFaltantes {
    Clear-Host
    Write-Header "MODULO 4 - BAIXAR FALTANTES"

    if (-not (Test-Spotdl)) {
        Write-ErrorMsg "spotdl nao encontrado"
        return
    }

    Ensure-Folder $pastaExtrasCombinadas

    $todasFontes = @()
    foreach ($pl in $playlists) {
        $safeName = Get-SafeName $pl.Nome
        $m3uGerado = Join-Path $pl.Folder "$safeName.m3u"
        $todasFontes += @{ Nome = $pl.Nome; M3u = $m3uGerado; Pasta = $pl.Folder }
    }
    foreach ($pl in $playlistsCombinadas) {
        $todasFontes += @{ Nome = $pl.Nome; M3u = $pl.M3u; Pasta = $pastaExtrasCombinadas }
    }

    $faltantesPorChave = @{}
    foreach ($fonte in $todasFontes) {
        $entradas = Read-NaoEncontrados $fonte.M3u
        foreach ($e in $entradas) {
            $rotuloCru = if ($e.Titulo) { $e.Titulo } else { $e.PathOriginal }
            $rotulo = Get-NomeLimpo $rotuloCru
            $chave = Get-NormalizedKey $rotulo
            if (-not $chave) { continue }
            if (-not $faltantesPorChave.ContainsKey($chave)) {
                $faltantesPorChave[$chave] = [PSCustomObject]@{ 
                    Rotulo = $rotulo
                    Playlists = New-Object System.Collections.Generic.List[string]
                    Pasta = $fonte.Pasta
                }
            }
            $faltantesPorChave[$chave].Playlists.Add($fonte.Nome)
        }
    }

    if ($faltantesPorChave.Count -eq 0) {
        Write-Success "Nenhuma música marcada como não encontrada"
        return
    }

    Write-Info "Total de músicas faltantes: $($faltantesPorChave.Count)"
    Write-Section "Músicas não encontradas:"

    $i = 1
    foreach ($chave in ($faltantesPorChave.Keys | Sort-Object)) {
        $info = $faltantesPorChave[$chave]
        $playlistsTxt = ($info.Playlists | Select-Object -Unique) -join ", "
        Write-Host "$i. $($info.Rotulo) [$playlistsTxt]" -ForegroundColor Yellow
        $i++
    }

    Write-Host ""
    $resposta = Read-Host "Deseja baixar alguma? Digite S para sim ou N para não"

    if ($resposta.ToUpper() -ne "S") {
        Write-Info "Cancelado"
        return
    }

    Write-Host ""
    Write-Section "Escolha quantas quiser (números separados por vírgula, ex: 1,3,5)"
    Write-Info "Deixe em branco para todas"
    $numeros_str = Read-Host "Números"

    if ([string]::IsNullOrWhiteSpace($numeros_str)) {
        $chavesEscolhidas = @($faltantesPorChave.Keys | Sort-Object)
    } else {
        $numeros = $numeros_str -split "[,\s]+" | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ - 1 }
        $chavesEscolhidas = @()
        $i = 0
        foreach ($chave in ($faltantesPorChave.Keys | Sort-Object)) {
            if ($numeros -contains $i) { $chavesEscolhidas += $chave }
            $i++
        }
    }

    if ($chavesEscolhidas.Count -eq 0) {
        Write-Warning "Nenhuma música selecionada"
        return
    }

    Write-Host ""
    Write-Section "Baixando $($chavesEscolhidas.Count) música(s)..."

    $sucessos = 0
    $falhas = 0
    $i = 1

    foreach ($chave in $chavesEscolhidas) {
        $info = $faltantesPorChave[$chave]
        Write-Host "`n[$i/$($chavesEscolhidas.Count)] $($info.Rotulo)" -ForegroundColor White
        $i++

        Ensure-Folder $info.Pasta
        $cookiePath = Join-Path $info.Pasta "cookies.txt"

        Push-Location $info.Pasta
        try {
            $antes = Get-AudioFileCount $info.Pasta

            Write-Info "Tentando via spotdl..."
            if (Test-Path $cookiePath) {
                spotdl download "$($info.Rotulo)" --format $Format --bitrate disable --threads $Threads --cookie-file "cookies.txt" --overwrite skip 2>&1 | ForEach-Object { Write-Info $_ }
            } else {
                spotdl download "$($info.Rotulo)" --format $Format --bitrate disable --threads $Threads --overwrite skip 2>&1 | ForEach-Object { Write-Info $_ }
            }

            $depois = Get-AudioFileCount $info.Pasta

            if ($depois -gt $antes) {
                Write-Success "Baixado com sucesso"
                $sucessos++
            }
            else {
                Write-Warning "Não encontrado no Spotify"

                if (Test-YtDlp) {
                    $ytLink = Read-Host "Cole o link do YouTube (ou deixe em branco para pular)"
                    if (-not [string]::IsNullOrWhiteSpace($ytLink)) {
                        Write-Info "Baixando via YouTube..."
                        $ytCmd = "yt-dlp -x --audio-format $Format --audio-quality 0 --embed-thumbnail --add-metadata -o `"%(title)s.%(ext)s`" `"$($ytLink.Trim())`""
                        if (Test-Path $cookiePath) { $ytCmd += " --cookies `"cookies.txt`"" }
                        Invoke-Expression $ytCmd

                        $depoisYt = Get-AudioFileCount $info.Pasta
                        if ($depoisYt -gt $antes) {
                            Write-Success "Baixado via YouTube"
                            $sucessos++
                        } else {
                            Write-ErrorMsg "Falha no download do YouTube"
                            $falhas++
                        }
                    } else {
                        Write-Warning "Pulado"
                        $falhas++
                    }
                } else {
                    Write-ErrorMsg "yt-dlp não encontrado. Instale com: pip install -U yt-dlp"
                    $falhas++
                }
            }
        }
        catch {
            Write-ErrorMsg "Erro: $_"
            $falhas++
        }
        finally {
            Pop-Location
        }
    }

    Write-Header "MODULO 4 CONCLUIDO"
    Write-Success "Baixadas com sucesso: $sucessos"
    if ($falhas -gt 0) { Write-ErrorMsg "Falharam: $falhas" }
    Write-Info "Execute o Módulo 3 de novo para atualizar os m3u's com as novas músicas"
}

# ============================================================
#  MENU PRINCIPAL
# ============================================================
function Show-Menu {
    Clear-Host
    Write-Header "PLAYLIST MANAGER - MENU PRINCIPAL"
    Write-Host ""
    Write-Host "1. Gerenciar playlist (CSV) + validar" -ForegroundColor Yellow
    Write-Host "2. Remover duplicatas" -ForegroundColor Yellow
    Write-Host "3. Corrigir e ordenar M3U por artista" -ForegroundColor Yellow
    Write-Host "4. Baixar músicas faltantes" -ForegroundColor Yellow
    Write-Host "5. Fluxo completo (2 > 3 > 4)" -ForegroundColor Yellow
    Write-Host "0. Sair" -ForegroundColor DarkGray
    Write-Host ""
}

$sair = $false
while (-not $sair) {
    Show-Menu
    $escolha = Read-Host "Escolha uma opção"

    switch ($escolha) {
        "1" { Invoke-Modulo1-GerenciarPlaylists }
        "2" { Invoke-Modulo2-RemoverDuplicatas }
        "3" { Invoke-Modulo3-CorrigirM3u }
        "4" { Invoke-Modulo4-BaixarFaltantes }
        "5" {
            Write-Header "FLUXO COMPLETO"
            Invoke-Modulo2-RemoverDuplicatas
            Write-Host ""
            Write-Info "Pressione Enter para seguir pro Módulo 3..."
            Read-Host | Out-Null
            Invoke-Modulo3-CorrigirM3u
            Write-Host ""
            Write-Info "Pressione Enter para seguir pro Módulo 4..."
            Read-Host | Out-Null
            Invoke-Modulo4-BaixarFaltantes
        }
        "0" { $sair = $true }
        default { Write-ErrorMsg "Opção inválida" }
    }

    if (-not $sair) {
        Write-Host ""
        Write-Info "Pressione Enter para voltar ao menu..."
        Read-Host | Out-Null
    }
}

Write-Host ""
Write-Host "Até mais!" -ForegroundColor Cyan
