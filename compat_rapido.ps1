# RedM Compat Rapido / RedM Compat Quick - chequeo simple de compatibilidad para el servidor del cliente.
# No instala nada: usa PowerShell y Windows Forms, que vienen con Windows. Solo LEE archivos: no modifica,
# borra ni mueve nada, no se conecta a internet y no envía datos.
#
# AVISO / DISCLAIMER: se entrega "tal cual", sin garantías. El reporte es informativo y puede no detectar todos
# los problemas. Se usa bajo la responsabilidad de quien lo ejecuta. Provided "as is", without warranty; the
# report is informational and may not detect every issue. Use at your own risk.
#
# Uso normal: doble clic en "Generar reporte de compatibilidad.bat", elegir el asset, la carpeta resources y
# los .cfg, y tocar "Generar reporte". Queda un .txt en la carpeta "reportes" (al lado de este script).
# Uso por consola (sin ventana):
#   powershell -ExecutionPolicy Bypass -File compat_rapido.ps1 -Asset <carpeta> -Resources <carpeta> [-Cfg a.cfg,b.cfg] [-Out <archivo.txt>] [-Lang es|en] -NoGui
param(
    [string]$Asset = "",
    [string]$Resources = "",
    [string[]]$Cfg = @(),
    [string]$Out = "",
    [switch]$NoGui,
    [string]$Lang = ""
)

$ErrorActionPreference = "Stop"
$Version = "1.1.0"
# -Cfg puede venir como "a.cfg,b.cfg" o "a.cfg; b.cfg"
$Cfg = @($Cfg | ForEach-Object { $_ -split '[,;]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })

# ------------------------------------------------------------------ idioma / language

if (-not $Lang) {
    $Lang = $(if ((Get-Culture).TwoLetterISOLanguageName -eq "es") { "es" } else { "en" })
}
$script:Lang = $(if ($Lang -eq "es") { "es" } else { "en" })

$TXT = @{
    es = @{
        p_asset = "Leyendo el asset..."
        p_find = "Buscando recursos..."
        p_cfg = "Leyendo los .cfg..."
        p_res = "Revisando recursos ({0}/{1})..."
        h_other = "[{0}] {1} quiere reemplazar el mismo archivo que {2} ({3}), pero en minúsculas su nombre da {4}: NO lo reemplaza. Parece un intento de compatibilidad mal nombrado."
        h_mine = "[{0}] El archivo {1} del asset dice {2}, pero en RedM (minúsculas) es {3}: NO reemplaza al archivo {2}. Hay que renombrarlo con una colisión en minúsculas."
        h_who = "    y estos recursos SÍ reemplazan {0}: {1}"
        title = "REPORTE DE COMPATIBILIDAD  -  RedM Compat Rápido v{0}"
        date = "Fecha: {0}   Equipo: {1}"
        asset = "Asset revisado: {0}"
        folder = "  Carpeta: {0}"
        nstream = "  Archivos en stream: {0}"
        loaded_yes = "  Se carga: SÍ (posición {0} de {1} en el orden de arranque)"
        loaded_no = "  Se carga: NO APARECE en los .cfg (ensure/start). Si no está, el asset no se carga."
        resfolder = "Carpeta de recursos: {0}  ({1} recursos encontrados)"
        cfgread = "Cfg leídos: {0}"
        cfgmiss = "Exec que no se encontraron: {0}"
        catused = "Categorías arrancadas enteras: {0}"
        catunk = "Categorías del cfg que no están en la carpeta: {0}"
        nocfg = "Cfg: no se eligió ninguno -> no se sabe qué recursos se cargan ni cuál gana."
        ok = "RESULTADO: OK - no hay choques de archivos con otros recursos que se carguen."
        bad = "RESULTADO: HAY PROBLEMAS DE COMPATIBILIDAD - {0} recurso(s) con archivos distintos que chocan{1}."
        bad_hash = ", y avisos de nombres con hash"
        need = "RECURSOS QUE CHOCAN CON EL ASSET (el creador del asset va a necesitar estas carpetas para hacer la compatibilidad):"
        dupmark = "  [NOMBRE REPETIDO: el servidor usa una sola carpeta con este nombre]"
        needline = "  - {0}   ({1} archivo(s) distintos: {2}){3}"
        same = "Chocan pero con archivos IDÉNTICOS (no hay problema): {0}"
        ymf = "Solo chocan los _manifest.ymf (aviso: cada _manifest.ymf debería tener un nombre único; no está confirmado que cause problemas): {0}"
        notloaded = "Chocan pero NO se cargan (no están en los .cfg, no afectan): {0}"
        hashhdr = "NOMBRES CON HASH (_0x...)"
        hashexp = "RedM pasa el nombre del archivo a minúsculas antes de calcular su hash."
        detail = "DETALLE POR ARCHIVO"
        dline = "* {0}   (tipo {1}, clave {2})"
        starts = "  [arranca #{0}]"
        notl = "  [NO se carga]"
        enc = "  [cifrado]"
        theasset = "{0} (el asset)"
        identical = "IDÉNTICO"
        different = "DISTINTO"
        dupwin = "{0} (hay {1} carpetas con ese nombre; el servidor usa una sola)"
        wins = "    -> Se carga (normalmente el que arranca último): {0}"
        ymfnote = "    (los _manifest.ymf necesitan un nombre único por recurso)"
        dupmine = "ARCHIVOS REPETIDOS DENTRO DEL ASSET (RedM carga uno solo):"
        dupres = "RECURSOS CON EL MISMO NOMBRE EN DOS CARPETAS (el servidor usa uno solo):"
        end = "Fin del reporte."
        saved = "Reporte guardado en: {0}"
        dir = "reportes"
        g_title = "RedM Compat Rápido v{0}"
        g_lang = "Idioma:"
        g_a = "1. El asset a revisar (la carpeta del recurso que te pasó el creador):"
        g_r = "2. La carpeta resources de tu servidor:"
        g_c = "3. Los .cfg del servidor (server.cfg; si tenés resources.cfg u otros, elegilos también):"
        g_pick = "Elegir..."
        g_go = "Generar reporte"
        g_status = "Elegí las 3 cosas y tocá Generar reporte. No se modifica ningún archivo del servidor."
        d_asset = "Elegí la carpeta del asset a revisar"
        d_res = "Elegí la carpeta resources del servidor"
        d_cfg = "Elegí los .cfg del servidor (podés elegir varios con Ctrl)"
        d_filter = "Archivos cfg (*.cfg)|*.cfg|Todos (*.*)|*.*"
        m_noasset = "Elegí la carpeta del asset."
        m_nores = "Elegí la carpeta resources."
        m_found = "Se encontraron problemas de compatibilidad."
        m_none = "No se encontraron problemas de compatibilidad."
        m_where = "El reporte quedó en:"
        m_send = "Mandale este archivo al creador del asset."
        m_ready = "Reporte listo"
        m_err = "Error: {0}"
        status_done = "Listo: {0}"
        disc_hdr = "AVISO"
        disc = "Esta herramienta solo LEE archivos: no modifica, borra ni mueve nada del servidor, no se conecta a internet y no envía datos. Se entrega ""tal cual"", sin garantías. El reporte es informativo, se basa en los nombres de archivo y en los .cfg elegidos, y puede no detectar todos los problemas (por ejemplo, conflictos dentro de archivos cifrados o de scripts). Quien la ejecuta lo hace bajo su propia responsabilidad; el autor no se hace responsable por daños, pérdidas ni decisiones tomadas a partir de este reporte. Antes de cambiar algo en el servidor, hacé una copia de seguridad."
        g_disc = "Solo lee archivos, no modifica nada ni usa internet. Se entrega tal cual, sin garantías; usala bajo tu responsabilidad."
    }
    en = @{
        p_asset = "Reading the asset..."
        p_find = "Looking for resources..."
        p_cfg = "Reading the .cfg files..."
        p_res = "Checking resources ({0}/{1})..."
        h_other = "[{0}] {1} tries to replace the same file as {2} ({3}), but in lowercase its name hashes to {4}: it does NOT replace it. Looks like a badly named compatibility file."
        h_mine = "[{0}] The asset file {1} says {2}, but in RedM (lowercase) it is {3}: it does NOT replace file {2}. It must be renamed with a lowercase collision name."
        h_who = "    and these resources DO replace {0}: {1}"
        title = "COMPATIBILITY REPORT  -  RedM Compat Quick v{0}"
        date = "Date: {0}   Computer: {1}"
        asset = "Asset checked: {0}"
        folder = "  Folder: {0}"
        nstream = "  Files in stream: {0}"
        loaded_yes = "  Loaded: YES (position {0} of {1} in the start order)"
        loaded_no = "  Loaded: NOT FOUND in the .cfg files (ensure/start). If it is not there, the asset does not load."
        resfolder = "Resources folder: {0}  ({1} resources found)"
        cfgread = "Cfg files read: {0}"
        cfgmiss = "Exec files not found: {0}"
        catused = "Whole categories started: {0}"
        catunk = "Categories in the cfg that are not in the folder: {0}"
        nocfg = "Cfg: none selected -> it is unknown which resources load and which file wins."
        ok = "RESULT: OK - no file conflicts with other loaded resources."
        bad = "RESULT: COMPATIBILITY PROBLEMS FOUND - {0} resource(s) with different files that conflict{1}."
        bad_hash = ", plus hashed-name warnings"
        need = "RESOURCES THAT CONFLICT WITH THE ASSET (the asset creator will need these folders to make them compatible):"
        dupmark = "  [DUPLICATE NAME: the server uses only one folder with this name]"
        needline = "  - {0}   ({1} different file(s): {2}){3}"
        same = "Conflict but with IDENTICAL files (no problem): {0}"
        ymf = "Only the _manifest.ymf files conflict (note: each _manifest.ymf should have a unique name; not confirmed to cause problems): {0}"
        notloaded = "Conflict but are NOT loaded (not in the .cfg files, no effect): {0}"
        hashhdr = "HASHED NAMES (_0x...)"
        hashexp = "RedM lowercases the file name before computing its hash."
        detail = "DETAILS PER FILE"
        dline = "* {0}   (type {1}, key {2})"
        starts = "  [starts #{0}]"
        notl = "  [NOT loaded]"
        enc = "  [encrypted]"
        theasset = "{0} (the asset)"
        identical = "IDENTICAL"
        different = "DIFFERENT"
        dupwin = "{0} ({1} folders have that name; the server uses only one)"
        wins = "    -> Loaded (usually the one that starts last): {0}"
        ymfnote = "    (_manifest.ymf files need a unique name per resource)"
        dupmine = "DUPLICATE FILES INSIDE THE ASSET (RedM loads only one):"
        dupres = "RESOURCES WITH THE SAME NAME IN TWO FOLDERS (the server uses only one):"
        end = "End of report."
        saved = "Report saved to: {0}"
        dir = "reports"
        g_title = "RedM Compat Quick v{0}"
        g_lang = "Language:"
        g_a = "1. The asset to check (the resource folder the creator gave you):"
        g_r = "2. Your server resources folder:"
        g_c = "3. The server .cfg files (server.cfg; if you also have resources.cfg or others, select them too):"
        g_pick = "Browse..."
        g_go = "Create report"
        g_status = "Pick the 3 items and click Create report. No server file is modified."
        d_asset = "Select the folder of the asset to check"
        d_res = "Select the server resources folder"
        d_cfg = "Select the server .cfg files (hold Ctrl to select several)"
        d_filter = "cfg files (*.cfg)|*.cfg|All files (*.*)|*.*"
        m_noasset = "Select the asset folder."
        m_nores = "Select the resources folder."
        m_found = "Compatibility problems were found."
        m_none = "No compatibility problems were found."
        m_where = "The report was saved to:"
        m_send = "Send this file to the asset creator."
        m_ready = "Report ready"
        m_err = "Error: {0}"
        status_done = "Done: {0}"
        disc_hdr = "DISCLAIMER"
        disc = "This tool only READS files: it does not modify, delete or move anything on the server, does not connect to the internet and does not send any data. It is provided ""as is"", without warranty of any kind. The report is informational, is based on file names and on the selected .cfg files, and may not detect every issue (for example, conflicts inside encrypted files or scripts). Whoever runs it does so at their own risk; the author is not liable for any damage, loss or decision made based on this report. Back up your server before changing anything."
        g_disc = "Read-only: it does not modify anything or use the internet. Provided as is, without warranty; use at your own risk."
    }
}

function T([string]$key) { return $TXT[$script:Lang][$key] }


$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$NotStreamed = @(".lua", ".md", ".txt", ".json", ".xml", ".fxap", ".cfg", ".bat", ".js", ".html", ".png", ".jpg",
                 ".jpeg", ".zip", ".rar", ".7z", ".ini", ".log", ".meta", ".ps1", ".css", ".ttf", ".otf", ".ogg",
                 ".mp3", ".wav", ".svg", ".gif", ".webp", ".sql", ".yml", ".yaml", ".ts", ".map")

Add-Type -TypeDefinition @"
using System;
using System.Text;
public static class RcJoaat {
    public static uint Hash(string s) {
        uint h = 0;
        foreach (byte b in Encoding.UTF8.GetBytes(s.ToLowerInvariant())) {
            h += b; h += h << 10; h ^= h >> 6;
        }
        h += h << 3; h ^= h >> 11; h += h << 15;
        return h;
    }
}
"@

function Hex([uint32]$h) { return ("0x{0:X8}" -f $h) }

function Size-Text([long]$n) {
    if ($n -ge 1048576) { return ("{0:N1} MB" -f ($n / 1048576.0)) }
    return ("{0} KB" -f [Math]::Max(1, [Math]::Round($n / 1024.0)))
}

function Is-Escrow([string]$path) {
    try {
        $fs = [IO.File]::OpenRead($path)
        try { $b = New-Object byte[] 4; [void]$fs.Read($b, 0, 4) } finally { $fs.Close() }
        return ([Text.Encoding]::ASCII.GetString($b) -eq "FXAP")
    } catch { return $false }
}

function Md5([string]$path) {
    return (Get-FileHash -Algorithm MD5 -LiteralPath $path).Hash
}

# ------------------------------------------------------------------ recursos

function Find-Resources([string]$root, [string]$exclude) {
    $out = New-Object System.Collections.Generic.List[string]
    $ex = ""
    if ($exclude) { $ex = [IO.Path]::GetFullPath($exclude).TrimEnd('\').ToLowerInvariant() }
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push([IO.Path]::GetFullPath($root))
    while ($stack.Count -gt 0) {
        $d = $stack.Pop()
        $dl = $d.TrimEnd('\').ToLowerInvariant()
        if ($ex -and ($dl -eq $ex -or $dl.StartsWith($ex + "\"))) { continue }
        try {
            $files = [IO.Directory]::GetFiles($d) | ForEach-Object { [IO.Path]::GetFileName($_).ToLowerInvariant() }
            $subs = [IO.Directory]::GetDirectories($d)
        } catch { continue }
        $isRes = ($files -contains "fxmanifest.lua") -or ($files -contains "__resource.lua")
        if (-not $isRes) {
            foreach ($s in $subs) { if ([IO.Path]::GetFileName($s).ToLowerInvariant() -eq "stream") { $isRes = $true } }
        }
        if ($isRes) { $out.Add($d); continue }
        foreach ($s in $subs) {
            $n = [IO.Path]::GetFileName($s).ToLowerInvariant()
            if ($n -in @(".git", "node_modules", "cache") -or $n.StartsWith("compat_") -or $n.EndsWith("_codex")) { continue }
            $stack.Push($s)
        }
    }
    return ,$out
}

function Stream-Files([string]$resPath) {
    $list = New-Object System.Collections.Generic.List[object]
    $base = [IO.Path]::GetFullPath($resPath).TrimEnd('\')
    try { $all = [IO.Directory]::EnumerateFiles($base, "*", [IO.SearchOption]::AllDirectories) } catch { return ,$list }
    foreach ($f in $all) {
        $rel = $f.Substring($base.Length + 1)
        $parts = $rel.ToLowerInvariant().Split('\')
        if ($parts.Length -lt 2 -or -not ($parts[0..($parts.Length - 2)] -contains "stream")) { continue }
        $name = [IO.Path]::GetFileName($f)
        $ext = [IO.Path]::GetExtension($name).ToLowerInvariant()
        if (-not $ext -or $NotStreamed -contains $ext) { continue }
        $stem = [IO.Path]::GetFileNameWithoutExtension($name)
        $key = [RcJoaat]::Hash($stem)
        $intended = $null
        if ($stem -match '_0x([0-9a-fA-F]{8})$') { $intended = [Convert]::ToUInt32($Matches[1], 16) }
        $list.Add([pscustomobject]@{ Path = $f; Rel = $rel; Name = $name; Ext = $ext; Key = $key; Intended = $intended })
    }
    return ,$list
}

# ------------------------------------------------------------------ cfg

function Cfg-Tokens([string]$line) {
    $out = New-Object System.Collections.Generic.List[string]
    $cur = ""; $q = $null; $i = 0
    while ($i -lt $line.Length) {
        $c = $line[$i]
        if ($q) { if ($c -eq $q) { $q = $null } else { $cur += $c } }
        elseif ($c -eq '"' -or $c -eq "'") { $q = $c }
        elseif ($c -eq '#' -or ($c -eq '/' -and $i + 1 -lt $line.Length -and $line[$i + 1] -eq '/')) { break }
        elseif ([char]::IsWhiteSpace($c)) { if ($cur) { $out.Add($cur); $cur = "" } }
        else { $cur += $c }
        $i++
    }
    if ($cur) { $out.Add($cur) }
    return ,$out
}

function Parse-Cfgs([string[]]$cfgFiles, $categories) {
    $state = @{ Order = New-Object System.Collections.Generic.List[string]; Pos = @{}; Files = New-Object System.Collections.Generic.List[string];
                Missing = New-Object System.Collections.Generic.List[string]; Stopped = @{}; CatUsed = New-Object System.Collections.Generic.List[string];
                CatUnknown = New-Object System.Collections.Generic.List[string] }
    if (-not $cfgFiles -or $cfgFiles.Count -eq 0) { return $state }
    $base = Split-Path -Parent ([IO.Path]::GetFullPath($cfgFiles[0]))
    $startRes = {
        param($n)
        $n = $n.ToLowerInvariant()
        $state.Stopped.Remove($n)
        if (-not $state.Pos.ContainsKey($n)) { $state.Pos[$n] = $state.Order.Count; $state.Order.Add($n) }
    }
    $run = $null
    $run = {
        param($p)
        $full = [IO.Path]::GetFullPath($p)
        foreach ($f in $state.Files) { if ($f -ieq $full) { return } }
        $state.Files.Add($full)
        foreach ($line in [IO.File]::ReadAllLines($full)) {
            $t = Cfg-Tokens $line
            if ($t.Count -lt 2) { continue }
            $cmd = $t[0].ToLowerInvariant(); $arg = $t[1]
            if ($cmd -in @("ensure", "start", "restart")) {
                if ($arg.StartsWith("[") -and $arg.EndsWith("]")) {
                    $members = $categories[$arg.ToLowerInvariant()]
                    if ($members) { $state.CatUsed.Add($arg); foreach ($m in ($members | Sort-Object)) { & $startRes $m } }
                    else { $state.CatUnknown.Add($arg) }
                } else { & $startRes $arg }
            } elseif ($cmd -eq "stop") { $state.Stopped[$arg.ToLowerInvariant()] = $true }
            elseif ($cmd -eq "exec") {
                $rel = $arg.Replace('/', '\')
                $hit = $null
                foreach ($c in @((Join-Path $base $rel), (Join-Path (Split-Path -Parent $full) $rel))) {
                    if (-not $hit -and (Test-Path -LiteralPath $c -PathType Leaf)) { $hit = $c }
                }
                if ($hit) { & $run $hit } else { $state.Missing.Add($arg) }
            }
        }
    }
    foreach ($c in $cfgFiles) { if ($c -and (Test-Path -LiteralPath $c -PathType Leaf)) { & $run $c } }
    foreach ($s in @($state.Stopped.Keys)) { if ($state.Pos.ContainsKey($s)) { $state.Pos.Remove($s) } }
    return $state
}

# ------------------------------------------------------------------ análisis

function Run-Check([string]$assetPath, [string]$resRoot, [string[]]$cfgFiles, [scriptblock]$progress) {
    $assetPath = [IO.Path]::GetFullPath($assetPath).TrimEnd('\')
    $resRoot = [IO.Path]::GetFullPath($resRoot).TrimEnd('\')
    $assetName = [IO.Path]::GetFileName($assetPath)
    $L = New-Object System.Collections.Generic.List[string]
    $w = { param($s) $L.Add($s) }

    & $progress (T "p_asset")
    $mine = Stream-Files $assetPath
    $mineMap = @{}
    foreach ($f in $mine) { $k = "$($f.Ext)|$($f.Key)"; if (-not $mineMap.ContainsKey($k)) { $mineMap[$k] = New-Object System.Collections.Generic.List[object] }; $mineMap[$k].Add($f) }

    & $progress (T "p_find")
    $resPaths = Find-Resources $resRoot $assetPath
    $categories = @{}
    foreach ($p in $resPaths) {
        $nm = [IO.Path]::GetFileName($p).ToLowerInvariant()
        foreach ($part in $p.Split('\')) {
            if ($part.StartsWith("[") -and $part.EndsWith("]")) {
                $ck = $part.ToLowerInvariant()
                if (-not $categories.ContainsKey($ck)) { $categories[$ck] = New-Object System.Collections.Generic.List[string] }
                $categories[$ck].Add($nm)
            }
        }
    }
    & $progress (T "p_cfg")
    $cfgState = Parse-Cfgs $cfgFiles $categories
    $haveCfg = $cfgState.Files.Count -gt 0
    $posOf = { param($name) $n = $name.ToLowerInvariant(); if ($cfgState.Pos.ContainsKey($n)) { return [int]$cfgState.Pos[$n] }; return $null }

    # recursos y choques
    $collisions = @{}        # "ext|key" -> lista de @{ Res; File }
    $hashNotes = New-Object System.Collections.Generic.List[string]
    $allOtherKeys = @{}      # "ext|key" -> lista de "recurso: archivo"
    $names = @{}
    foreach ($p in $resPaths) {
        $rl = [IO.Path]::GetFileName($p).ToLowerInvariant()
        if (-not $names.ContainsKey($rl)) { $names[$rl] = New-Object System.Collections.Generic.List[string] }
        $names[$rl].Add($p)
    }
    $i = 0
    foreach ($p in $resPaths) {
        $i++
        if ($i % 10 -eq 0) { & $progress ((T "p_res") -f $i, $resPaths.Count) }
        $rn = [IO.Path]::GetFileName($p)
        # si hay varias carpetas con el mismo nombre de recurso, se muestran con su ruta
        if ($names[$rn.ToLowerInvariant()].Count -gt 1) { $rn = "{0}  ({1})" -f $rn, $p.Substring($resRoot.Length).TrimStart('\') }
        $files = Stream-Files $p
        foreach ($f in $files) {
            $k = "$($f.Ext)|$($f.Key)"
            if (-not $allOtherKeys.ContainsKey($k)) { $allOtherKeys[$k] = New-Object System.Collections.Generic.List[string] }
            $allOtherKeys[$k].Add("$rn : $($f.Rel)")
            if ($mineMap.ContainsKey($k)) {
                if (-not $collisions.ContainsKey($k)) { $collisions[$k] = New-Object System.Collections.Generic.List[object] }
                $collisions[$k].Add([pscustomobject]@{ Res = $rn; ResPath = $p; File = $f; Pos = (& $posOf ([IO.Path]::GetFileName($p))) })
            }
            if ($null -ne $f.Intended -and $f.Intended -ne $f.Key -and $mineMap.ContainsKey("$($f.Ext)|$($f.Intended)")) {
                $mf = $mineMap["$($f.Ext)|$($f.Intended)"][0]
                $hashNotes.Add(((T "h_other") -f $rn, $f.Rel, $mf.Name, (Hex $f.Intended), (Hex $f.Key)))
            }
        }
    }
    foreach ($f in $mine) {
        if ($null -ne $f.Intended -and $f.Intended -ne $f.Key) {
            $hashNotes.Add(((T "h_mine") -f $assetName, $f.Rel, (Hex $f.Intended), (Hex $f.Key)))
            $k2 = "$($f.Ext)|$($f.Intended)"
            if ($allOtherKeys.ContainsKey($k2)) { $hashNotes.Add(((T "h_who") -f (Hex $f.Intended), ($allOtherKeys[$k2] -join "; "))) }
        }
    }

    # ---------------- reporte
    $myPos = & $posOf $assetName
    $now = Get-Date
    & $w ((T "title") -f $Version)
    & $w ((T "date") -f $now.ToString("yyyy-MM-dd HH:mm"), $env:COMPUTERNAME)
    & $w ((T "asset") -f $assetName)
    & $w ((T "folder") -f $assetPath)
    & $w ((T "nstream") -f $mine.Count)
    if ($haveCfg) {
        if ($null -ne $myPos) { & $w ((T "loaded_yes") -f ($myPos + 1), $cfgState.Pos.Count) }
        else { & $w (T "loaded_no") }
    }
    & $w ((T "resfolder") -f $resRoot, $resPaths.Count)
    if ($haveCfg) {
        & $w ((T "cfgread") -f ($cfgState.Files -join " | "))
        if ($cfgState.Missing.Count) { & $w ((T "cfgmiss") -f ($cfgState.Missing -join ", ")) }
        if ($cfgState.CatUsed.Count) { & $w ((T "catused") -f ($cfgState.CatUsed -join ", ")) }
        if ($cfgState.CatUnknown.Count) { & $w ((T "catunk") -f ($cfgState.CatUnknown -join ", ")) }
    } else {
        & $w (T "nocfg")
    }
    & $w ""

    # por recurso (por carpeta: dos recursos con el mismo nombre se cuentan aparte).
    # Los _manifest.ymf van aparte: chocan entre casi todos los MLO y no está confirmado que sea un problema.
    $byRes = @{}
    $md5Cache = @{}
    $getMd5 = { param($path) if (-not $md5Cache.ContainsKey($path)) { $md5Cache[$path] = Md5 $path }; return $md5Cache[$path] }
    foreach ($k in $collisions.Keys) {
        foreach ($o in $collisions[$k]) {
            if (-not $byRes.ContainsKey($o.ResPath)) {
                $byRes[$o.ResPath] = [pscustomobject]@{ Res = $o.Res; Path = $o.ResPath; Pos = $o.Pos; N = 0; Diff = 0; YmfDiff = 0
                                                        Files = New-Object System.Collections.Generic.List[string] }
            }
            $mf = $mineMap[$k][0]
            $same = (& $getMd5 $mf.Path) -eq (& $getMd5 $o.File.Path)
            $o | Add-Member -NotePropertyName Same -NotePropertyValue $same -Force
            $o | Add-Member -NotePropertyName Escrow -NotePropertyValue (Is-Escrow $o.File.Path) -Force
            $r = $byRes[$o.ResPath]
            $r.N++
            if (-not $same) {
                if ($o.File.Ext -eq ".ymf") { $r.YmfDiff++ } else { $r.Diff++; $r.Files.Add($o.File.Name) }
            }
        }
    }
    $loaded = @($byRes.Values | Where-Object { -not $haveCfg -or $null -ne $_.Pos })
    $notLoaded = @($byRes.Values | Where-Object { $haveCfg -and $null -eq $_.Pos })
    $problems = @($loaded | Where-Object { $_.Diff -gt 0 })
    $ymfOnly = @($loaded | Where-Object { $_.Diff -eq 0 -and $_.YmfDiff -gt 0 })

    if ($problems.Count -eq 0 -and $hashNotes.Count -eq 0) {
        & $w (T "ok")
    } else {
        $extra = ""
        if ($hashNotes.Count) { $extra = T "bad_hash" }
        & $w ((T "bad") -f $problems.Count, $extra)
    }
    & $w ""
    if ($problems.Count) {
        & $w (T "need")
        foreach ($r in ($problems | Sort-Object Res)) {
            $dupMark = ""
            if ($names[[IO.Path]::GetFileName($r.Path).ToLowerInvariant()].Count -gt 1) { $dupMark = T "dupmark" }
            & $w ((T "needline") -f $r.Res, $r.Diff, (($r.Files | Select-Object -Unique) -join ", "), $dupMark)
            & $w ("      {0}" -f $r.Path)
        }
        & $w ""
    }
    $sameOnly = @($loaded | Where-Object { $_.Diff -eq 0 -and $_.YmfDiff -eq 0 })
    if ($sameOnly.Count) { & $w ((T "same") -f (($sameOnly | Sort-Object Res | ForEach-Object { $_.Res }) -join ", ")) }
    if ($ymfOnly.Count) { & $w ((T "ymf") -f (($ymfOnly | Sort-Object Res | ForEach-Object { $_.Res }) -join ", ")) }
    if ($notLoaded.Count) { & $w ((T "notloaded") -f (($notLoaded | Sort-Object Res | ForEach-Object { $_.Res }) -join ", ")) }
    & $w ""

    if ($hashNotes.Count) {
        & $w (T "hashhdr")
        & $w (T "hashexp")
        foreach ($n in $hashNotes) { & $w ("  " + $n) }
        & $w ""
    }

    if ($collisions.Count) {
        & $w (T "detail")
        foreach ($k in ($collisions.Keys | Sort-Object { $mineMap[$_][0].Ext -ne ".ymap" }, { $mineMap[$_][0].Name })) {
            $mf = $mineMap[$k][0]
            & $w ((T "dline") -f $mf.Name, $mf.Ext.TrimStart('.'), (Hex $mf.Key))
            $tag = ""
            if ($haveCfg) { if ($null -ne $myPos) { $tag = (T "starts") -f ($myPos + 1) } else { $tag = T "notl" } }
            $enc = ""
            if (Is-Escrow $mf.Path) { $enc = T "enc" }
            $assetLabel = (T "theasset") -f $assetName
            & $w ("    {0}: {1}  {2}{3}{4}" -f $assetLabel, $mf.Rel, (Size-Text (Get-Item -LiteralPath $mf.Path).Length), $tag, $enc)
            $cands = New-Object System.Collections.Generic.List[object]
            if ($null -ne $myPos) { $cands.Add([pscustomobject]@{ Res = $assetLabel; Pos = $myPos }) }
            foreach ($o in ($collisions[$k] | Sort-Object Res)) {
                $tag = ""
                if ($haveCfg) { if ($null -ne $o.Pos) { $tag = (T "starts") -f ($o.Pos + 1); $cands.Add([pscustomobject]@{ Res = $o.Res; Pos = $o.Pos }) } else { $tag = T "notl" } }
                $st = if ($o.Same) { T "identical" } else { T "different" }
                $enc = ""
                if ($o.Escrow) { $enc = T "enc" }
                & $w ("    {0}: {1}  {2}  {3}{4}{5}" -f $o.Res, $o.File.Rel, (Size-Text (Get-Item -LiteralPath $o.File.Path).Length), $st, $tag, $enc)
            }
            if ($haveCfg -and $cands.Count -gt 1) {
                $top = ($cands | Measure-Object -Property Pos -Maximum).Maximum
                $tops = @($cands | Where-Object { $_.Pos -eq $top })
                if ($tops.Count -eq 1) { $winTxt = $tops[0].Res }
                else { $winTxt = (T "dupwin") -f (($tops[0].Res -split '  \(')[0]), $tops.Count }
                & $w ((T "wins") -f $winTxt)
            }
            if ($mf.Ext -eq ".ymf") { & $w (T "ymfnote") }
        }
        & $w ""
    }

    $dupMine = @($mineMap.Keys | Where-Object { $mineMap[$_].Count -gt 1 })
    if ($dupMine.Count) {
        & $w (T "dupmine")
        foreach ($k in $dupMine) { & $w ("  " + (($mineMap[$k] | ForEach-Object { $_.Rel }) -join "  |  ")) }
        & $w ""
    }
    $dupRes = @($names.Keys | Where-Object { $names[$_].Count -gt 1 })
    if ($dupRes.Count) {
        & $w (T "dupres")
        foreach ($n in $dupRes) { & $w ("  {0}: {1}" -f $n, ($names[$n] -join "  |  ")) }
        & $w ""
    }
    & $w ("-" * 70)
    & $w (T "disc_hdr")
    & $w (T "disc")
    & $w ""
    & $w (T "end")
    return [pscustomobject]@{ Text = ($L -join "`r`n"); Problems = $problems.Count; Notes = $hashNotes.Count; AssetName = $assetName }
}

function Save-Report($result, [string]$outPath) {
    if (-not $outPath) {
        $dir = Join-Path $ScriptDir (T "dir")
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
        $safe = ($result.AssetName -replace '[^A-Za-z0-9_.-]', '_')
        $outPath = Join-Path $dir ("compat_{0}_{1}.txt" -f $safe, (Get-Date).ToString("yyyy-MM-dd_HH-mm"))
    }
    [IO.File]::WriteAllText($outPath, $result.Text, (New-Object System.Text.UTF8Encoding($true)))
    return $outPath
}

function Default-Cfgs([string]$resRoot) {
    $d = [IO.Path]::GetFullPath($resRoot)
    for ($i = 0; $i -lt 4; $i++) {
        $p = Join-Path $d "server.cfg"
        if (Test-Path -LiteralPath $p -PathType Leaf) { return @($p) }
        $parent = Split-Path -Parent $d
        if (-not $parent -or $parent -eq $d) { break }
        $d = $parent
    }
    return @()
}

# ------------------------------------------------------------------ sin ventana

if ($NoGui) {
    if (-not $Cfg -or $Cfg.Count -eq 0) { $Cfg = Default-Cfgs $Resources }
    $r = Run-Check $Asset $Resources $Cfg { param($s) Write-Host $s }
    $p = Save-Report $r $Out
    Write-Host ""
    Write-Host $r.Text
    Write-Host ""
    Write-Host ((T "saved") -f $p)
    exit 0
}

# ------------------------------------------------------------------ ventana

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Size = New-Object System.Drawing.Size(720, 420)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)

function Add-Row([int]$y) {
    $l = New-Object System.Windows.Forms.Label
    $l.Location = New-Object System.Drawing.Point(16, $y); $l.Size = New-Object System.Drawing.Size(670, 20)
    $form.Controls.Add($l)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point(16, ($y + 22)); $t.Size = New-Object System.Drawing.Size(560, 24)
    $form.Controls.Add($t)
    $b = New-Object System.Windows.Forms.Button
    $b.Location = New-Object System.Drawing.Point(586, ($y + 20)); $b.Size = New-Object System.Drawing.Size(100, 28)
    $form.Controls.Add($b)
    return @($l, $t, $b)
}

$rowA = Add-Row 44
$rowR = Add-Row 108
$rowC = Add-Row 172
$txtA = $rowA[1]; $txtR = $rowR[1]; $txtC = $rowC[1]
$txtA.Text = $Asset; $txtR.Text = $Resources; $txtC.Text = ($Cfg -join "; ")

$langLabel = New-Object System.Windows.Forms.Label
$langLabel.Location = New-Object System.Drawing.Point(16, 14); $langLabel.Size = New-Object System.Drawing.Size(80, 22)
$form.Controls.Add($langLabel)
$langBox = New-Object System.Windows.Forms.ComboBox
$langBox.DropDownStyle = "DropDownList"
$langBox.Location = New-Object System.Drawing.Point(100, 11); $langBox.Size = New-Object System.Drawing.Size(130, 24)
[void]$langBox.Items.Add("Español")
[void]$langBox.Items.Add("English")
$form.Controls.Add($langBox)

$status = New-Object System.Windows.Forms.Label
$status.Location = New-Object System.Drawing.Point(16, 282); $status.Size = New-Object System.Drawing.Size(670, 40)
$form.Controls.Add($status)

$disc = New-Object System.Windows.Forms.Label
$disc.Location = New-Object System.Drawing.Point(16, 330); $disc.Size = New-Object System.Drawing.Size(670, 36)
$disc.ForeColor = [System.Drawing.Color]::DimGray
$disc.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$form.Controls.Add($disc)

$go = New-Object System.Windows.Forms.Button
$go.Location = New-Object System.Drawing.Point(16, 238); $go.Size = New-Object System.Drawing.Size(180, 36)
$go.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($go)

function Apply-Texts {
    $form.Text = (T "g_title") -f $Version
    $langLabel.Text = T "g_lang"
    $rowA[0].Text = T "g_a"; $rowR[0].Text = T "g_r"; $rowC[0].Text = T "g_c"
    $rowA[2].Text = T "g_pick"; $rowR[2].Text = T "g_pick"; $rowC[2].Text = T "g_pick"
    $go.Text = T "g_go"
    $status.Text = T "g_status"
    $disc.Text = T "g_disc"
}
$langBox.SelectedIndex = $(if ($script:Lang -eq "en") { 1 } else { 0 })
$langBox.Add_SelectedIndexChanged({
    $script:Lang = $(if ($langBox.SelectedIndex -eq 1) { "en" } else { "es" })
    Apply-Texts
})
Apply-Texts

function Pick-Folder([string]$desc, [string]$start) {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = $desc
    if ($start -and (Test-Path -LiteralPath $start)) { $dlg.SelectedPath = $start }
    if ($dlg.ShowDialog($form) -eq "OK") { return $dlg.SelectedPath }
    return $null
}
$rowA[2].Add_Click({ $p = Pick-Folder (T "d_asset") $txtA.Text; if ($p) { $txtA.Text = $p } })
$rowR[2].Add_Click({
    $p = Pick-Folder (T "d_res") $txtR.Text
    if ($p) {
        $txtR.Text = $p
        if (-not $txtC.Text) { $txtC.Text = ((Default-Cfgs $p) -join "; ") }
    }
})
$rowC[2].Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = T "d_filter"
    $dlg.Multiselect = $true
    $dlg.Title = T "d_cfg"
    $first = ($txtC.Text -split ";")[0].Trim()
    if ($first -and (Test-Path -LiteralPath $first)) { $dlg.InitialDirectory = Split-Path -Parent $first }
    elseif ($txtR.Text -and (Test-Path -LiteralPath $txtR.Text)) { $dlg.InitialDirectory = $txtR.Text }
    if ($dlg.ShowDialog($form) -eq "OK") {
        # server.cfg primero
        $sel = @($dlg.FileNames | Sort-Object { [IO.Path]::GetFileName($_) -ne "server.cfg" })
        $txtC.Text = ($sel -join "; ")
    }
})

$go.Add_Click({
    $a = $txtA.Text.Trim(); $rr = $txtR.Text.Trim()
    if (-not (Test-Path -LiteralPath $a -PathType Container)) { [System.Windows.Forms.MessageBox]::Show((T "m_noasset")) | Out-Null; return }
    if (-not (Test-Path -LiteralPath $rr -PathType Container)) { [System.Windows.Forms.MessageBox]::Show((T "m_nores")) | Out-Null; return }
    $cfgs = @($txtC.Text -split ";" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $go.Enabled = $false
    $langBox.Enabled = $false
    $form.Cursor = "WaitCursor"
    try {
        $r = Run-Check $a $rr $cfgs { param($s) $status.Text = $s; [System.Windows.Forms.Application]::DoEvents() }
        $p = Save-Report $r ""
        $status.Text = (T "status_done") -f $p
        $msg = $(if ($r.Problems -gt 0 -or $r.Notes -gt 0) { T "m_found" } else { T "m_none" })
        $nl = [Environment]::NewLine
        [System.Windows.Forms.MessageBox]::Show(($msg + $nl + $nl + (T "m_where") + $nl + $p + $nl + $nl + (T "m_send")), (T "m_ready")) | Out-Null
        Start-Process explorer.exe "/select,`"$p`""
    } catch {
        $status.Text = (T "m_err") -f $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show(((T "m_err") -f $_.Exception.Message)) | Out-Null
    } finally {
        $go.Enabled = $true
        $langBox.Enabled = $true
        $form.Cursor = "Default"
    }
})

[void]$form.ShowDialog()
