#Requires -Version 5.1
<#
.SYNOPSIS
    Limpieza segura y configurable de temporales, caches, Windows.old y la Papelera.

.DESCRIPTION
    Inventaria y, solo si se pide, elimina contenido desechable:

      - archivos temporales del usuario y de Windows
      - caches de Windows (INetCache, miniaturas, WER, Delivery Optimization,
        descarga de Windows Update)
      - caches de navegadores (Edge, Chrome, Firefox, Brave, Opera), sin tocar
        contrasenas, marcadores ni historial
      - caches de NuGet (HTTP, v3, plugins y paquetes globales)
      - Windows.old
      - la Papelera de reciclaje

    El modo predeterminado es dry-run: mide e informa y no borra nada.
    Cada carpeta se inventaria en una sola pasada nativa de solo lectura.
    No se lanza un segundo recorrido para los archivos mas grandes.
    -Apply arma la eliminacion. Sin -Force, en una consola interactiva hay que
    escribir SI. -Force es la confirmacion explicita para el Programador de
    tareas. La whitelist se vuelve a comprobar antes de cada borrado.

    No se sigue un punto de union. No se borra la carpeta raiz de un objetivo
    salvo Windows.old, y solo si la ruta es exactamente X:\Windows.old.

.PARAMETER TempFiles
    Incluye carpetas temporales del usuario y C:\Windows\Temp.

.PARAMETER WindowsCache
    Incluye caches de Windows descritos en la ayuda.

.PARAMETER BrowserCache
    Incluye solo subcarpetas de cache de los navegadores detectados.

.PARAMETER NuGetCache
    Incluye los caches locales de NuGet, paquetes globales incluidos.

.PARAMETER WindowsOld
    Incluye X:\Windows.old si existe.

.PARAMETER RecycleBin
    Incluye la Papelera de reciclaje.

.PARAMETER All
    Incluye todas las categorias. Tambien es el comportamiento si no se elige
    ninguna.

.PARAMETER Apply
    Ejecuta el borrado. Sin este modificador el script solo informa.

.PARAMETER Force
    Omite la pregunta SI. No omite la whitelist ni el dry-run: hace falta
    -Apply. Pensado para una tarea no interactiva.

.PARAMETER OlderThanDays
    Solo cuenta, y solo borraria, archivos mas viejos que este numero de dias.
    Predeterminado: 7. 0 incluye todos los archivos del objetivo.
    No aplica a la Papelera. La pasada sigue visitando la carpeta para conocer
    la fecha de cada archivo; no borra nada en dry-run.

.PARAMETER Top
    Cantidad de archivos mas grandes que se listan. De 1 a 100. Predeterminado 15.

.PARAMETER Whitelist
    Rutas adicionales que no se pueden borrar, ni ellas ni su contenido.
    Se suman a la whitelist integrada.

.PARAMETER LogPath
    JSONL de auditoria. Predeterminado: logs\SmartCleanup.jsonl.
    Se escribe tambien en dry-run.

.PARAMETER ReportPath
    JSON de auditoria. Solo se escribe junto con -Apply. En dry-run el JSON
    sale por la canalizacion y no se crea el archivo.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\SmartCleanup.ps1

    Dry-run de todas las categorias.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\SmartCleanup.ps1 -TempFiles -BrowserCache

    Dry-run solo de temporales y caches de navegador.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\SmartCleanup.ps1 -NuGetCache -Apply

    Borra el cache de NuGet despues de escribir SI.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\SmartCleanup.ps1 -TempFiles -OlderThanDays 7 -Apply -Force

    Tarea programada: temporales con mas de 7 dias, sin pregunta interactiva.

.NOTES
    Autor: LocalAgents
    Requiere: Windows PowerShell 5.1
    La medicion no usa robocopy ni consulta el catalogo de Windows Update.
    Codigos de salida (Common.psm1):
      0  inventario completo, sin rechazos ni lecturas incompletas
      1  termino con objetivos rechazados o mediciones incompletas
      2  reservado; esta rutina no marca Critical
      3  el script fallo antes de cerrar el reporte
    Un dry-run grande no es Critical: solo informa.
    Helpers compartidos: scripts\Common.psm1.
    No eleva privilegios. Las rutas sin permiso quedan en el informe.

    Programador de tareas (temporales de mas de 7 dias, a las 07:30):
      $script = 'C:\ProyectosDeSoftware\LocalAgents\scripts\SmartCleanup.ps1'
      $arg = "-NoProfile -ExecutionPolicy Bypass -File `"$script`" -TempFiles -OlderThanDays 7 -Apply -Force"
      $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arg
      $trigger = New-ScheduledTaskTrigger -Daily -At 7:30am
      Register-ScheduledTask -TaskName 'LocalAgents-SmartCleanup' -Action $action -Trigger $trigger -Description 'Limpieza segura de temporales'
#>
[CmdletBinding()]
param(
    [switch]$TempFiles,
    [switch]$WindowsCache,
    [switch]$BrowserCache,
    [switch]$NuGetCache,
    [switch]$WindowsOld,
    [switch]$RecycleBin,
    [switch]$All,

    [switch]$Apply,
    [switch]$Force,

    [ValidateRange(0, 3650)]
    [int]$OlderThanDays = 7,

    [ValidateRange(1, 100)]
    [int]$Top = 15,

    [string[]]$Whitelist,

    [string]$LogPath,
    [string]$ReportPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -Force -DisableNameChecking

$script:Apply = [bool]$Apply
Set-DryRunMode -DryRun (-not [bool]$Apply)
$script:ExitCode = 0
$script:Whitelist = $null
$script:RobocopyPath = $null
$script:DeletedFiles = 0
$script:DeletedBytes = [int64]0


function ConvertTo-Int64Loose {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $digits = ($Text -replace '[^\d]', '')
    if ([string]::IsNullOrWhiteSpace($digits)) { return $null }
    return [int64]$digits
}

function Get-NormalizedPath {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $full = [System.IO.Path]::GetFullPath($Path)
    if ($full.Length -gt 3) { $full = $full.TrimEnd('\') }
    return $full
}

function Test-IsDriveRoot {
    param([string]$Path)
    return ($Path -match '^[A-Za-z]:\\$')
}

function Test-IsUnderPath {
    param(
        [string]$Path,
        [string]$Root,
        [switch]$AllowEqual
    )
    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($Root)) { return $false }
    $left = $Path.TrimEnd('\')
    $right = $Root.TrimEnd('\')
    if ($left.Equals($right, [System.StringComparison]::OrdinalIgnoreCase)) {
        return [bool]$AllowEqual
    }
    $prefix = $right + '\'
    return $left.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-ProtectedPath {
    param([Parameter(Mandatory = $true)][string]$Path)
    foreach ($entry in $script:Whitelist) {
        if (Test-IsUnderPath -Path $Path -Root $entry -AllowEqual) { return $true }
    }
    return $false
}

function Test-ForbiddenRoot {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (Test-IsDriveRoot -Path $Path) { return $true }
    $systemDrive = $env:SystemDrive
    if ([string]::IsNullOrWhiteSpace($systemDrive)) { $systemDrive = 'C:' }
    $forbidden = @(
        (Get-NormalizedPath -Path ($systemDrive + '\Windows')),
        (Get-NormalizedPath -Path ($systemDrive + '\Users')),
        (Get-NormalizedPath -Path $env:USERPROFILE),
        (Get-NormalizedPath -Path ($systemDrive + '\Program Files')),
        (Get-NormalizedPath -Path ($systemDrive + '\Program Files (x86)')),
        (Get-NormalizedPath -Path ($systemDrive + '\ProgramData')),
        (Get-NormalizedPath -Path ($systemDrive + '\Windows\System32')),
        (Get-NormalizedPath -Path ($systemDrive + '\Windows\SysWOW64'))
    )
    foreach ($item in $forbidden) {
        if ([string]::IsNullOrWhiteSpace($item)) { continue }
        if ($Path.Equals($item, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Test-AllowRootSafe {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Reason
    )
    $normalized = Get-NormalizedPath -Path $Path
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return ,([ordered]@{ ok = $false; path = $Path; reason = 'Ruta vacia.' })
    }
    if (Test-IsDriveRoot -Path $normalized) {
        return ,([ordered]@{ ok = $false; path = $normalized; reason = 'Una unidad completa no es un objetivo valido.' })
    }
    if (Test-ForbiddenRoot -Path $normalized) {
        return ,([ordered]@{ ok = $false; path = $normalized; reason = 'La ruta es una carpeta critica del sistema o del perfil.' })
    }
    if (Test-ProtectedPath -Path $normalized) {
        return ,([ordered]@{ ok = $false; path = $normalized; reason = 'La ruta esta dentro de la whitelist.' })
    }
    try {
        if (Test-Path -LiteralPath $normalized) {
            $attributes = [System.IO.File]::GetAttributes($normalized)
            if (($attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                return ,([ordered]@{ ok = $false; path = $normalized; reason = 'La ruta es un punto de union y no se sigue.' })
            }
        }
    }
    catch {
        return ,([ordered]@{ ok = $false; path = $normalized; reason = $_.Exception.Message })
    }
    return ,([ordered]@{ ok = $true; path = $normalized; reason = $Reason })
}

function Test-ApprovedForDeletion {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$AllowRoot,
        [switch]$AllowRootItself
    )
    $normalized = Get-NormalizedPath -Path $Path
    $root = Get-NormalizedPath -Path $AllowRoot
    if ([string]::IsNullOrWhiteSpace($normalized) -or [string]::IsNullOrWhiteSpace($root)) { return $false }
    if (Test-IsDriveRoot -Path $normalized) { return $false }
    if (Test-ForbiddenRoot -Path $normalized) { return $false }
    if (Test-ProtectedPath -Path $normalized) { return $false }
    if (Test-ProtectedPath -Path $root) { return $false }
    if ($normalized.Equals($root, [System.StringComparison]::OrdinalIgnoreCase)) {
        return [bool]$AllowRootItself
    }
    if (-not (Test-IsUnderPath -Path $normalized -Root $root)) { return $false }
    return $true
}


function Get-RobocopySource {
    param([Parameter(Mandatory = $true)][string]$Path)
    if ($Path.StartsWith('\\?\')) { return $Path }
    if ($Path.Length -lt 240) { return $Path }
    if ($Path.StartsWith('\\')) { return ('\\?\UNC\' + $Path.Substring(2)) }
    return ('\\?\' + $Path)
}

function Invoke-RobocopyMeasure {
    param([Parameter(Mandatory = $true)][string]$Path)

    $robocopyArgs = @(
        (Get-RobocopySource -Path $Path),
        'NULL',
        '/L', '/S', '/BYTES', '/XJ', '/R:0', '/W:0', '/NJH', '/NDL', '/NC', '/NFL', '/NP', '/MT:8'
    )
    if ($robocopyArgs -notcontains '/L') {
        throw 'Invariante rota: robocopy debe ejecutarse siempre con /L.'
    }

    $output = & $script:RobocopyPath @robocopyArgs 2>&1
    $exitCode = $LASTEXITCODE
    $text = (($output | ForEach-Object { "$_" }) -join "`n")
    $bytes = [int64]0
    $files = [int64]0
    $directoryErrors = [int64]0
    $parsedBytes = $false

    foreach ($line in ($text -split '\r?\n')) {
        if ($line -match '^\s*Bytes\s*:\s+(\S+)') {
            $parsed = ConvertTo-Int64Loose -Text $Matches[1]
            if ($null -ne $parsed) { $bytes = $parsed; $parsedBytes = $true }
            continue
        }
        if ($line -match '^\s*(?:Archivos|Files)\s*:\s+(\S+)') {
            $parsedFiles = ConvertTo-Int64Loose -Text $Matches[1]
            if ($null -ne $parsedFiles) { $files = $parsedFiles }
            continue
        }
        if ($line -match '^\s*(?:Director\.|Dirs)\s*:') {
            $parts = $line.Split(':', 2)
            if ($parts.Count -ge 2) {
                $columns = @($parts[1].Trim() -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                if ($columns.Count -ge 5) {
                    $parsedErrors = ConvertTo-Int64Loose -Text $columns[4]
                    if ($null -ne $parsedErrors) { $directoryErrors = $parsedErrors }
                }
            }
        }
    }

    $hardFailure = ($exitCode -ge 16) -or (-not $parsedBytes)
    $errorText = $null
    if ($hardFailure) {
        $snippet = ($text -replace '\s+', ' ').Trim()
        if ($snippet.Length -gt 400) { $snippet = $snippet.Substring(0, 400) }
        $errorText = "robocopy termino con codigo $exitCode. $snippet".Trim()
    }

    return ,([ordered]@{
        ok             = (-not $hardFailure)
        sizeBytes      = $(if ($hardFailure) { $null } else { $bytes })
        fileCount      = $(if ($hardFailure) { $null } else { $files })
        accessErrors   = $directoryErrors
        sizeIsComplete = (-not $hardFailure -and $directoryErrors -eq 0 -and $exitCode -lt 8)
        exitCode       = $exitCode
        error          = $errorText
    })
}

function Add-TopCandidate {
    param(
        [Parameter(Mandatory = $true)]$List,
        [Parameter(Mandatory = $true)]$Candidate,
        [Parameter(Mandatory = $true)][int]$Limit
    )
    if ($List.Count -lt $Limit) {
        $List.Add($Candidate)
        return
    }
    $minIndex = 0
    $minSize = [int64]$List[0].sizeBytes
    for ($index = 1; $index -lt $List.Count; $index++) {
        $size = [int64]$List[$index].sizeBytes
        if ($size -lt $minSize) {
            $minSize = $size
            $minIndex = $index
        }
    }
    if ([int64]$Candidate.sizeBytes -gt $minSize) {
        $List[$minIndex] = $Candidate
    }
}

function Get-CutoffUtc {
    if ($OlderThanDays -le 0) { return $null }
    return [DateTime]::UtcNow.AddDays(-1 * $OlderThanDays)
}

function Test-FileOldEnough {
    param([datetime]$LastWriteTimeUtc, $CutoffUtc)
    if ($null -eq $CutoffUtc) { return $true }
    return ($LastWriteTimeUtc -lt $CutoffUtc)
}

function Get-TopFiles {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][int]$Limit,
        [string]$Pattern
    )

    $best = New-Object System.Collections.Generic.List[object]
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($Root)
    $enumerated = 0
    $matched = 0
    $accessErrors = 0
    $protectedSkipped = 0
    $bytes = [int64]0
    $cutoff = Get-CutoffUtc
    $rootOnly = -not [string]::IsNullOrWhiteSpace($Pattern)

    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        try {
            $directory = New-Object System.IO.DirectoryInfo $current
            $files = $directory.EnumerateFiles()
            if ($rootOnly) { $files = $directory.EnumerateFiles($Pattern) }
            foreach ($file in $files) {
                try {
                    $enumerated++
                    if (($file.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                    if (-not (Test-FileOldEnough -LastWriteTimeUtc $file.LastWriteTimeUtc -CutoffUtc $cutoff)) { continue }
                    if (-not (Test-ApprovedForDeletion -Path $file.FullName -AllowRoot $Root)) {
                        $protectedSkipped++
                        continue
                    }
                    $matched++
                    $length = [int64]$file.Length
                    $bytes += $length
                    Add-TopCandidate -List $best -Limit $Limit -Candidate ([ordered]@{
                        path             = [string]$file.FullName
                        sizeBytes        = $length
                        lastWriteTimeUtc = ($file.LastWriteTimeUtc.ToString('yyyy-MM-ddTHH:mm:ss.fff', [System.Globalization.CultureInfo]::InvariantCulture) + 'Z')
                    })
                }
                catch {
                    $accessErrors++
                }
            }
            if ($rootOnly) { continue }
            foreach ($child in $directory.EnumerateDirectories()) {
                try {
                    if (($child.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                    if (Test-ProtectedPath -Path $child.FullName) {
                        $protectedSkipped++
                        continue
                    }
                    $stack.Push($child.FullName)
                }
                catch {
                    $accessErrors++
                }
            }
        }
        catch {
            $accessErrors++
        }
    }

    $ordered = New-Object System.Collections.Generic.List[object]
    foreach ($item in ($best | Sort-Object { $_.sizeBytes } -Descending)) {
        $ordered.Add($item)
    }

    return ,([ordered]@{
        enumerated       = $enumerated
        matched          = $matched
        sizeBytes        = $bytes
        accessErrors     = $accessErrors
        protectedSkipped = $protectedSkipped
        largest          = $ordered
    })
}

function Test-WhitelistInside {
    param([Parameter(Mandatory = $true)][string]$Root)
    foreach ($entry in $script:Whitelist) {
        if (Test-IsUnderPath -Path $entry -Root $Root) { return $true }
    }
    return $false
}

function New-Target {
    param(
        [Parameter(Mandatory = $true)][string]$Category,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Path,
        [ValidateSet('DirectoryContents', 'MatchingFiles', 'RemoveTree')]
        [string]$Mode = 'DirectoryContents',
        [string]$Pattern
    )
    return [ordered]@{
        category = $Category
        name     = $Name
        path     = $Path
        mode     = $Mode
        pattern  = $Pattern
    }
}

function Add-TargetIfSafe {
    param(
        $Bucket,
        [string]$Category,
        [string]$Name,
        [string]$Path,
        [string]$Mode = 'DirectoryContents',
        [string]$Pattern
    )
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    $decision = Test-AllowRootSafe -Path $Path
    if (-not $decision.ok) {
        $script:Rejected.Add([ordered]@{
            category = $Category
            name     = $Name
            path     = [string]$decision.path
            reason   = [string]$decision.reason
        })
        Write-Log -Level WARN -Event 'target_rejected' -Data @{
            category = $Category
            name     = $Name
            path     = [string]$decision.path
            reason   = [string]$decision.reason
        }
        return
    }
    $Bucket.Add((New-Target -Category $Category -Name $Name -Path ([string]$decision.path) -Mode $Mode -Pattern $Pattern))
}

function Get-SelectedCategoryNames {
    $names = New-Object System.Collections.Generic.List[string]
    $explicit = $TempFiles -or $WindowsCache -or $BrowserCache -or $NuGetCache -or $WindowsOld -or $RecycleBin
    $useAll = $All -or (-not $explicit)
    if ($useAll -or $TempFiles) { $names.Add('TempFiles') }
    if ($useAll -or $WindowsCache) { $names.Add('WindowsCache') }
    if ($useAll -or $BrowserCache) { $names.Add('BrowserCache') }
    if ($useAll -or $NuGetCache) { $names.Add('NuGetCache') }
    if ($useAll -or $WindowsOld) { $names.Add('WindowsOld') }
    if ($useAll -or $RecycleBin) { $names.Add('RecycleBin') }
    Write-Output -NoEnumerate $names.ToArray()
}

function Add-BrowserProfileCaches {
    param(
        $Bucket,
        [string]$Browser,
        [string]$UserData,
        [string[]]$RelativeCaches
    )
    if ([string]::IsNullOrWhiteSpace($UserData)) { return }
    if (-not (Test-Path -LiteralPath $UserData -PathType Container)) { return }
    $profiles = @(Get-ChildItem -LiteralPath $UserData -Force -Directory -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -eq 'Default' -or $_.Name -like 'Profile *' -or $_.Name -eq 'Guest Profile'
    })
    foreach ($profile in $profiles) {
        foreach ($relative in $RelativeCaches) {
            $cachePath = Join-Path $profile.FullName $relative
            if (Test-Path -LiteralPath $cachePath -PathType Container) {
                Add-TargetIfSafe -Bucket $Bucket -Category 'BrowserCache' -Name ("{0} {1} {2}" -f $Browser, $profile.Name, $relative) -Path $cachePath
            }
        }
    }
    foreach ($rootCache in @('ShaderCache', 'GrShaderCache', 'GPUCache')) {
        $cachePath = Join-Path $UserData $rootCache
        if (Test-Path -LiteralPath $cachePath -PathType Container) {
            Add-TargetIfSafe -Bucket $Bucket -Category 'BrowserCache' -Name ("{0} {1}" -f $Browser, $rootCache) -Path $cachePath
        }
    }
}

function Get-CleanupTargets {
    param([string[]]$Categories)

    $bucket = New-Object System.Collections.Generic.List[object]
    $systemDrive = $env:SystemDrive
    if ([string]::IsNullOrWhiteSpace($systemDrive)) { $systemDrive = 'C:' }
    $windows = Join-Path ($systemDrive + '\') 'Windows'

    foreach ($category in $Categories) {
        switch ($category) {
            'TempFiles' {
                $tempPaths = New-Object System.Collections.Generic.List[string]
                foreach ($candidate in @($env:TEMP, (Join-Path $env:LOCALAPPDATA 'Temp'), (Join-Path $windows 'Temp'))) {
                    if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
                    $full = Get-NormalizedPath -Path $candidate
                    if ([string]::IsNullOrWhiteSpace($full)) { continue }
                    $key = $full.ToUpperInvariant()
                    if ($tempPaths -contains $key) { continue }
                    $tempPaths.Add($key)
                    Add-TargetIfSafe -Bucket $bucket -Category 'TempFiles' -Name $full -Path $full
                }
            }
            'WindowsCache' {
                $windowsTargets = @(
                    @{ Name = 'INetCache'; Path = (Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\INetCache') },
                    @{ Name = 'WER ReportArchive'; Path = (Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\WER\ReportArchive') },
                    @{ Name = 'WER ReportQueue'; Path = (Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\WER\ReportQueue') },
                    @{ Name = 'WER ProgramData ReportArchive'; Path = (Join-Path $env:ProgramData 'Microsoft\Windows\WER\ReportArchive') },
                    @{ Name = 'WER ProgramData ReportQueue'; Path = (Join-Path $env:ProgramData 'Microsoft\Windows\WER\ReportQueue') },
                    @{ Name = 'Windows Update Download'; Path = (Join-Path $windows 'SoftwareDistribution\Download') },
                    @{ Name = 'Delivery Optimization'; Path = (Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\DeliveryOptimization\Cache') }
                )
                foreach ($item in $windowsTargets) {
                    Add-TargetIfSafe -Bucket $bucket -Category 'WindowsCache' -Name $item.Name -Path $item.Path
                }
                $explorer = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Explorer'
                if (Test-Path -LiteralPath $explorer -PathType Container) {
                    Add-TargetIfSafe -Bucket $bucket -Category 'WindowsCache' -Name 'Miniaturas thumbcache' -Path $explorer -Mode 'MatchingFiles' -Pattern 'thumbcache_*.db'
                    Add-TargetIfSafe -Bucket $bucket -Category 'WindowsCache' -Name 'Cache de iconos' -Path $explorer -Mode 'MatchingFiles' -Pattern 'iconcache_*.db'
                }
            }
            'BrowserCache' {
                $relative = @('Cache', 'Code Cache', 'GPUCache', 'GrShaderCache', 'ShaderCache', 'Service Worker\CacheStorage')
                Add-BrowserProfileCaches -Bucket $bucket -Browser 'Edge' -UserData (Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\User Data') -RelativeCaches $relative
                Add-BrowserProfileCaches -Bucket $bucket -Browser 'Chrome' -UserData (Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data') -RelativeCaches $relative
                Add-BrowserProfileCaches -Bucket $bucket -Browser 'Brave' -UserData (Join-Path $env:LOCALAPPDATA 'BraveSoftware\Brave-Browser\User Data') -RelativeCaches $relative
                $opera = Join-Path $env:APPDATA 'Opera Software\Opera Stable\Cache'
                if (Test-Path -LiteralPath $opera -PathType Container) {
                    Add-TargetIfSafe -Bucket $bucket -Category 'BrowserCache' -Name 'Opera Cache' -Path $opera
                }
                $firefoxRoot = Join-Path $env:APPDATA 'Mozilla\Firefox\Profiles'
                $firefoxLocal = Join-Path $env:LOCALAPPDATA 'Mozilla\Firefox\Profiles'
                foreach ($root in @($firefoxRoot, $firefoxLocal)) {
                    if (-not (Test-Path -LiteralPath $root -PathType Container)) { continue }
                    $profiles = @(Get-ChildItem -LiteralPath $root -Force -Directory -ErrorAction SilentlyContinue)
                    foreach ($profile in $profiles) {
                        foreach ($relativeCache in @('cache2', 'startupCache', 'jumpListCache')) {
                            $cachePath = Join-Path $profile.FullName $relativeCache
                            if (Test-Path -LiteralPath $cachePath -PathType Container) {
                                Add-TargetIfSafe -Bucket $bucket -Category 'BrowserCache' -Name ("Firefox {0} {1}" -f $profile.Name, $relativeCache) -Path $cachePath
                            }
                        }
                    }
                }
            }
            'NuGetCache' {
                $nugetTargets = @(
                    @{ Name = 'NuGet http-cache'; Path = (Join-Path $env:LOCALAPPDATA 'NuGet\http-cache') },
                    @{ Name = 'NuGet v3-cache'; Path = (Join-Path $env:LOCALAPPDATA 'NuGet\v3-cache') },
                    @{ Name = 'NuGet plugins-cache'; Path = (Join-Path $env:LOCALAPPDATA 'NuGet\plugins-cache') },
                    @{ Name = 'NuGet paquetes globales'; Path = (Join-Path $env:USERPROFILE '.nuget\packages') }
                )
                foreach ($item in $nugetTargets) {
                    Add-TargetIfSafe -Bucket $bucket -Category 'NuGetCache' -Name $item.Name -Path $item.Path
                }
            }
            'WindowsOld' {
                $old = Join-Path ($systemDrive + '\') 'Windows.old'
                Add-TargetIfSafe -Bucket $bucket -Category 'WindowsOld' -Name 'Windows.old' -Path $old -Mode 'RemoveTree'
            }
            'RecycleBin' {
                # La Papelera no es una carpeta que se recorra con robocopy.
            }
        }
    }

    Write-Output -NoEnumerate $bucket.ToArray()
}

function Measure-Target {
    param($Target)

    $path = [string]$Target.path
    $record = [ordered]@{
        category         = [string]$Target.category
        name             = [string]$Target.name
        path             = $path
        mode             = [string]$Target.mode
        exists           = $false
        included         = $false
        sizeBytes        = [int64]0
        fileCount        = [int64]0
        sizeIsComplete   = $true
        accessErrors     = 0
        protectedSkipped = 0
        error            = $null
        largestFiles     = (New-Object System.Collections.Generic.List[object])
        durationMs       = [int64]0
    }

    $rejected = $false
    foreach ($item in $script:Rejected) {
        if ([string]$item.path -eq $path -and [string]$item.category -eq [string]$Target.category -and [string]$item.name -eq [string]$Target.name) {
            $rejected = $true
            $record.error = [string]$item.reason
            $record.sizeIsComplete = $false
        }
    }
    if ($rejected) { return ,$record }

    if (-not (Test-Path -LiteralPath $path)) {
        $record.error = 'La ruta no existe.'
        return ,$record
    }
    $record.exists = $true
    $record.included = $true

    $pattern = [string]$Target.pattern
    $rootOnly = ($Target.mode -eq 'MatchingFiles') -or (-not [string]::IsNullOrWhiteSpace($pattern))
    $cutoffTicks = [int64]-1
    $cutoff = Get-CutoffUtc
    if ($null -ne $cutoff) { $cutoffTicks = [int64]$cutoff.Ticks }
    $protected = @()
    if ($null -ne $script:Whitelist) { $protected = @($script:Whitelist) }
    $walk = Measure-TreeFiles -Root $path -Pattern $pattern -CutoffUtcTicks $cutoffTicks -Top $Top -ProtectedRoots $protected -RootOnly:$rootOnly
    $record.sizeBytes = [int64]$walk.sizeBytes
    $record.fileCount = [int64]$walk.fileCount
    $record.accessErrors = [int]$walk.accessErrors
    $record.protectedSkipped = [int]$walk.protectedSkipped
    $record.sizeIsComplete = ([int]$walk.accessErrors -eq 0)
    $record.largestFiles = $walk.largest
    $record.durationMs = [int64]$walk.durationMs
    if ([int]$walk.accessErrors -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$walk.error)) {
        $record.error = [string]$walk.error
    }

    return ,$record
}

function Get-RecycleBinReport {
    param([int]$Limit)

    $files = New-Object System.Collections.Generic.List[object]
    $total = [int64]0
    $count = 0
    $errorText = $null
    try {
        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.NameSpace(10)
        if ($null -eq $folder) { throw 'No se pudo abrir la Papelera.' }
        foreach ($item in $folder.Items()) {
            $count++
            $size = [int64]0
            try {
                $raw = $item.ExtendedProperty('System.Size')
                if ($null -ne $raw -and $raw -ne '') { $size = [int64]$raw }
            }
            catch { }
            $total += $size
            Add-TopCandidate -List $files -Limit $Limit -Candidate ([ordered]@{
                path      = [string]$item.Path
                name      = [string]$item.Name
                sizeBytes = $size
            })
        }
    }
    catch {
        $errorText = $_.Exception.Message
    }

    $ordered = New-Object System.Collections.Generic.List[object]
    foreach ($item in ($files | Sort-Object { $_.sizeBytes } -Descending)) {
        $ordered.Add($item)
    }

    return ,([ordered]@{
        included   = $true
        itemCount  = $count
        sizeBytes  = $total
        error      = $errorText
        largest    = $ordered
    })
}

function Remove-ApprovedChildren {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [string]$Pattern,
        [switch]$RemoveRootWhenEmpty
    )

    if ((Test-DryRun)) {
        throw 'Invariante rota: Remove-ApprovedChildren no puede ejecutarse en dry-run.'
    }
    if (-not (Test-AllowRootSafe -Path $Root).ok) {
        throw "La raiz dejo de ser segura: $Root"
    }

    $directories = New-Object System.Collections.Generic.List[string]
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($Root)
    $rootOnly = -not [string]::IsNullOrWhiteSpace($Pattern)
    $cutoff = Get-CutoffUtc
    $deletedHere = 0
    $failedHere = 0

    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        try {
            $directory = New-Object System.IO.DirectoryInfo $current
            $files = $directory.EnumerateFiles()
            if ($rootOnly) { $files = $directory.EnumerateFiles($Pattern) }
            foreach ($file in $files) {
                try {
                    if (($file.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                    if (-not (Test-FileOldEnough -LastWriteTimeUtc $file.LastWriteTimeUtc -CutoffUtc $cutoff)) { continue }
                    if (-not (Test-ApprovedForDeletion -Path $file.FullName -AllowRoot $Root)) {
                        Write-Log -Level WARN -Event 'delete_blocked' -Data @{ path = $file.FullName; reason = 'whitelist' }
                        continue
                    }
                    $length = [int64]$file.Length
                    [System.IO.File]::SetAttributes($file.FullName, [System.IO.FileAttributes]::Normal)
                    [System.IO.File]::Delete($file.FullName)
                    $script:DeletedFiles++
                    $script:DeletedBytes += $length
                    $deletedHere++
                }
                catch {
                    $failedHere++
                    Write-Log -Level WARN -Event 'delete_failed' -Data @{
                        path    = $file.FullName
                        message = $_.Exception.Message
                    }
                }
            }
            if (-not $rootOnly) {
                foreach ($child in $directory.EnumerateDirectories()) {
                    if (($child.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                    if (-not (Test-ApprovedForDeletion -Path $child.FullName -AllowRoot $Root)) { continue }
                    $stack.Push($child.FullName)
                    $directories.Add($child.FullName)
                }
            }
        }
        catch {
            $failedHere++
            Write-Log -Level WARN -Event 'delete_failed' -Data @{
                path    = $current
                message = $_.Exception.Message
            }
        }
    }

    if (-not $rootOnly) {
        $sortedDirs = $directories.ToArray() | Sort-Object { $_.Length } -Descending
        foreach ($dirPath in $sortedDirs) {
            if (-not (Test-ApprovedForDeletion -Path $dirPath -AllowRoot $Root)) { continue }
            try {
                $remaining = @(Get-ChildItem -LiteralPath $dirPath -Force -ErrorAction SilentlyContinue)
                if ($remaining.Count -eq 0) {
                    [System.IO.Directory]::Delete($dirPath, $false)
                }
            }
            catch {
                Write-Log -Level WARN -Event 'delete_failed' -Data @{
                    path    = [string]$dirPath
                    message = $_.Exception.Message
                }
            }
        }
    }

    if ($RemoveRootWhenEmpty) {
        $windowsOld = Get-NormalizedPath -Path (Join-Path ($env:SystemDrive + '\') 'Windows.old')
        $rootNorm = Get-NormalizedPath -Path $Root
        if ($rootNorm.Equals($windowsOld, [System.StringComparison]::OrdinalIgnoreCase)) {
            try {
                $remainingRoot = @(Get-ChildItem -LiteralPath $rootNorm -Force -ErrorAction SilentlyContinue)
                if ($remainingRoot.Count -eq 0) {
                    if (Test-ApprovedForDeletion -Path $rootNorm -AllowRoot $rootNorm -AllowRootItself) {
                        [System.IO.Directory]::Delete($rootNorm, $false)
                    }
                }
            }
            catch {
                Write-Log -Level WARN -Event 'delete_failed' -Data @{
                    path    = $rootNorm
                    message = $_.Exception.Message
                }
            }
        }
    }

    return ,([ordered]@{ deleted = $deletedHere; failed = $failedHere })
}

function Clear-ApprovedRecycleBin {
    if ((Test-DryRun)) {
        throw 'Invariante rota: la Papelera no se vacia en dry-run.'
    }
    $command = Get-Command -Name Clear-RecycleBin -ErrorAction SilentlyContinue
    if (-not $command) { throw 'Clear-RecycleBin no esta disponible en esta version de PowerShell.' }
    Clear-RecycleBin -Force -ErrorAction Stop
    Write-Log -Level INFO -Event 'recyclebin_cleared' -Data @{ forced = [bool]$Force }
}


function Confirm-Deletion {
    param(
        [int64]$FileCount,
        [int64]$Bytes,
        [int]$RecycleItems
    )

    if ((Test-DryRun)) { return $false }
    if ($Force) {
        Write-Log -Level WARN -Event 'confirmation_forced' -Data @{
            fileCount    = $FileCount
            sizeBytes    = $Bytes
            recycleItems = $RecycleItems
        }
        return $true
    }
    if (-not (Test-InteractiveSession)) {
        throw 'La consola no es interactiva. No se borro nada. Repita con -Apply -Force para una tarea programada.'
    }

    Write-Host ''
    Write-Host ("Se eliminarian {0} archivos ({1}) y {2} elementos de la Papelera." -f $FileCount, (Format-ByteSize -Bytes $Bytes), $RecycleItems)
    $answer = Read-Host 'Escriba SI en mayusculas para continuar'
    $accepted = ($answer -eq 'SI')
    Write-Log -Level INFO -Event 'confirmation_answer' -Data @{ accepted = $accepted }
    return $accepted
}

function Initialize-Whitelist {
    $paths = New-Object System.Collections.Generic.List[string]
    $seen = @{}

    $folderNames = @(
        'Desktop',
        'MyDocuments',
        'MyPictures',
        'MyMusic',
        'MyVideos',
        'Favorites'
    )
    foreach ($name in $folderNames) {
        try { $paths.Add([Environment]::GetFolderPath($name)) } catch { }
    }

    $extra = @(
        (Join-Path $env:USERPROFILE 'Downloads'),
        (Join-Path $env:USERPROFILE 'Documents'),
        (Join-Path $env:USERPROFILE 'Desktop'),
        (Join-Path $env:USERPROFILE 'Pictures'),
        (Join-Path $env:USERPROFILE 'Music'),
        (Join-Path $env:USERPROFILE 'Videos'),
        (Join-Path $env:USERPROFILE 'OneDrive'),
        (Join-Path $env:USERPROFILE '.ssh'),
        (Join-Path $env:USERPROFILE '.gnupg'),
        (Join-Path $env:USERPROFILE 'Contacts'),
        (Join-Path ($env:SystemDrive + '\') 'Windows\System32'),
        (Join-Path ($env:SystemDrive + '\') 'Windows\SysWOW64'),
        (Join-Path ($env:SystemDrive + '\') 'Windows\WinSxS'),
        (Join-Path ($env:SystemDrive + '\') 'Windows\Fonts'),
        (Join-Path ($env:SystemDrive + '\') 'Windows\SystemApps'),
        (Join-Path ($env:SystemDrive + '\') 'Program Files'),
        (Join-Path ($env:SystemDrive + '\') 'Program Files (x86)'),
        (Get-RepositoryRoot)
    )
    foreach ($item in $extra) { $paths.Add($item) }
    if ($Whitelist) {
        foreach ($item in $Whitelist) {
            if ([string]::IsNullOrWhiteSpace($item)) { continue }
            if (-not [System.IO.Path]::IsPathRooted($item)) {
                throw "La whitelist adicional debe ser una ruta absoluta: $item"
            }
            $paths.Add($item)
        }
    }

    $normalized = New-Object System.Collections.Generic.List[string]
    foreach ($item in $paths) {
        if ([string]::IsNullOrWhiteSpace($item)) { continue }
        $full = Get-NormalizedPath -Path $item
        if ([string]::IsNullOrWhiteSpace($full)) { continue }
        if (Test-IsDriveRoot -Path $full) {
            throw "La whitelist no puede ser una unidad completa: $full"
        }
        $key = $full.ToUpperInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $normalized.Add($full)
    }
    return ,$normalized
}

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }

    if ($Force -and -not $Apply) {
        Write-Warning '-Force no cambia el dry-run. Hace falta -Apply para borrar.'
    }
    if ($WhatIfPreference) {
        Set-DryRunMode -DryRun $true
    }

    $repoRoot = Get-RepositoryRoot
    $stamp = [DateTime]::Now.ToString('yyyyMMdd-HHmmss')
    if ([string]::IsNullOrWhiteSpace($LogPath)) {
        $LogPath = Join-Path (Join-Path $repoRoot 'logs') 'SmartCleanup.jsonl'
    }
    if ([string]::IsNullOrWhiteSpace($ReportPath)) {
        $ReportPath = Join-Path (Join-Path $repoRoot 'reports') ("SmartCleanup-{0}.json" -f $stamp)
    }
    $ReportPath = [System.IO.Path]::GetFullPath($ReportPath)
    Initialize-CommonSession -ScriptName 'SmartCleanup' -LogPath ([System.IO.Path]::GetFullPath($LogPath)) -DryRun (Test-DryRun)
    Set-LogProperty -Name 'apply' -Value ([bool]$script:Apply)

    $script:Rejected = New-Object System.Collections.Generic.List[object]
    $whitelistList = Initialize-Whitelist
    $script:Whitelist = $whitelistList

    $categoryNames = Get-SelectedCategoryNames
    if ($null -eq $categoryNames) { $categoryNames = @() }

    $modeLabel = 'DRY-RUN'
    if (-not (Test-DryRun)) { $modeLabel = 'APPLY' }

    Write-Host 'SmartCleanup 1.0'
    Write-Host ("Modo: {0}" -f $modeLabel)
    if ((Test-DryRun)) {
        Write-Host 'Dry-run predeterminado: no se borrara ningun archivo.'
    }
    else {
        Write-Host 'APPLY armado. No se borra nada hasta la confirmacion y la whitelist.'
    }
    Write-Host ("Log: {0}" -f (Get-SessionLogPath))
    Write-Host ("Categorias: {0}" -f ($categoryNames -join ', '))
    Write-Host ("Antiguedad minima: {0} dias" -f $OlderThanDays)

    Write-Log -Level INFO -Event 'run_start' -Data @{
        categories    = ($categoryNames -join ',')
        olderThanDays = $OlderThanDays
        top           = $Top
        force         = [bool]$Force
        reportPath    = $ReportPath
    }

    $targets = Get-CleanupTargets -Categories $categoryNames
    if ($null -eq $targets) { $targets = @() }

    $results = New-Object System.Collections.Generic.List[object]
    $totalBytes = [int64]0
    $totalFiles = [int64]0
    foreach ($target in $targets) {
        Write-Host ("Midiendo {0}: {1}" -f $target.category, $target.path)
        $measured = Measure-Target -Target $target
        $results.Add($measured)
        if ($measured.included -and $null -ne $measured.sizeBytes) {
            $totalBytes += [int64]$measured.sizeBytes
            $totalFiles += [int64]$measured.fileCount
        }
        $level = 'INFO'
        if ($measured.error) { $level = 'WARN' }
        Write-Log -Level $level -Event 'target_measured' -Data @{
            category   = [string]$measured.category
            path       = [string]$measured.path
            exists     = [bool]$measured.exists
            included   = [bool]$measured.included
            sizeBytes  = [int64]$measured.sizeBytes
            fileCount  = [int64]$measured.fileCount
            complete   = [bool]$measured.sizeIsComplete
            durationMs = [int64]$measured.durationMs
            message    = [string]$measured.error
        }
    }

    $recycleReport = [ordered]@{
        included  = $false
        itemCount = 0
        sizeBytes = [int64]0
        error     = $null
        largest   = (New-Object System.Collections.Generic.List[object])
    }
    $cleanRecycle = $false
    foreach ($name in $categoryNames) {
        if ($name -eq 'RecycleBin') { $cleanRecycle = $true }
    }
    if ($cleanRecycle) {
        Write-Host 'Leyendo la Papelera de reciclaje...'
        $recycleReport = Get-RecycleBinReport -Limit $Top
        Write-Log -Level INFO -Event 'recyclebin_measured' -Data @{
            itemCount = [int]$recycleReport.itemCount
            sizeBytes = [int64]$recycleReport.sizeBytes
            message   = [string]$recycleReport.error
        }
    }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($item in $results) {
        $sizeLabel = ''
        if ($item.included) { $sizeLabel = Format-ByteSize -Bytes ([int64]$item.sizeBytes) }
        $note = [string]$item.error
        if ($note.Length -gt 90) { $note = $note.Substring(0, 90) }
        $rows.Add([pscustomobject]@{
            Categoria = $item.category
            Existe    = $item.exists
            Archivos  = $(if ($item.included) { $item.fileCount } else { '' })
            Tamano    = $sizeLabel
            Completo  = $(if ($item.included) { $item.sizeIsComplete } else { '' })
            Ruta      = $item.path
            Nota      = $note
        })
    }
    Write-ConsoleTable -Title 'Objetivos' -Rows $rows

    if ($cleanRecycle) {
        $recycleRows = New-Object System.Collections.Generic.List[object]
        foreach ($item in $recycleReport.largest) {
            $recycleRows.Add([pscustomobject]@{
                Tamano   = (Format-ByteSize -Bytes ([int64]$item.sizeBytes))
                Elemento = $(if ([string]::IsNullOrWhiteSpace([string]$item.path)) { $item.name } else { $item.path })
            })
        }
        Write-Host ''
        Write-Host ("Papelera: {0} elementos, {1}." -f $recycleReport.itemCount, (Format-ByteSize -Bytes ([int64]$recycleReport.sizeBytes)))
        if ($recycleReport.error) { Write-Host ("Aviso de Papelera: {0}" -f $recycleReport.error) }
        Write-ConsoleTable -Title 'Elementos mas grandes de la Papelera' -Rows $recycleRows
    }

    $topPool = New-Object System.Collections.Generic.List[object]
    foreach ($item in $results) {
        foreach ($file in $item.largestFiles) {
            $topPool.Add([ordered]@{
                category         = [string]$item.category
                path             = [string]$file.path
                sizeBytes        = [int64]$file.sizeBytes
                lastWriteTimeUtc = [string]$file.lastWriteTimeUtc
            })
        }
    }
    $topRows = New-Object System.Collections.Generic.List[object]
    $topFiles = New-Object System.Collections.Generic.List[object]
    $sortedTop = @($topPool.ToArray() | Sort-Object { $_.sizeBytes } -Descending | Select-Object -First $Top)
    foreach ($file in $sortedTop) {
        if ($null -eq $file) { continue }
        $topFiles.Add($file)
        $topRows.Add([pscustomobject]@{
            Tamano    = (Format-ByteSize -Bytes ([int64]$file.sizeBytes))
            Categoria = $file.category
            Archivo   = $file.path
        })
    }
    Write-ConsoleTable -Title ("Archivos mas grandes que se limpiarian (top {0})" -f $Top) -Rows $topRows

    $rejectedRows = New-Object System.Collections.Generic.List[object]
    foreach ($item in $script:Rejected) {
        $rejectedRows.Add([pscustomobject]@{
            Categoria = $item.category
            Ruta      = $item.path
            Motivo    = $item.reason
        })
    }
    if ($rejectedRows.Count -gt 0) {
        Write-ConsoleTable -Title 'Objetivos rechazados por seguridad' -Rows $rejectedRows
    }

    $deletionPerformed = $false
    $cancelled = $false
    $confirmed = $false
    if ((Test-DryRun)) {
        Write-Log -Level INFO -Event 'deletion_skipped' -Data @{ reason = 'DryRun' }
        Write-Host ''
        Write-Host ("DRY-RUN: no se escribio {0} y no se elimino ningun archivo." -f $ReportPath)
    }
    else {
        $confirmed = Confirm-Deletion -FileCount $totalFiles -Bytes $totalBytes -RecycleItems ([int]$recycleReport.itemCount)
        if (-not $confirmed) {
            $cancelled = $true
            Write-Log -Level INFO -Event 'deletion_cancelled' -Data @{ reason = 'El usuario no escribio SI.' }
            Write-Host 'Operacion cancelada. No se elimino ningun archivo.'
        }
        else {
            foreach ($target in $results) {
                if (-not $target.included) { continue }
                if (-not $target.exists) { continue }
                $removeRoot = ($target.mode -eq 'RemoveTree')
                $pattern = $null
                foreach ($source in $targets) {
                    if ([string]$source.path -eq [string]$target.path -and [string]$source.name -eq [string]$target.name) {
                        $pattern = [string]$source.pattern
                    }
                }
                Write-Log -Level INFO -Event 'delete_started' -Data @{ path = [string]$target.path }
                $outcome = Remove-ApprovedChildren -Root ([string]$target.path) -Pattern $pattern -RemoveRootWhenEmpty:$removeRoot
                Write-Log -Level INFO -Event 'delete_finished' -Data @{
                    path    = [string]$target.path
                    deleted = [int]$outcome.deleted
                    failed  = [int]$outcome.failed
                }
            }
            if ($cleanRecycle -and [int]$recycleReport.itemCount -gt 0) {
                Clear-ApprovedRecycleBin
            }
            $deletionPerformed = $true
            Ensure-Directory -Path (Split-Path -Parent $ReportPath)
        }
    }

    $stopwatch.Stop()
    $recommendations = New-Object System.Collections.Generic.List[string]
    if ((Test-DryRun)) {
        $recommendations.Add(("Dry-run: se podrian liberar {0} en {1} archivos de caches y temporales, mas {2} en la Papelera." -f (Format-ByteSize -Bytes $totalBytes), $totalFiles, (Format-ByteSize -Bytes ([int64]$recycleReport.sizeBytes))))
        $recommendations.Add('No se elimino ningun archivo. Para borrar, ejecute de nuevo con -Apply y confirme con SI.')
    }
    elseif ($cancelled) {
        $recommendations.Add('La confirmacion no fue SI. No se elimino ningun archivo.')
    }
    else {
        $recommendations.Add(("Se eliminaron {0} archivos ({1})." -f $script:DeletedFiles, (Format-ByteSize -Bytes $script:DeletedBytes)))
    }
    if ($script:Rejected.Count -gt 0) {
        $recommendations.Add(("{0} objetivo(s) se rechazaron por la whitelist o por ser rutas criticas." -f $script:Rejected.Count))
    }

    $cleanupStatus = 'Ok'
    if ($script:Rejected.Count -gt 0) { $cleanupStatus = 'Warning' }
    foreach ($measured in $results) {
        $measuredAccess = 0
        $measuredComplete = $true
        if ($measured -is [System.Collections.IDictionary]) {
            if ((Test-DictionaryKey -Dictionary $measured -Name 'accessErrors') -and $null -ne $measured['accessErrors']) {
                $measuredAccess = [int]$measured['accessErrors']
            }
            if (Test-DictionaryKey -Dictionary $measured -Name 'sizeIsComplete') {
                $measuredComplete = [bool]$measured['sizeIsComplete']
            }
        }
        if ($measuredAccess -gt 0 -or -not $measuredComplete) { $cleanupStatus = 'Warning' }
    }

    $report = [ordered]@{
        schemaVersion = '1.0'
        generatedAt   = (Get-UtcTimestamp)
        dryRun        = [bool](Test-DryRun)
        apply         = [bool]$script:Apply
        confirmed     = [bool]$confirmed
        cancelled     = [bool]$cancelled
        deletionPerformed = [bool]$deletionPerformed
        hostname      = [string]$env:COMPUTERNAME
        userName      = [string]$env:USERNAME
        runId         = (Get-RunId)
        parameters    = [ordered]@{
            categories    = $categoryNames
            olderThanDays = $OlderThanDays
            top           = $Top
            force         = [bool]$Force
            logPath       = (Get-SessionLogPath)
            reportPath    = $ReportPath
        }
        whitelist     = $script:Whitelist
        rejected      = $script:Rejected
        targets       = $results
        recycleBin    = $recycleReport
        largestFiles  = $topFiles
        summary       = [ordered]@{
            overallStatus      = $cleanupStatus
            targetCount        = $results.Count
            fileCount          = $totalFiles
            sizeBytes          = $totalBytes
            recycleItemCount   = [int]$recycleReport.itemCount
            recycleSizeBytes   = [int64]$recycleReport.sizeBytes
            deletedFiles       = [int]$script:DeletedFiles
            deletedBytes       = [int64]$script:DeletedBytes
            durationMs         = [int64]$stopwatch.ElapsedMilliseconds
            recommendations    = $recommendations
        }
    }

    $json = ConvertTo-ReportJson -Report $report
    if ($deletionPerformed) {
        [void](Write-JsonReport -Path $ReportPath -Json $json)
        Write-Log -Level INFO -Event 'report_written' -Data @{ reportPath = $ReportPath }
        Write-Host ("Reporte escrito: {0}" -f $ReportPath)
    }
    else {
        Write-Log -Level INFO -Event 'report_skipped' -Data @{
            reportPath = $ReportPath
            reason     = $(if ((Test-DryRun)) { 'DryRun' } else { 'Cancelled' })
        }
    }

    Write-Log -Level INFO -Event 'run_completed' -Data @{
        dryRun       = [bool](Test-DryRun)
        deletedFiles = [int]$script:DeletedFiles
        sizeBytes    = $totalBytes
        durationMs   = [int64]$stopwatch.ElapsedMilliseconds
    }

    Write-Host ''
    Write-Host ("Total estimado en caches y temporales: {0} en {1} archivos. Duracion: {2:N1} s." -f (Format-ByteSize -Bytes $totalBytes), $totalFiles, ($stopwatch.ElapsedMilliseconds / 1000))
    Write-Output $json
    $script:ExitCode = Get-ReportExitCode -Status $cleanupStatus
}
catch {
    Write-ScriptError -ErrorRecord $_
    $script:ExitCode = (Get-FatalExitCode)
}

Write-Host ("Codigo de salida: {0}" -f $script:ExitCode)
if ($MyInvocation.InvocationName -ne '.') {
    exit $script:ExitCode
}
