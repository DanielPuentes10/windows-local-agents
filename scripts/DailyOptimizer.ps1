#Requires -Version 5.1
<#
.SYNOPSIS
    Rutina diaria de optimizacion de Windows 11, en dry-run salvo que se pida -Apply.

.DESCRIPTION
    Ejecuta, cada una con su propio manejo de errores:

      - Temporales: mide en una sola pasada de solo lectura y, solo con
        -Apply, borra archivos mas viejos que -OlderThanDays dentro de
        TEMP del usuario y de Windows\Temp.
      - Actualizaciones: lee reports\WindowsUpdateSnapshot.json si
        UpdateChecker lo escribio dentro de -UpdateCacheHours. No abre
        el catalogo de Windows Update. El reinicio pendiente se lee del
        registro local. Nunca descarga ni instala actualizaciones.
      - Memoria: informa el uso y, solo con -Apply, recorta el conjunto de
        trabajo de procesos de la sesion actual que superan el umbral.
        No vacia la lista de espera ni cierra procesos.
      - Procesos: muestra los que mas CPU y RAM consumen. Nunca los detiene.
      - Indice: informa Windows Search y, solo con -Apply, inicia la tarea
        integrada IndexerAutomaticMaintenance. No reconstruye el indice
        ni compacta Windows.edb.

    Sin -Apply no se borra, no se recorta memoria y no se inicia el
    mantenimiento. -Force solo omite la confirmacion SI; no activa el
    borrado por si solo.

    Si no se indica ninguna tarea, se ejecutan las cinco.

.PARAMETER TempFiles
    Limpia temporales. Si no se pasa ninguna tarea, entra igual.

.PARAMETER PendingUpdates
    Revisa actualizaciones pendientes. No instala nada.

.PARAMETER FreeMemory
    Evalua y, con -Apply, recorta conjuntos de trabajo elegibles.

.PARAMETER ReviewProcesses
    Lista los procesos que mas recursos consumen.

.PARAMETER SearchIndex
    Revisa el indice de busqueda y, con -Apply, dispara su mantenimiento.

.PARAMETER All
    Ejecuta las cinco tareas. Es el comportamiento por defecto.

.PARAMETER Apply
    Aplica las acciones que modifican el equipo. Sin este modificador
    el script solo informa.

.PARAMETER Force
    Con -Apply, no pide la confirmacion SI. Pensado para el Programador
    de tareas. Sin -Apply no cambia nada.

.PARAMETER OlderThanDays
    Antiguedad minima, en dias, de un temporal para contarlo y poder borrarlo.
    De 0 a 3650. 0 incluye todos. Predeterminado: 7. La pasada sigue siendo
    de solo lectura en dry-run.

.PARAMETER Top
    Cuantos procesos o archivos grandes se listan. De 1 a 50.
    Predeterminado: 10.

.PARAMETER MemoryThresholdMB
    Conjunto de trabajo minimo, en MB, para ofrecer un proceso al recorte
    y para marcarlo como pesado. De 50 a 8192. Predeterminado: 500.

.PARAMETER CpuSampleSeconds
    Segundos entre las dos muestras de CPU. De 1 a 15. Predeterminado: 2.

.PARAMETER UpdateTimeoutSeconds
    Se conserva para tareas ya creadas. Esta rutina ya no espera una
    busqueda del catalogo. De 5 a 120. Predeterminado: 45.

.PARAMETER UpdateCacheHours
    Edad maxima, en horas, del snapshot de UpdateChecker que se reutiliza.
    De 1 a 168. Predeterminado: 20. No provoca una consulta nueva.

.PARAMETER Whitelist
    Rutas absolutas extra que no se pueden borrar ni recorrer.

.PARAMETER ReportPath
    JSON del resultado. Si se omite, se crea
    reports\DailyOptimizer-yyyyMMdd-HHmmss.json.
    Se escribe tambien en dry-run.

.PARAMETER LogPath
    JSONL de auditoria. Predeterminado: logs\DailyOptimizer.jsonl.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\DailyOptimizer.ps1

    Dry-run de las cinco tareas.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\DailyOptimizer.ps1 -TempFiles -ReviewProcesses -OlderThanDays 7

    Solo mide temporales de mas de 7 dias y revisa procesos.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\DailyOptimizer.ps1 -TempFiles -SearchIndex -OlderThanDays 7 -Apply -Force

    Borra temporales viejos e inicia el mantenimiento del indice, sin preguntar.

.NOTES
    Autor: LocalAgents
    Requiere: Windows PowerShell 5.1
    Codigos de salida (Common.psm1):
      0  el reporte termino con estado Ok
      1  el reporte termino con estado Warning
      2  el reporte termino con estado Critical o Error
      3  el script fallo antes de cerrar el reporte
    Helpers compartidos: scripts\Common.psm1.

    Programador de tareas (cada dia a las 07:45). No instala actualizaciones
    y no consulta el catalogo: UpdateChecker debe haber corrido antes, una
    sola vez, para dejar el snapshot. El recorte de memoria queda fuera de
    este ejemplo porque puede provocar una pausa breve al volver a paginar
    las aplicaciones:
      $script = 'C:\ProyectosDeSoftware\LocalAgents\scripts\DailyOptimizer.ps1'
      $arg = "-NoProfile -ExecutionPolicy Bypass -File `"$script`" -TempFiles -PendingUpdates -ReviewProcesses -SearchIndex -OlderThanDays 7 -Apply -Force"
      $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arg
      $trigger = New-ScheduledTaskTrigger -Daily -At 7:45am
      Register-ScheduledTask -TaskName 'LocalAgents-DailyOptimizer' -Action $action -Trigger $trigger -Description 'Optimizacion diaria en modo apply'
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$TempFiles,
    [switch]$PendingUpdates,
    [switch]$FreeMemory,
    [switch]$ReviewProcesses,
    [switch]$SearchIndex,
    [switch]$All,

    [switch]$Apply,
    [switch]$Force,

    [ValidateRange(0, 3650)]
    [int]$OlderThanDays = 7,

    [ValidateRange(1, 50)]
    [int]$Top = 10,

    [ValidateRange(50, 8192)]
    [int]$MemoryThresholdMB = 500,

    [ValidateRange(1, 15)]
    [int]$CpuSampleSeconds = 2,

    [ValidateRange(5, 120)]
    [int]$UpdateTimeoutSeconds = 45,

    [ValidateRange(1, 168)]
    [int]$UpdateCacheHours = 20,

    [string[]]$Whitelist,

    [string]$ReportPath,

    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -Force -DisableNameChecking

$script:ExitCode = 0
$script:Whitelist = $null
$script:UpdateJob = $null
$script:UpdateTemp = $null
$script:UpdateDeadline = $null
$script:MemoryNativeReady = $false


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
    param([Parameter(Mandatory = $true)][string]$Path)
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
    return ,([ordered]@{ ok = $true; path = $normalized; reason = $null })
}

function Test-ApprovedForDeletion {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$AllowRoot
    )
    $normalized = Get-NormalizedPath -Path $Path
    $root = Get-NormalizedPath -Path $AllowRoot
    if ([string]::IsNullOrWhiteSpace($normalized) -or [string]::IsNullOrWhiteSpace($root)) { return $false }
    if (Test-IsDriveRoot -Path $normalized) { return $false }
    if (Test-ForbiddenRoot -Path $normalized) { return $false }
    if (Test-ProtectedPath -Path $normalized) { return $false }
    if (Test-ProtectedPath -Path $root) { return $false }
    if ($normalized.Equals($root, [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
    if (-not (Test-IsUnderPath -Path $normalized -Root $root)) { return $false }
    return $true
}

function Test-KnownTempRoot {
    param([Parameter(Mandatory = $true)][string]$Path)
    $normalized = Get-NormalizedPath -Path $Path
    if ([string]::IsNullOrWhiteSpace($normalized)) { return $false }
    $known = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($env:TEMP)) { $known.Add((Get-NormalizedPath -Path $env:TEMP)) }
    if (-not [string]::IsNullOrWhiteSpace($env:SystemRoot)) { $known.Add((Get-NormalizedPath -Path (Join-Path $env:SystemRoot 'Temp'))) }
    if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) { $known.Add((Get-NormalizedPath -Path (Join-Path $env:LOCALAPPDATA 'Temp'))) }
    foreach ($item in $known) {
        if ([string]::IsNullOrWhiteSpace($item)) { continue }
        if ($normalized.Equals($item, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Initialize-Whitelist {
    $paths = New-Object System.Collections.Generic.List[string]
    $folderNames = @('Desktop', 'MyDocuments', 'MyPictures', 'MyMusic', 'MyVideos', 'Favorites')
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
    $seen = @{}
    foreach ($item in $paths) {
        if ([string]::IsNullOrWhiteSpace($item)) { continue }
        $full = Get-NormalizedPath -Path $item
        if ([string]::IsNullOrWhiteSpace($full)) { continue }
        $key = $full.ToLowerInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $normalized.Add($full)
    }
    return ,$normalized.ToArray()
}


function New-SkippedTask {
    param([string]$Note)
    return [ordered]@{
        status  = 'Skipped'
        enabled = $false
        error   = $null
        note    = $Note
    }
}

function Add-LargestItem {
    param(
        $Items,
        [int]$Limit,
        $Entry
    )
    $kept = Get-ItemCount -Value $Items
    if ($kept -lt $Limit) {
        $Items.Add($Entry)
        return
    }
    $minIndex = -1
    $minBytes = [int64]::MaxValue
    $index = 0
    foreach ($existing in $Items) {
        $size = [int64]$existing.bytes
        if ($size -lt $minBytes) {
            $minBytes = $size
            $minIndex = $index
        }
        $index++
    }
    if ($minIndex -ge 0 -and [int64]$Entry.bytes -gt $minBytes) {
        $Items.RemoveAt($minIndex)
        $Items.Add($Entry)
    }
}

function Get-CutoffUtc {
    if ($OlderThanDays -le 0) { return [DateTime]::MinValue }
    return [DateTime]::UtcNow.AddDays(-1 * $OlderThanDays)
}

function Get-TempRoots {
    $items = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    $candidates = @(
        @{ name = 'UserTemp'; path = [string]$env:TEMP },
        @{ name = 'WindowsTemp'; path = $(if ([string]::IsNullOrWhiteSpace($env:SystemRoot)) { '' } else { Join-Path $env:SystemRoot 'Temp' }) }
    )
    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace([string]$candidate.path)) { continue }
        $normalized = Get-NormalizedPath -Path ([string]$candidate.path)
        if ([string]::IsNullOrWhiteSpace($normalized)) { continue }
        $key = $normalized.ToLowerInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $items.Add([ordered]@{
            name = [string]$candidate.name
            path = $normalized
        })
    }
    return ,$items.ToArray()
}

function Measure-TempRoot {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $decision = Test-AllowRootSafe -Path $Path
    if (-not $decision.ok) {
        return [ordered]@{
            name       = $Name
            path       = [string]$decision.path
            accepted   = $false
            reason     = [string]$decision.reason
            fileCount  = 0
            bytes      = [int64]0
            errorCount = 0
            error      = $null
            largest    = (New-Object System.Collections.Generic.List[object])
        }
    }
    if (-not (Test-KnownTempRoot -Path ([string]$decision.path))) {
        return [ordered]@{
            name       = $Name
            path       = [string]$decision.path
            accepted   = $false
            reason     = 'La ruta no es una carpeta Temp conocida.'
            fileCount  = 0
            bytes      = [int64]0
            errorCount = 0
            error      = $null
            largest    = (New-Object System.Collections.Generic.List[object])
        }
    }
    if (-not (Test-Path -LiteralPath ([string]$decision.path))) {
        return [ordered]@{
            name       = $Name
            path       = [string]$decision.path
            accepted   = $false
            reason     = 'La carpeta no existe.'
            fileCount  = 0
            bytes      = [int64]0
            errorCount = 0
            error      = $null
            largest    = (New-Object System.Collections.Generic.List[object])
        }
    }

    $cutoffTicks = [int64]-1
    if ($OlderThanDays -gt 0) { $cutoffTicks = [int64](Get-CutoffUtc).Ticks }
    $protected = @()
    if ($null -ne $script:Whitelist) { $protected = @($script:Whitelist) }
    $walk = Measure-TreeFiles -Root ([string]$decision.path) -CutoffUtcTicks $cutoffTicks -Top $Top -ProtectedRoots $protected
    $largest = New-Object System.Collections.Generic.List[object]
    foreach ($hit in $walk.largest) {
        if ($null -eq $hit) { continue }
        $largest.Add([ordered]@{
            path         = [string]$hit.path
            bytes        = [int64]$hit.sizeBytes
            lastWriteUtc = [string]$hit.lastWriteTimeUtc
        })
    }
    $firstError = $null
    if ([int]$walk.accessErrors -gt 0) { $firstError = [string]$walk.error }

    return [ordered]@{
        name       = $Name
        path       = [string]$decision.path
        accepted   = $true
        reason     = $null
        fileCount  = [int64]$walk.fileCount
        bytes      = [int64]$walk.sizeBytes
        errorCount = [int]$walk.accessErrors
        error      = $firstError
        largest    = $largest
        durationMs = [int64]$walk.durationMs
    }
}

function Get-TempReport {
    $roots = New-Object System.Collections.Generic.List[object]
    $largest = New-Object System.Collections.Generic.List[object]
    $fileCount = 0
    $bytes = [int64]0
    $acceptedRoots = 0
    $failedRoots = 0
    $rootErrors = 0

    foreach ($root in (Get-TempRoots)) {
        $measured = Measure-TempRoot -Name ([string]$root.name) -Path ([string]$root.path)
        $roots.Add($measured)
        if ($measured.accepted) { $acceptedRoots++ } else { $failedRoots++ }
        if ($measured.error) { $rootErrors++ }
        $fileCount += [int]$measured.fileCount
        $bytes += [int64]$measured.bytes
        foreach ($item in $measured.largest) {
            Add-LargestItem -Items $largest -Limit $Top -Entry $item
        }
        Write-Log -Level INFO -Event 'temp_measured' -Data @{
            name      = [string]$measured.name
            path      = [string]$measured.path
            accepted  = [bool]$measured.accepted
            fileCount = [int]$measured.fileCount
            bytes     = [int64]$measured.bytes
            reason    = [string]$measured.reason
            error     = [string]$measured.error
        }
    }

    $status = 'Ok'
    if ($fileCount -gt 0 -or $rootErrors -gt 0 -or $failedRoots -gt 0) { $status = 'Warning' }
    if ($acceptedRoots -eq 0) { $status = 'Error' }

    return [ordered]@{
        status          = $status
        enabled         = $true
        olderThanDays   = $OlderThanDays
        acceptedRoots   = $acceptedRoots
        fileCount       = $fileCount
        bytes           = $bytes
        deletedFiles    = 0
        deletedBytes    = [int64]0
        deleteFailed    = 0
        applied         = $false
        roots           = $roots
        largest         = $largest
        error           = $null
        note            = 'Dry-run no borra. -Apply borra solo estos temporales.'
    }
}

function Remove-TempRootContents {
    param([Parameter(Mandatory = $true)][string]$Root)

    if ((Test-DryRun)) {
        throw 'Invariante rota: no se borran temporales en dry-run.'
    }
    if (-not (Test-KnownTempRoot -Path $Root)) {
        throw "La raiz dejo de ser una carpeta Temp conocida: $Root"
    }
    $decision = Test-AllowRootSafe -Path $Root
    if (-not $decision.ok) {
        throw ("La raiz dejo de ser segura: {0}" -f [string]$decision.reason)
    }

    $cutoff = Get-CutoffUtc
    $directories = New-Object System.Collections.Generic.List[string]
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push([string]$decision.path)
    $deletedHere = 0
    $failedHere = 0
    $deletedBytes = [int64]0

    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        try {
            $directory = New-Object System.IO.DirectoryInfo $current
            foreach ($file in $directory.EnumerateFiles()) {
                try {
                    if (($file.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                    if ($file.LastWriteTimeUtc -ge $cutoff -and $OlderThanDays -gt 0) { continue }
                    if (-not (Test-ApprovedForDeletion -Path $file.FullName -AllowRoot ([string]$decision.path))) { continue }
                    $length = [int64]$file.Length
                    [System.IO.File]::SetAttributes($file.FullName, [System.IO.FileAttributes]::Normal)
                    [System.IO.File]::Delete($file.FullName)
                    $deletedHere++
                    $deletedBytes += $length
                }
                catch {
                    $failedHere++
                    Write-Log -Level WARN -Event 'delete_failed' -Data @{
                        path    = $file.FullName
                        message = $_.Exception.Message
                    }
                }
            }
            foreach ($child in $directory.EnumerateDirectories()) {
                if (($child.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                if (-not (Test-ApprovedForDeletion -Path $child.FullName -AllowRoot ([string]$decision.path))) { continue }
                $stack.Push($child.FullName)
                $directories.Add($child.FullName)
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

    $sortedDirs = @($directories.ToArray() | Sort-Object { $_.Length } -Descending)
    foreach ($dirPath in $sortedDirs) {
        if ([string]::IsNullOrWhiteSpace([string]$dirPath)) { continue }
        if (-not (Test-ApprovedForDeletion -Path ([string]$dirPath) -AllowRoot ([string]$decision.path))) { continue }
        try {
            $remaining = New-Object System.IO.DirectoryInfo ([string]$dirPath)
            $hasChild = $false
            foreach ($entry in $remaining.EnumerateFileSystemInfos()) {
                $hasChild = $true
                break
            }
            if (-not $hasChild) {
                [System.IO.Directory]::Delete([string]$dirPath, $false)
            }
        }
        catch {
            Write-Log -Level WARN -Event 'delete_failed' -Data @{
                path    = [string]$dirPath
                message = $_.Exception.Message
            }
        }
    }

    return ,([ordered]@{
        deleted      = $deletedHere
        deletedBytes = $deletedBytes
        failed       = $failedHere
    })
}


function Start-PendingUpdateSearch {
    param([int]$TimeoutSeconds)
    Write-Log -Level INFO -Event 'update_search_skipped' -Data @{
        reason                = 'La cadena diaria no abre el catalogo desde DailyOptimizer.'
        timeoutSecondsIgnored = $TimeoutSeconds
    }
}

function Clear-PendingUpdateJob {
    if ($script:UpdateJob) {
        try { Stop-Job -Job $script:UpdateJob -ErrorAction SilentlyContinue } catch { }
        try { Remove-Job -Job $script:UpdateJob -Force -ErrorAction SilentlyContinue } catch { }
        $script:UpdateJob = $null
    }
    if ($script:UpdateTemp -and (Test-Path -LiteralPath $script:UpdateTemp)) {
        try { [System.IO.File]::Delete($script:UpdateTemp) } catch { }
    }
    $script:UpdateTemp = $null
}

function Complete-PendingUpdateSearch {
    $serviceState = 'Unknown'
    $serviceStart = $null
    try {
        $service = Get-Service -Name 'wuauserv' -ErrorAction Stop
        $serviceState = [string]$service.Status
        $serviceStart = [string]$service.StartType
    }
    catch {
        $serviceState = 'NotFound'
    }

    $reboot = Get-RebootPending
    $updates = New-Object System.Collections.Generic.List[object]
    $queried = $false
    $errorText = $null
    $truncated = $false
    $queryMode = 'missing'
    $snapshotAt = $null
    $snapshot = Read-WindowsUpdateSnapshot -MaxAgeHours $UpdateCacheHours
    if ($null -ne $snapshot) {
        $queryMode = 'snapshot'
        $snapshotAt = Get-JsonText -Object $snapshot -Name 'generatedAt'
        $truncated = Get-JsonBool -Object $snapshot -Name 'truncated'
        $nodes = Get-JsonField -Object $snapshot -Name 'updates'
        if ($null -ne $nodes) {
            foreach ($update in @($nodes)) {
                if ($null -eq $update) { continue }
                $updates.Add([ordered]@{
                    title      = (Get-JsonText -Object $update -Name 'title')
                    kb         = (Get-JsonText -Object $update -Name 'kb')
                    severity   = (Get-JsonText -Object $update -Name 'severity')
                    downloaded = (Get-JsonBool -Object $update -Name 'downloaded')
                    priority   = (Get-JsonText -Object $update -Name 'priority')
                })
            }
        }
        $queried = $true
        $snapshotError = Get-JsonText -Object $snapshot -Name 'error'
        if (-not [string]::IsNullOrWhiteSpace($snapshotError)) { $errorText = $snapshotError }
        Write-Log -Level INFO -Event 'update_snapshot_used' -Data @{
            generatedAt = $snapshotAt
            count       = (Get-ItemCount -Value $updates)
        }
    }
    else {
        $errorText = 'No hay una consulta vigente de Windows Update. UpdateChecker es la unica que abre el catalogo y deja reports\WindowsUpdateSnapshot.json.'
        Write-Log -Level WARN -Event 'update_snapshot_missing' -Data @{ maxAgeHours = $UpdateCacheHours }
    }

    $pendingCount = Get-ItemCount -Value $updates
    $status = 'Ok'
    if ($reboot.pending -or $pendingCount -gt 0 -or $errorText -or $serviceState -eq 'NotFound') { $status = 'Warning' }

    return [ordered]@{
        status            = $status
        enabled           = $true
        serviceState      = $serviceState
        serviceStartType  = $serviceStart
        queried           = $queried
        queryMode         = $queryMode
        snapshotAt        = $snapshotAt
        pendingCount      = $pendingCount
        truncated         = $truncated
        rebootPending     = [bool]$reboot.pending
        rebootReasons     = $reboot.reasons
        installsPerformed = 0
        updates           = $updates
        error             = $errorText
        note              = 'Esta tarea nunca instala actualizaciones ni abre el catalogo de Windows Update.'
    }
}

function Get-ProcessSample {
    $processorCount = [int][Environment]::ProcessorCount
    if ($processorCount -lt 1) { $processorCount = 1 }
    $first = @{}
    $names = @{}
    $sessions = @{}
    $workingSets = @{}

    foreach ($process in @(Get-Process -ErrorAction SilentlyContinue)) {
        $processId = 0
        try { $processId = [int]$process.Id } catch { continue }
        $processName = ''
        try { $processName = [string]$process.ProcessName } catch { }
        $names[$processId] = $processName
        try { $sessions[$processId] = [int]$process.SessionId } catch { $sessions[$processId] = -1 }
        try { $workingSets[$processId] = [int64]$process.WorkingSet64 } catch { $workingSets[$processId] = [int64]0 }
        try { $first[$processId] = [int64]$process.TotalProcessorTime.Ticks } catch { }
    }

    Start-Sleep -Seconds $CpuSampleSeconds

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($process in @(Get-Process -ErrorAction SilentlyContinue)) {
        $processId = 0
        try { $processId = [int]$process.Id } catch { continue }
        $processName = ''
        try { $processName = [string]$process.ProcessName } catch { }
        if ([string]::IsNullOrWhiteSpace($processName) -and $names.ContainsKey($processId)) {
            $processName = [string]$names[$processId]
        }
        $sessionId = -1
        try { $sessionId = [int]$process.SessionId } catch { }
        $workingSet = [int64]0
        try { $workingSet = [int64]$process.WorkingSet64 } catch { }
        $cpuPercent = $null
        if ($first.ContainsKey($processId)) {
            try {
                $deltaTicks = [int64]$process.TotalProcessorTime.Ticks - [int64]$first[$processId]
                if ($deltaTicks -lt 0) { $deltaTicks = 0 }
                $deltaSeconds = $deltaTicks / [double][TimeSpan]::TicksPerSecond
                $capacity = $CpuSampleSeconds * $processorCount
                if ($capacity -gt 0) {
                    $cpuPercent = [math]::Round((100.0 * $deltaSeconds) / $capacity, 1)
                }
            }
            catch { }
        }
        $rows.Add([ordered]@{
            pid              = $processId
            name             = $processName
            sessionId        = $sessionId
            cpuPercent       = $cpuPercent
            workingSetBytes  = $workingSet
        })
    }

    return [ordered]@{
        sampleSeconds    = $CpuSampleSeconds
        processorCount   = $processorCount
        rows             = $rows
        error            = $null
    }
}

function Test-ReviewNoiseProcess {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $true }
    switch ($Name.ToLowerInvariant()) {
        'idle' { return $true }
        'system' { return $true }
        'registry' { return $true }
        'memory compression' { return $true }
        'secure system' { return $true }
        default { return $false }
    }
}

function Test-TrimProtectedProcess {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $true }
    switch ($Name.ToLowerInvariant()) {
        'idle' { return $true }
        'system' { return $true }
        'registry' { return $true }
        'memory compression' { return $true }
        'secure system' { return $true }
        'csrss' { return $true }
        'lsass' { return $true }
        'lsaiso' { return $true }
        'wininit' { return $true }
        'winlogon' { return $true }
        'services' { return $true }
        'smss' { return $true }
        'dwm' { return $true }
        'fontdrvhost' { return $true }
        'wmiprvse' { return $true }
        'msmpeng' { return $true }
        'nissrv' { return $true }
        'securityhealthservice' { return $true }
        'svchost' { return $true }
        'searchindexer' { return $true }
        'audiodg' { return $true }
        'spoolsv' { return $true }
        default { return $false }
    }
}

function Get-TopProcesses {
    param(
        $Rows,
        [string]$SortKey,
        [int]$Limit,
        [switch]$ExcludeNoise
    )
    $filtered = New-Object System.Collections.Generic.List[object]
    foreach ($row in $Rows) {
        if ($ExcludeNoise -and (Test-ReviewNoiseProcess -Name ([string]$row.name))) { continue }
        if ($SortKey -eq 'cpuPercent' -and $null -eq $row.cpuPercent) { continue }
        $filtered.Add($row)
    }
    $sorted = @($filtered.ToArray() | Sort-Object { $_.$SortKey } -Descending | Select-Object -First $Limit)
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($row in $sorted) {
        if ($null -eq $row) { continue }
        $items.Add($row)
    }
    return ,$items
}

function Get-MemorySnapshot {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $total = [int64]$os.TotalVisibleMemorySize * 1KB
    $free = [int64]$os.FreePhysicalMemory * 1KB
    $used = $total - $free
    $percent = 0.0
    if ($total -gt 0) { $percent = [math]::Round((100.0 * $used) / $total, 2) }
    return ,([ordered]@{
        totalBytes  = $total
        freeBytes   = $free
        usedBytes   = $used
        usedPercent = $percent
    })
}

function Build-MemoryReport {
    param($Sample)
    $memory = Get-MemorySnapshot
    $thresholdBytes = [int64]$MemoryThresholdMB * 1MB
    $currentSession = -1
    try { $currentSession = [int](Get-Process -Id $PID).SessionId } catch { }
    $eligible = New-Object System.Collections.Generic.List[object]
    $eligibleBytes = [int64]0
    foreach ($row in $Sample.rows) {
        if ([int64]$row.workingSetBytes -lt $thresholdBytes) { continue }
        if ([int]$row.sessionId -ne $currentSession) { continue }
        if (Test-TrimProtectedProcess -Name ([string]$row.name)) { continue }
        $eligibleBytes += [int64]$row.workingSetBytes
        $eligible.Add($row)
    }
    $eligibleCount = Get-ItemCount -Value $eligible
    $sortedEligible = Get-TopProcesses -Rows $eligible -SortKey 'workingSetBytes' -Limit 500
    $status = 'Ok'
    if ($memory.usedPercent -ge 90 -or $eligibleCount -gt 0) { $status = 'Warning' }
    return [ordered]@{
        status                 = $status
        enabled                = $true
        totalBytes             = $memory.totalBytes
        freeBytes              = $memory.freeBytes
        usedBytes              = $memory.usedBytes
        usedPercent            = $memory.usedPercent
        thresholdMb            = $MemoryThresholdMB
        sessionId              = $currentSession
        eligibleCount          = $eligibleCount
        eligibleWorkingSetBytes = $eligibleBytes
        trimmedCount           = 0
        applied                = $false
        eligible               = $sortedEligible
        error                  = $null
        note                   = 'El recorte vacia el conjunto de trabajo. No cierra procesos ni limpia la cache en espera.'
    }
}

function Build-ProcessReport {
    param($Sample)
    $thresholdBytes = [int64]$MemoryThresholdMB * 1MB
    $topCpu = Get-TopProcesses -Rows $Sample.rows -SortKey 'cpuPercent' -Limit $Top -ExcludeNoise
    $topMemory = Get-TopProcesses -Rows $Sample.rows -SortKey 'workingSetBytes' -Limit $Top -ExcludeNoise
    $heavy = 0
    foreach ($row in $Sample.rows) {
        if (Test-ReviewNoiseProcess -Name ([string]$row.name)) { continue }
        $cpuHigh = ($null -ne $row.cpuPercent -and [double]$row.cpuPercent -ge 15)
        $ramHigh = ([int64]$row.workingSetBytes -ge $thresholdBytes)
        if ($cpuHigh -or $ramHigh) { $heavy++ }
    }
    $status = 'Ok'
    if ($heavy -gt 0) { $status = 'Warning' }
    return [ordered]@{
        status           = $status
        enabled          = $true
        sampleSeconds    = $Sample.sampleSeconds
        processorCount   = $Sample.processorCount
        heavyCount       = $heavy
        stoppedCount     = 0
        topCpu           = $topCpu
        topMemory        = $topMemory
        error            = $null
        note             = 'La revision nunca detiene procesos. Idle y System se omiten del ranking.'
    }
}

function Initialize-MemoryNative {
    if ($script:MemoryNativeReady) { return }
    if (-not ('DailyOptimizerNative' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class DailyOptimizerNative {
    [DllImport("psapi.dll", SetLastError = true)]
    public static extern bool EmptyWorkingSet(IntPtr hProcess);
}
'@
    }
    $script:MemoryNativeReady = $true
}

function Invoke-WorkingSetTrim {
    param($Eligible)
    if ((Test-DryRun)) {
        throw 'Invariante rota: no se recorta memoria en dry-run.'
    }
    Initialize-MemoryNative
    $thresholdBytes = [int64]$MemoryThresholdMB * 1MB
    $currentSession = -1
    try { $currentSession = [int](Get-Process -Id $PID).SessionId } catch { }
    $trimmed = 0
    $failed = 0
    foreach ($row in $Eligible) {
        $process = $null
        try {
            $process = Get-Process -Id ([int]$row.pid) -ErrorAction Stop
            if ([string]$process.ProcessName -ne [string]$row.name) { continue }
            if ([int]$process.SessionId -ne $currentSession) { continue }
            if (Test-TrimProtectedProcess -Name ([string]$process.ProcessName)) { continue }
            if ([int64]$process.WorkingSet64 -lt $thresholdBytes) { continue }
            $handle = $process.Handle
            $ok = [DailyOptimizerNative]::EmptyWorkingSet($handle)
            if ($ok) { $trimmed++ } else { $failed++ }
        }
        catch {
            $failed++
            Write-Log -Level WARN -Event 'memory_trim_failed' -Data @{
                pid     = [int]$row.pid
                name    = [string]$row.name
                message = $_.Exception.Message
            }
        }
        finally {
            if ($process) { try { $process.Dispose() } catch { } }
        }
    }
    return ,([ordered]@{ trimmed = $trimmed; failed = $failed })
}

function Get-SearchReport {
    $serviceState = 'Unknown'
    $serviceStart = $null
    try {
        $service = Get-Service -Name 'WSearch' -ErrorAction Stop
        $serviceState = [string]$service.Status
        $serviceStart = [string]$service.StartType
    }
    catch {
        $serviceState = 'NotFound'
    }

    $indexPath = Join-Path $env:ProgramData 'Microsoft\Search\Data\Applications\Windows\Windows.edb'
    $indexBytes = $null
    $indexError = $null
    if (Test-Path -LiteralPath $indexPath) {
        try {
            $indexBytes = [int64](New-Object System.IO.FileInfo $indexPath).Length
        }
        catch {
            $indexError = $_.Exception.Message
        }
    }
    else {
        $indexError = 'Windows.edb no esta en la ruta predeterminada.'
    }

    $taskPath = '\Microsoft\Windows\Shell\'
    $taskName = 'IndexerAutomaticMaintenance'
    $taskFound = $false
    $taskState = $null
    $lastRun = $null
    $lastResult = $null
    $taskError = $null
    try {
        $task = Get-ScheduledTask -TaskPath $taskPath -TaskName $taskName -ErrorAction Stop
        $taskFound = $true
        $taskState = [string]$task.State
        $info = Get-ScheduledTaskInfo -InputObject $task
        $lastRun = Format-Utc -Value $info.LastRunTime
        $lastResult = [int]$info.LastTaskResult
    }
    catch {
        $taskError = $_.Exception.Message
    }

    $canStart = $taskFound -and $serviceStart -ne 'Disabled' -and $serviceState -ne 'NotFound'
    $status = 'Ok'
    if (-not $taskFound -or $serviceState -ne 'Running' -or $indexError) { $status = 'Warning' }
    if ($serviceState -eq 'NotFound') { $status = 'Error' }

    return [ordered]@{
        status          = $status
        enabled         = $true
        serviceState    = $serviceState
        serviceStartType = $serviceStart
        indexPath       = $indexPath
        indexBytes      = $indexBytes
        indexError      = $indexError
        taskPath        = ($taskPath + $taskName)
        taskFound       = $taskFound
        taskState       = $taskState
        lastRunUtc      = $lastRun
        lastTaskResult  = $lastResult
        canStart        = $canStart
        maintenanceStarted = $false
        taskError       = $taskError
        error           = $null
        note            = 'Optimize inicia IndexerAutomaticMaintenance. No reconstruye el indice.'
    }
}

function Start-SearchMaintenance {
    if ((Test-DryRun)) {
        throw 'Invariante rota: el mantenimiento del indice no se inicia en dry-run.'
    }
    $task = Get-ScheduledTask -TaskPath '\Microsoft\Windows\Shell\' -TaskName 'IndexerAutomaticMaintenance' -ErrorAction Stop
    Start-ScheduledTask -InputObject $task
}


function Confirm-Optimization {
    param(
        [int]$TempFilesToDelete,
        [int64]$TempBytes,
        [int]$ProcessesToTrim,
        [bool]$StartIndex
    )
    if ((Test-DryRun)) { return $false }
    if ($Force) {
        Write-Log -Level WARN -Event 'confirmation_forced' -Data @{
            tempFiles = $TempFilesToDelete
            tempBytes = $TempBytes
            processes = $ProcessesToTrim
            search    = $StartIndex
        }
        return $true
    }
    if (-not (Test-InteractiveSession)) {
        throw 'La consola no es interactiva. No se aplico nada. Repita con -Apply -Force para una tarea programada.'
    }
    Write-Host ''
    Write-Host 'Se aplicarian estas acciones:'
    if ($TempFilesToDelete -gt 0) {
        Write-Host ("- Borrar {0} temporales ({1}) con mas de {2} dias." -f $TempFilesToDelete, (Format-ByteSize -Bytes $TempBytes), $OlderThanDays)
    }
    if ($ProcessesToTrim -gt 0) {
        Write-Host ("- Recortar el conjunto de trabajo de {0} procesos." -f $ProcessesToTrim)
    }
    if ($StartIndex) {
        Write-Host '- Iniciar la tarea IndexerAutomaticMaintenance.'
    }
    $answer = Read-Host 'Escriba SI en mayusculas para continuar'
    $accepted = ($answer -eq 'SI')
    Write-Log -Level INFO -Event 'confirmation_answer' -Data @{ accepted = $accepted }
    return $accepted
}

function Invoke-OptimizerTask {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][scriptblock]$Action
    )
    try {
        $value = & $Action
        $status = Get-DictionaryValue -Object $value -Name 'status'
        Write-Log -Level INFO -Event 'task_completed' -Data @{
            task   = $Name
            status = [string]$status
        }
        return $value
    }
    catch {
        $message = $_.Exception.Message
        Write-Log -Level ERROR -Event 'task_failed' -Data @{
            task    = $Name
            message = $message
        }
        Write-Warning ("{0}: {1}" -f $Name, $message)
        return [ordered]@{
            status  = 'Error'
            enabled = $true
            error   = $message
        }
    }
}

function New-ProcessConsoleRows {
    param($Rows)
    $items = New-Object System.Collections.Generic.List[object]
    if ($null -eq $Rows) { return ,$items }
    foreach ($row in $Rows) {
        $cpuText = 'n/d'
        $cpuValue = Get-DictionaryValue -Object $row -Name 'cpuPercent'
        if ($null -ne $cpuValue) { $cpuText = ('{0:N1} %' -f [double]$cpuValue) }
        $bytes = Get-DictionaryValue -Object $row -Name 'workingSetBytes'
        if ($null -eq $bytes) { $bytes = 0 }
        $items.Add([pscustomobject]@{
            Proceso = (Get-DictionaryValue -Object $row -Name 'name')
            Pid     = (Get-DictionaryValue -Object $row -Name 'pid')
            CPU     = $cpuText
            RAM     = (Format-ByteSize -Bytes ([int64]$bytes))
        })
    }
    return ,$items
}

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }

    if ($WhatIfPreference) { $Apply = $false }
    Set-DryRunMode -DryRun (-not [bool]$Apply)

    $runAll = $All -or -not ($TempFiles -or $PendingUpdates -or $FreeMemory -or $ReviewProcesses -or $SearchIndex)
    $doTemp = $runAll -or [bool]$TempFiles
    $doUpdates = $runAll -or [bool]$PendingUpdates
    $doMemory = $runAll -or [bool]$FreeMemory
    $doProcesses = $runAll -or [bool]$ReviewProcesses
    $doSearch = $runAll -or [bool]$SearchIndex

    $repoRoot = Get-RepositoryRoot
    $stamp = [DateTime]::Now.ToString('yyyyMMdd-HHmmss')
    if ([string]::IsNullOrWhiteSpace($ReportPath)) {
        $ReportPath = Join-Path (Join-Path $repoRoot 'reports') ("DailyOptimizer-{0}.json" -f $stamp)
    }
    if ([string]::IsNullOrWhiteSpace($LogPath)) {
        $LogPath = Join-Path (Join-Path $repoRoot 'logs') 'DailyOptimizer.jsonl'
    }
    $ReportPath = [System.IO.Path]::GetFullPath($ReportPath)
    Initialize-CommonSession -ScriptName 'DailyOptimizer' -LogPath ([System.IO.Path]::GetFullPath($LogPath)) -DryRun (Test-DryRun)
    Ensure-Directory -Path (Split-Path -Parent $ReportPath)
    $script:Whitelist = Initialize-Whitelist

    $modeText = 'dry-run'
    if (-not (Test-DryRun)) { $modeText = 'apply' }

    Write-Host 'DailyOptimizer 1.0'
    Write-Host ("Modo: {0}" -f $modeText)
    Write-Host ("Tareas: temporales={0}; actualizaciones={1}; memoria={2}; procesos={3}; indice={4}" -f $doTemp, $doUpdates, $doMemory, $doProcesses, $doSearch)
    Write-Host ("Antiguedad de temporales: {0} dias. Umbral de RAM: {1} MB. Muestra de CPU: {2} s." -f $OlderThanDays, $MemoryThresholdMB, $CpuSampleSeconds)

    Write-Log -Level INFO -Event 'run_start' -Data @{
        dryRun          = [bool](Test-DryRun)
        tempFiles       = $doTemp
        pendingUpdates  = $doUpdates
        freeMemory      = $doMemory
        reviewProcesses = $doProcesses
        searchIndex     = $doSearch
        olderThanDays   = $OlderThanDays
        top             = $Top
        memoryThresholdMb = $MemoryThresholdMB
        reportPath      = $ReportPath
    }

    if ($doUpdates) {
        Write-Log -Level INFO -Event 'update_search_skipped' -Data @{
            reason          = 'Se reutiliza el snapshot de UpdateChecker.'
            updateCacheHours = $UpdateCacheHours
        }
    }

    $sample = $null
    if ($doMemory -or $doProcesses) {
        $sample = Invoke-OptimizerTask -Name 'processSample' -Action { Get-ProcessSample }
    }

    $tempReport = New-SkippedTask -Note 'Tarea desactivada.'
    if ($doTemp) {
        $tempReport = Invoke-OptimizerTask -Name 'tempFiles' -Action { Get-TempReport }
    }

    $searchReport = New-SkippedTask -Note 'Tarea desactivada.'
    if ($doSearch) {
        $searchReport = Invoke-OptimizerTask -Name 'searchIndex' -Action { Get-SearchReport }
    }

    $updateReport = New-SkippedTask -Note 'Tarea desactivada.'
    if ($doUpdates) {
        $updateReport = Invoke-OptimizerTask -Name 'pendingUpdates' -Action { Complete-PendingUpdateSearch }
    }

    $memoryReport = New-SkippedTask -Note 'Tarea desactivada.'
    $processReport = New-SkippedTask -Note 'Tarea desactivada.'
    $sampleStatus = Get-DictionaryValue -Object $sample -Name 'status'
    if ($doMemory) {
        if ($sampleStatus -eq 'Error') {
            $memoryReport = [ordered]@{ status = 'Error'; enabled = $true; error = (Get-DictionaryValue -Object $sample -Name 'error') }
        }
        else {
            $memoryReport = Invoke-OptimizerTask -Name 'freeMemory' -Action { Build-MemoryReport -Sample $sample }
        }
    }
    if ($doProcesses) {
        if ($sampleStatus -eq 'Error') {
            $processReport = [ordered]@{ status = 'Error'; enabled = $true; error = (Get-DictionaryValue -Object $sample -Name 'error') }
        }
        else {
            $processReport = Invoke-OptimizerTask -Name 'reviewProcesses' -Action { Build-ProcessReport -Sample $sample }
        }
    }

    $tempEligible = 0
    $tempBytes = [int64]0
    if ($doTemp -and (Get-DictionaryValue -Object $tempReport -Name 'status') -ne 'Error') {
        $countValue = Get-DictionaryValue -Object $tempReport -Name 'fileCount'
        $byteValue = Get-DictionaryValue -Object $tempReport -Name 'bytes'
        if ($null -ne $countValue) { $tempEligible = [int]$countValue }
        if ($null -ne $byteValue) { $tempBytes = [int64]$byteValue }
    }
    $trimCount = 0
    if ($doMemory -and (Get-DictionaryValue -Object $memoryReport -Name 'status') -ne 'Error') {
        $eligibleValue = Get-DictionaryValue -Object $memoryReport -Name 'eligibleCount'
        if ($null -ne $eligibleValue) { $trimCount = [int]$eligibleValue }
    }
    $canStartSearch = $false
    if ($doSearch -and (Get-DictionaryValue -Object $searchReport -Name 'status') -ne 'Error') {
        $canStartValue = Get-DictionaryValue -Object $searchReport -Name 'canStart'
        if ($canStartValue) { $canStartSearch = $true }
    }

    $needsApply = ($tempEligible -gt 0) -or ($trimCount -gt 0) -or $canStartSearch
    $confirmed = $false
    $cancelled = $false
    if (-not (Test-DryRun) -and $needsApply) {
        $confirmed = Confirm-Optimization -TempFilesToDelete $tempEligible -TempBytes $tempBytes -ProcessesToTrim $trimCount -StartIndex:$canStartSearch
        if (-not $confirmed) {
            $cancelled = $true
            Write-Log -Level INFO -Event 'apply_cancelled' -Data @{ reason = 'El usuario no escribio SI.' }
        }
    }
    elseif ((Test-DryRun)) {
        Write-Log -Level INFO -Event 'apply_skipped' -Data @{ reason = 'DryRun' }
    }

    $appliedActions = New-Object System.Collections.Generic.List[string]
    if ($confirmed) {
        if ($doTemp -and $tempEligible -gt 0) {
            $deletedFiles = 0
            $deletedBytes = [int64]0
            $deleteFailed = 0
            foreach ($root in (Get-DictionaryValue -Object $tempReport -Name 'roots')) {
                $accepted = Get-DictionaryValue -Object $root -Name 'accepted'
                $rootCount = Get-DictionaryValue -Object $root -Name 'fileCount'
                if (-not $accepted -or [int]$rootCount -le 0) { continue }
                $rootPath = [string](Get-DictionaryValue -Object $root -Name 'path')
                Write-Log -Level INFO -Event 'delete_started' -Data @{ path = $rootPath }
                $outcome = Remove-TempRootContents -Root $rootPath
                $deletedFiles += [int]$outcome.deleted
                $deletedBytes += [int64]$outcome.deletedBytes
                $deleteFailed += [int]$outcome.failed
            }
            $tempReport.deletedFiles = $deletedFiles
            $tempReport.deletedBytes = $deletedBytes
            $tempReport.deleteFailed = $deleteFailed
            $tempReport.applied = $true
            $appliedActions.Add(("Temporales borrados: {0} ({1}). Fallos: {2}." -f $deletedFiles, (Format-ByteSize -Bytes $deletedBytes), $deleteFailed))
            Write-Log -Level INFO -Event 'delete_finished' -Data @{
                deleted = $deletedFiles
                bytes   = $deletedBytes
                failed  = $deleteFailed
            }
        }
        if ($doMemory -and $trimCount -gt 0) {
            $eligibleRows = Get-DictionaryValue -Object $memoryReport -Name 'eligible'
            $trim = Invoke-WorkingSetTrim -Eligible $eligibleRows
            $memoryReport.trimmedCount = [int]$trim.trimmed
            $memoryReport.applied = $true
            $after = Get-MemorySnapshot
            $memoryReport.freeBytes = $after.freeBytes
            $memoryReport.usedBytes = $after.usedBytes
            $memoryReport.usedPercent = $after.usedPercent
            $appliedActions.Add(("Conjuntos de trabajo recortados: {0}. Fallos: {1}." -f [int]$trim.trimmed, [int]$trim.failed))
            Write-Log -Level INFO -Event 'memory_trim_finished' -Data @{
                trimmed = [int]$trim.trimmed
                failed  = [int]$trim.failed
            }
        }
        if ($doSearch -and $canStartSearch) {
            Start-SearchMaintenance
            $searchReport.maintenanceStarted = $true
            $appliedActions.Add('Mantenimiento del indice iniciado.')
            Write-Log -Level INFO -Event 'search_maintenance_started' -Data @{
                task = '\Microsoft\Windows\Shell\IndexerAutomaticMaintenance'
            }
        }
    }

    $statuses = New-Object System.Collections.Generic.List[string]
    foreach ($section in @($tempReport, $updateReport, $memoryReport, $processReport, $searchReport)) {
        $one = [string](Get-DictionaryValue -Object $section -Name 'status')
        if ($one -and $one -ne 'Skipped') { $statuses.Add($one) }
    }
    $overall = Get-WorstStatus -Statuses ($statuses.ToArray())

    $findings = New-Object System.Collections.Generic.List[string]
    if ((Test-DryRun)) {
        $findings.Add('Modo dry-run: no se borro, no se recorto memoria y no se inicio el mantenimiento del indice.')
    }
    elseif ($cancelled) {
        $findings.Add('Apply cancelado: no se escribio SI. No se modifico el equipo.')
    }
    elseif ((Get-ItemCount -Value $appliedActions) -eq 0) {
        $findings.Add('Apply activo, pero no habia acciones que ejecutar.')
    }
    else {
        foreach ($action in $appliedActions) { $findings.Add([string]$action) }
    }
    if ($doTemp) {
        $findings.Add(("Temporales con mas de {0} dias: {1} archivos ({2})." -f $OlderThanDays, $tempEligible, (Format-ByteSize -Bytes $tempBytes)))
    }
    if ($doUpdates) {
        $pendingText = Get-DictionaryValue -Object $updateReport -Name 'pendingCount'
        $rebootText = Get-DictionaryValue -Object $updateReport -Name 'rebootPending'
        $queryModeText = [string](Get-DictionaryValue -Object $updateReport -Name 'queryMode')
        $findings.Add(("Actualizaciones de Windows pendientes: {0}. Reinicio pendiente: {1}. Origen: {2}." -f $pendingText, $rebootText, $queryModeText))
        $updateError = Get-DictionaryValue -Object $updateReport -Name 'error'
        if ($updateError) { $findings.Add(("Busqueda de actualizaciones: {0}" -f $updateError)) }
    }
    if ($doMemory) {
        $usedText = Get-DictionaryValue -Object $memoryReport -Name 'usedPercent'
        $eligibleBytes = Get-DictionaryValue -Object $memoryReport -Name 'eligibleWorkingSetBytes'
        if ($null -eq $eligibleBytes) { $eligibleBytes = 0 }
        if ($null -ne $usedText) {
            $findings.Add(("Memoria en uso: {0:N2}%. Procesos elegibles para recorte: {1} ({2} de conjunto de trabajo)." -f [double]$usedText, $trimCount, (Format-ByteSize -Bytes ([int64]$eligibleBytes))))
        }
    }
    if ($doProcesses) {
        $heavyText = Get-DictionaryValue -Object $processReport -Name 'heavyCount'
        $findings.Add(("Procesos por encima del umbral (CPU >= 15% o RAM >= {0} MB): {1}." -f $MemoryThresholdMB, $heavyText))
    }
    if ($doSearch) {
        $searchState = Get-DictionaryValue -Object $searchReport -Name 'serviceState'
        $indexBytes = Get-DictionaryValue -Object $searchReport -Name 'indexBytes'
        $indexText = 'n/d'
        if ($null -ne $indexBytes) { $indexText = Format-ByteSize -Bytes ([int64]$indexBytes) }
        $findings.Add(("Windows Search: {0}. Windows.edb: {1}." -f $searchState, $indexText))
    }

    $stopwatch.Stop()
    $report = [ordered]@{
        schemaVersion = '1.0'
        generatedAt   = (Get-UtcTimestamp)
        hostname      = [string]$env:COMPUTERNAME
        userName      = [string]$env:USERNAME
        runId         = (Get-RunId)
        dryRun        = [bool](Test-DryRun)
        parameters    = [ordered]@{
            olderThanDays       = $OlderThanDays
            top                 = $Top
            memoryThresholdMb   = $MemoryThresholdMB
            cpuSampleSeconds    = $CpuSampleSeconds
            updateTimeoutSeconds = $UpdateTimeoutSeconds
            tempFiles           = $doTemp
            pendingUpdates      = $doUpdates
            freeMemory          = $doMemory
            reviewProcesses     = $doProcesses
            searchIndex         = $doSearch
        }
        tasks         = [ordered]@{
            tempFiles        = $tempReport
            pendingUpdates   = $updateReport
            freeMemory       = $memoryReport
            reviewProcesses  = $processReport
            searchIndex      = $searchReport
        }
        summary       = [ordered]@{
            overallStatus  = $overall
            cancelled      = $cancelled
            actionsApplied = $appliedActions
            findings       = $findings
            durationMs     = [int64]$stopwatch.ElapsedMilliseconds
        }
    }

    $json = Write-JsonReport -Path $ReportPath -Report $report
    Write-Log -Level INFO -Event 'report_written' -Data @{ path = $ReportPath }

    $taskRows = New-Object System.Collections.Generic.List[object]
    $taskRows.Add([pscustomobject]@{ Tarea = 'Temporales'; Estado = [string](Get-DictionaryValue -Object $tempReport -Name 'status'); Detalle = ("{0} archivos, {1}" -f $tempEligible, (Format-ByteSize -Bytes $tempBytes)) })
    $pendingShown = Get-DictionaryValue -Object $updateReport -Name 'pendingCount'
    $taskRows.Add([pscustomobject]@{ Tarea = 'Actualizaciones'; Estado = [string](Get-DictionaryValue -Object $updateReport -Name 'status'); Detalle = ("{0} pendientes" -f $pendingShown) })
    $usedShown = Get-DictionaryValue -Object $memoryReport -Name 'usedPercent'
    $memoryDetail = 'omitida'
    if ($null -ne $usedShown) { $memoryDetail = ("{0:N2}% en uso, {1} elegibles" -f [double]$usedShown, $trimCount) }
    $taskRows.Add([pscustomobject]@{ Tarea = 'Memoria'; Estado = [string](Get-DictionaryValue -Object $memoryReport -Name 'status'); Detalle = $memoryDetail })
    $heavyShown = Get-DictionaryValue -Object $processReport -Name 'heavyCount'
    $taskRows.Add([pscustomobject]@{ Tarea = 'Procesos'; Estado = [string](Get-DictionaryValue -Object $processReport -Name 'status'); Detalle = ("{0} sobre el umbral" -f $heavyShown) })
    $searchShown = Get-DictionaryValue -Object $searchReport -Name 'serviceState'
    $taskRows.Add([pscustomobject]@{ Tarea = 'Indice'; Estado = [string](Get-DictionaryValue -Object $searchReport -Name 'status'); Detalle = [string]$searchShown })
    Write-ConsoleTable -Title 'Resumen de tareas' -Rows $taskRows

    if ($doTemp) {
        $fileRows = New-Object System.Collections.Generic.List[object]
        foreach ($file in (Get-DictionaryValue -Object $tempReport -Name 'largest')) {
            $fileBytes = Get-DictionaryValue -Object $file -Name 'bytes'
            if ($null -eq $fileBytes) { $fileBytes = 0 }
            $fileRows.Add([pscustomobject]@{
                Bytes = (Format-ByteSize -Bytes ([int64]$fileBytes))
                Ruta  = (Get-DictionaryValue -Object $file -Name 'path')
            })
        }
        Write-ConsoleTable -Title 'Temporales mas grandes elegibles' -Rows $fileRows
    }
    if ($doProcesses) {
        Write-ConsoleTable -Title 'Procesos por CPU' -Rows (New-ProcessConsoleRows -Rows (Get-DictionaryValue -Object $processReport -Name 'topCpu'))
        Write-ConsoleTable -Title 'Procesos por RAM' -Rows (New-ProcessConsoleRows -Rows (Get-DictionaryValue -Object $processReport -Name 'topMemory'))
    }

    Write-Host ''
    Write-Host ("Estado global: {0}" -f $overall)
    foreach ($finding in $findings) {
        Write-Host ("- {0}" -f $finding)
    }
    Write-Host ("Reporte JSON: {0}" -f $ReportPath)
    Write-Host ("Duracion: {0:N1} s" -f ($stopwatch.ElapsedMilliseconds / 1000.0))

    Write-Log -Level INFO -Event 'run_completed' -Data @{
        overallStatus = $overall
        dryRun        = [bool](Test-DryRun)
        durationMs    = [int64]$stopwatch.ElapsedMilliseconds
        reportPath    = $ReportPath
    }
    $script:ExitCode = Get-ReportExitCode -Status $overall
}
catch {
    Write-ScriptError -ErrorRecord $_
    $script:ExitCode = (Get-FatalExitCode)
}
finally {
    Clear-PendingUpdateJob
}

Write-Host ("Codigo de salida: {0}" -f $script:ExitCode)
if ($MyInvocation.InvocationName -ne '.') {
    exit $script:ExitCode
}
