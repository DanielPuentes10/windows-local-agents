#Requires -Version 5.1
<#
.SYNOPSIS
    Analiza el uso de disco, las carpetas mas grandes y los archivos temporales.

.DESCRIPTION
    Script de solo lectura. Resume los volumenes fijos con Win32_LogicalDisk,
    mide las carpetas de primer nivel de -ScanRoot con robocopy en modo lista
    (/L, no copia ni borra) y inventaria las ubicaciones temporales estandar.

    En modo normal escribe un reporte JSON en reports\ y anade eventos JSONL
    en logs\. Con -DryRun hace el mismo analisis y el mismo log, pero no crea
    ni sobrescribe el archivo de reporte: el JSON sale por la canalizacion.

    No elimina archivos, no vacia temporales y no requiere administrador.
    Las rutas inaccesibles quedan registradas y los tamanos pueden ser menores
    que el uso real. Los puntos de union (junctions) se omiten para no contar
    dos veces la misma carpeta.

.PARAMETER Drive
    Unidades fijas a resumir, por ejemplo C: o D. Si se omite, se usan todas
    las unidades locales fijas. No cambia la raiz del escaneo de carpetas.

.PARAMETER ScanRoot
    Carpeta cuyos directorios inmediatos se miden. El valor predeterminado es
    la raiz de la unidad del sistema, normalmente C:\.

.PARAMETER Top
    Cantidad de carpetas mas grandes que se incluyen en el reporte. De 1 a 100.
    El valor predeterminado es 15.

.PARAMETER MaxDepth
    Profundidad maxima de medicion. 0 significa arbol completo. Un valor mayor
    que 0 aplica robocopy /LEV y marca el tamano como parcial.

.PARAMETER TopTempFiles
    Cantidad de archivos temporales mas grandes que se listan. De 1 a 100.
    El valor predeterminado es 20.

.PARAMETER ReportPath
    Ruta del JSON de salida. Si se omite, se genera
    reports\DiskHealth-yyyyMMdd-HHmmss.json junto al repositorio del script.

.PARAMETER LogPath
    Ruta del log JSONL. Si se omite, se usa logs\DiskHealthAnalyzer.jsonl.
    El log se escribe tambien en -DryRun.

.PARAMETER DryRun
    Analiza y registra, pero no escribe el archivo de reporte.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\DiskHealthAnalyzer.ps1 -DryRun

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\DiskHealthAnalyzer.ps1 -ScanRoot "$env:USERPROFILE" -Top 10

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\DiskHealthAnalyzer.ps1 -Drive C: -ReportPath .\reports\disco.json

.NOTES
    Autor: LocalAgents
    Requiere: Windows PowerShell 5.1 o superior y robocopy.exe
    Codigos de salida (Common.psm1):
      0  el reporte termino con estado Ok
      1  el reporte termino con estado Warning
      2  el reporte termino con estado Critical o Error
      3  el script fallo antes de cerrar el reporte
    Helpers compartidos: scripts\Common.psm1.
    Umbrales: aviso >= 80% de uso, critico >= 90%.

    Programador de tareas (reporte real, una vez al dia a las 08:00):
      $script = 'C:\ProyectosDeSoftware\LocalAgents\scripts\DiskHealthAnalyzer.ps1'
      $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$script`""
      $trigger = New-ScheduledTaskTrigger -Daily -At 8:00am
      Register-ScheduledTask -TaskName 'LocalAgents-DiskHealth' -Action $action -Trigger $trigger -Description 'Reporte de salud de disco'
#>
[CmdletBinding()]
param(
    [string[]]$Drive,

    [ValidateNotNullOrEmpty()]
    [string]$ScanRoot,

    [ValidateRange(1, 100)]
    [int]$Top = 15,

    [ValidateRange(0, 64)]
    [int]$MaxDepth = 0,

    [ValidateRange(1, 100)]
    [int]$TopTempFiles = 20,

    [string]$ReportPath,

    [string]$LogPath,

    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -Force -DisableNameChecking

$script:WarningUsedPercent = 80
$script:CriticalUsedPercent = 90
$script:TempAttentionBytes = 1GB
$script:ExitCode = 0


function Get-NormalizedDirectory {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        throw "La ruta no existe o no es una carpeta: $Path"
    }

    $full = [System.IO.Path]::GetFullPath($Path)
    if ($full.Length -gt 3) {
        $full = $full.TrimEnd('\')
    }
    return $full
}

function Resolve-DriveId {
    param([Parameter(Mandatory = $true)][string]$Value)

    $text = $Value.Trim()
    if ($text -match '^[A-Za-z]$') {
        return ($text.ToUpperInvariant() + ':')
    }
    if ($text -match '^[A-Za-z]:\\?$') {
        return ($text.Substring(0, 1).ToUpperInvariant() + ':')
    }

    throw "Unidad no valida: $Value. Use el formato C: o C."
}

function Get-VolumeStatus {
    param([double]$UsedPercent)

    if ($UsedPercent -ge $script:CriticalUsedPercent) { return 'Critical' }
    if ($UsedPercent -ge $script:WarningUsedPercent) { return 'Warning' }
    return 'Ok'
}


function ConvertTo-Int64Loose {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $digits = ($Text -replace '[^\d]', '')
    if ([string]::IsNullOrWhiteSpace($digits)) { return $null }
    return [int64]$digits
}

function Get-RobocopySource {
    param([Parameter(Mandatory = $true)][string]$Path)

    if ($Path.StartsWith('\\?\')) { return $Path }
    if ($Path.Length -lt 240) { return $Path }
    if ($Path.StartsWith('\\')) {
        return ('\\?\UNC\' + $Path.Substring(2))
    }
    return ('\\?\' + $Path)
}

function Invoke-RobocopyMeasure {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$Depth = 0
    )

    # /L lista sin copiar. NULL no se materializa como destino.
    $robocopyArgs = @(
        (Get-RobocopySource -Path $Path),
        'NULL',
        '/L',
        '/S',
        '/BYTES',
        '/XJ',
        '/R:0',
        '/W:0',
        '/NJH',
        '/NDL',
        '/NC',
        '/NFL',
        '/NP',
        '/MT:8'
    )

    if ($Depth -gt 0) {
        $robocopyArgs += ('/LEV:{0}' -f $Depth)
    }

    if ($robocopyArgs -notcontains '/L') {
        throw 'Invariante rota: robocopy debe ejecutarse siempre con /L.'
    }

    $started = [System.Diagnostics.Stopwatch]::StartNew()
    $output = & $script:RobocopyPath @robocopyArgs 2>&1
    $exitCode = $LASTEXITCODE
    $started.Stop()

    $text = (($output | ForEach-Object { "$_" }) -join "`n")
    $bytes = [int64]0
    $files = [int64]0
    $directoryErrors = [int64]0
    $parsedBytes = $false

    foreach ($line in ($text -split '\r?\n')) {
        if ($line -match '^\s*Bytes\s*:\s+(\S+)') {
            $parsedValue = ConvertTo-Int64Loose -Text $Matches[1]
            if ($null -ne $parsedValue) {
                $bytes = $parsedValue
                $parsedBytes = $true
            }
            continue
        }

        if ($line -match '^\s*(?:Archivos|Files)\s*:\s+(\S+)') {
            $parsedFiles = ConvertTo-Int64Loose -Text $Matches[1]
            if ($null -ne $parsedFiles) { $files = $parsedFiles }
            continue
        }

        if ($line -match '^\s*(?:Director\.|Dirs)\s*:') {
            $columns = @($line.Split(':', 2)[1].Trim() -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($columns.Count -ge 5) {
                $parsedErrors = ConvertTo-Int64Loose -Text $columns[4]
                if ($null -ne $parsedErrors) { $directoryErrors = $parsedErrors }
            }
        }
    }

    $hardFailure = ($exitCode -ge 16) -or (-not $parsedBytes)
    $partial = ($Depth -gt 0) -or ($directoryErrors -gt 0) -or ($exitCode -ge 8)
    $errorText = $null
    if ($hardFailure) {
        $snippet = $text.Trim()
        if ($snippet.Length -gt 500) { $snippet = $snippet.Substring(0, 500) }
        $errorText = ("robocopy termino con codigo {0}. {1}" -f $exitCode, $snippet).Trim()
        $errorText = ($errorText -replace '\s+', ' ')
    }

    return ,([ordered]@{
        ok               = (-not $hardFailure)
        partial          = [bool]($partial -or $hardFailure)
        sizeIsComplete   = (-not $hardFailure -and -not $partial)
        exitCode         = $exitCode
        sizeBytes        = $(if ($hardFailure) { $null } else { $bytes })
        fileCount        = $(if ($hardFailure) { $null } else { $files })
        accessErrors     = $directoryErrors
        elapsedMs        = [int]$started.ElapsedMilliseconds
        error            = $errorText
    })
}

function Get-FixedVolumes {
    param([string[]]$DriveFilter)

    $disks = @(Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3')
    if ($DriveFilter -and $DriveFilter.Count -gt 0) {
        $wanted = @($DriveFilter | ForEach-Object { Resolve-DriveId -Value $_ })
        $available = @($disks | ForEach-Object { $_.DeviceID })
        foreach ($id in $wanted) {
            if ($available -notcontains $id) {
                throw ("La unidad {0} no esta disponible. Unidades fijas: {1}" -f $id, ($available -join ', '))
            }
        }
        $disks = @($disks | Where-Object { $wanted -contains $_.DeviceID })
    }

    $items = New-Object System.Collections.Generic.List[object]
    foreach ($disk in $disks) {
        $sizeBytes = [int64]$disk.Size
        $freeBytes = [int64]$disk.FreeSpace
        $usedBytes = $sizeBytes - $freeBytes
        $usedPercent = 0.0
        if ($sizeBytes -gt 0) {
            $usedPercent = [math]::Round((100.0 * $usedBytes) / $sizeBytes, 2)
        }

        $items.Add([ordered]@{
            deviceId    = [string]$disk.DeviceID
            volumeName  = [string]$disk.VolumeName
            fileSystem  = [string]$disk.FileSystem
            sizeBytes   = $sizeBytes
            freeBytes   = $freeBytes
            usedBytes   = $usedBytes
            usedPercent = $usedPercent
            status      = (Get-VolumeStatus -UsedPercent $usedPercent)
        })
    }

    Write-Output -NoEnumerate $items.ToArray()
}

function Get-ChildDirectories {
    param([Parameter(Mandatory = $true)][string]$Root)

    $measured = New-Object System.Collections.Generic.List[object]
    $junctions = New-Object System.Collections.Generic.List[object]
    $entries = @(Get-ChildItem -LiteralPath $Root -Force -Directory -ErrorAction SilentlyContinue)

    foreach ($entry in $entries) {
        $isJunction = (($entry.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
        if ($isJunction) {
            $junctions.Add([ordered]@{
                name = [string]$entry.Name
                path = [string]$entry.FullName
            })
        }
        else {
            $measured.Add($entry)
        }
    }

    return ,([ordered]@{
        directories = $measured
        junctions   = $junctions
    })
}

function Get-DirectFiles {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][int]$Limit
    )

    $files = New-Object System.Collections.Generic.List[object]
    $accessErrors = 0
    $totalBytes = [int64]0
    $totalCount = 0

    try {
        $directory = New-Object System.IO.DirectoryInfo $Root
        foreach ($file in $directory.EnumerateFiles()) {
            try {
                $length = [int64]$file.Length
                $totalBytes += $length
                $totalCount++
                $files.Add([ordered]@{
                    path             = [string]$file.FullName
                    sizeBytes        = $length
                    lastWriteTimeUtc = ($file.LastWriteTimeUtc.ToString('yyyy-MM-ddTHH:mm:ss.fff', [System.Globalization.CultureInfo]::InvariantCulture) + 'Z')
                })
            }
            catch {
                $accessErrors++
            }
        }
    }
    catch {
        return ,([ordered]@{
            accessible   = $false
            totalBytes   = $null
            totalCount   = $null
            accessErrors = 1
            largest      = (New-Object System.Collections.Generic.List[object])
            error        = $_.Exception.Message
        })
    }

    $largest = New-Object System.Collections.Generic.List[object]
    foreach ($item in @($files | Sort-Object { $_.sizeBytes } -Descending | Select-Object -First $Limit)) {
        $largest.Add($item)
    }

    return ,([ordered]@{
        accessible   = $true
        totalBytes   = $totalBytes
        totalCount   = $totalCount
        accessErrors = $accessErrors
        largest      = $largest
        error        = $null
    })
}

function Get-LargestFiles {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][int]$Limit
    )

    $best = New-Object System.Collections.Generic.List[object]
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($Root)
    $enumerated = 0
    $accessErrors = 0

    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        try {
            $directory = New-Object System.IO.DirectoryInfo $current
            foreach ($file in $directory.EnumerateFiles()) {
                try {
                    $enumerated++
                    $candidate = [ordered]@{
                        path             = [string]$file.FullName
                        sizeBytes        = [int64]$file.Length
                        lastWriteTimeUtc = ($file.LastWriteTimeUtc.ToString('yyyy-MM-ddTHH:mm:ss.fff', [System.Globalization.CultureInfo]::InvariantCulture) + 'Z')
                    }
                    if ($best.Count -lt $Limit) {
                        $best.Add($candidate)
                        continue
                    }

                    $minimumIndex = 0
                    $minimumSize = [int64]$best[0].sizeBytes
                    for ($index = 1; $index -lt $best.Count; $index++) {
                        $size = [int64]$best[$index].sizeBytes
                        if ($size -lt $minimumSize) {
                            $minimumSize = $size
                            $minimumIndex = $index
                        }
                    }
                    if ([int64]$candidate.sizeBytes -gt $minimumSize) {
                        $best[$minimumIndex] = $candidate
                    }
                }
                catch {
                    $accessErrors++
                }
            }

            foreach ($child in $directory.EnumerateDirectories()) {
                try {
                    if (($child.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
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
    foreach ($item in @($best | Sort-Object { $_.sizeBytes } -Descending)) {
        $ordered.Add($item)
    }

    return ,([ordered]@{
        enumerated   = $enumerated
        accessErrors = $accessErrors
        largest      = $ordered
    })
}

function Get-TempLocations {
    $candidates = @(
        $env:TEMP,
        $(if ($env:LOCALAPPDATA) { Join-Path $env:LOCALAPPDATA 'Temp' }),
        $(if ($env:WINDIR) { Join-Path $env:WINDIR 'Temp' })
    )

    $seen = @{}
    $locations = New-Object System.Collections.Generic.List[string]
    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        try {
            $full = [System.IO.Path]::GetFullPath($candidate)
        }
        catch {
            continue
        }
        if ($full.Length -gt 3) { $full = $full.TrimEnd('\') }
        $key = $full.ToUpperInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $locations.Add($full)
    }

    # Un arreglo normal evita que @() reciba List[T] y falle en Windows PowerShell 5.1.
    Write-Output -NoEnumerate $locations.ToArray()
}


$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {

    try {
        [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
    }
    catch {
        # La consola del host puede no permitir cambiar la codificacion.
    }

    $repoRoot = Get-RepositoryRoot
    $stamp = [DateTime]::Now.ToString('yyyyMMdd-HHmmss')
    if ([string]::IsNullOrWhiteSpace($ReportPath)) {
        $ReportPath = Join-Path (Join-Path $repoRoot 'reports') ("DiskHealth-{0}.json" -f $stamp)
    }
    if ([string]::IsNullOrWhiteSpace($LogPath)) {
        $LogPath = Join-Path (Join-Path $repoRoot 'logs') 'DiskHealthAnalyzer.jsonl'
    }

    $ReportPath = [System.IO.Path]::GetFullPath($ReportPath)
    Initialize-CommonSession -ScriptName 'DiskHealthAnalyzer' -LogPath ([System.IO.Path]::GetFullPath($LogPath)) -DryRun ([bool]$DryRun)

    $robocopyCommand = Get-Command -Name robocopy.exe -CommandType Application -ErrorAction SilentlyContinue
    if (-not $robocopyCommand) {
        throw 'robocopy.exe no esta disponible en PATH.'
    }
    $script:RobocopyPath = $robocopyCommand.Source

    if ([string]::IsNullOrWhiteSpace($ScanRoot)) {
        $systemDrive = $env:SystemDrive
        if ([string]::IsNullOrWhiteSpace($systemDrive)) { $systemDrive = 'C:' }
        if ($systemDrive -match '^[A-Za-z]:$') { $systemDrive = $systemDrive + '\' }
        $ScanRoot = $systemDrive
    }
    $ScanRoot = Get-NormalizedDirectory -Path $ScanRoot

    $modeLabel = 'NORMAL'
    if ((Test-DryRun)) { $modeLabel = 'DRY-RUN' }

    Write-Host 'DiskHealthAnalyzer 1.0'
    Write-Host ("Modo: {0}" -f $modeLabel)
    Write-Host 'Este script no modifica archivos de usuario ni elimina temporales.'
    if ((Test-DryRun)) {
        Write-Host 'Dry-run: el reporte JSON no se escribira en disco. El log estructurado si se actualiza.'
    }
    else {
        Write-Host ("El reporte JSON se escribira en: {0}" -f $ReportPath)
    }
    Write-Host ("Log: {0}" -f (Get-SessionLogPath))
    Write-Host ("Raiz de carpetas: {0}" -f $ScanRoot)

    Write-Log -Level INFO -Event 'run_start' -Data @{
        scanRoot     = $ScanRoot
        top          = $Top
        maxDepth     = $MaxDepth
        topTempFiles = $TopTempFiles
        reportPath   = $ReportPath
        driveFilter  = $(if ($Drive) { ($Drive -join ',') } else { '' })
    }

    Write-Host ''
    Write-Host 'Leyendo volumenes...'
    $volumes = Get-FixedVolumes -DriveFilter $Drive
    if ($null -eq $volumes) { $volumes = @() }
    Write-Log -Level INFO -Event 'volumes_collected' -Data @{
        count = $volumes.Count
    }

    $volumeRows = New-Object System.Collections.Generic.List[object]
    foreach ($volume in $volumes) {
        $volumeRows.Add([pscustomobject]@{
            Unidad = $volume.deviceId
            Nombre = $volume.volumeName
            FS     = $volume.fileSystem
            Tamano = (Format-ByteSize -Bytes $volume.sizeBytes)
            Libre  = (Format-ByteSize -Bytes $volume.freeBytes)
            Uso    = ('{0:N2} %' -f $volume.usedPercent)
            Estado = $volume.status
        })
    }
    Write-ConsoleTable -Title 'Volumenes fijos' -Rows $volumeRows

    $volumeById = @{}
    foreach ($volume in $volumes) {
        $volumeById[$volume.deviceId] = $volume
    }

    $scanDrive = [System.IO.Path]::GetPathRoot($ScanRoot).TrimEnd('\')
    $scanVolume = $null
    if ($volumeById.ContainsKey($scanDrive)) {
        $scanVolume = $volumeById[$scanDrive]
    }

    Write-Host ''
    Write-Host 'Midiendo carpetas de primer nivel (robocopy /L)...'
    Write-Log -Level INFO -Event 'folder_scan_started' -Data @{
        scanRoot = $ScanRoot
        maxDepth = $MaxDepth
    }

    $childInfo = Get-ChildDirectories -Root $ScanRoot
    $folderResults = New-Object System.Collections.Generic.List[object]
    $folderIndex = 0
    $folderTotal = $childInfo.directories.Count

    foreach ($directory in $childInfo.directories) {
        $folderIndex++
        Write-Host ("[{0}/{1}] {2}" -f $folderIndex, $folderTotal, $directory.FullName)
        $measure = Invoke-RobocopyMeasure -Path $directory.FullName -Depth $MaxDepth
        $percentOfVolume = $null
        if ($measure.ok -and $scanVolume -and $scanVolume.sizeBytes -gt 0) {
            $percentOfVolume = [math]::Round((100.0 * [int64]$measure.sizeBytes) / [int64]$scanVolume.sizeBytes, 2)
        }

        $record = [ordered]@{
            name             = [string]$directory.Name
            path             = [string]$directory.FullName
            sizeBytes        = $measure.sizeBytes
            fileCount        = $measure.fileCount
            percentOfVolume  = $percentOfVolume
            sizeIsComplete   = [bool]$measure.sizeIsComplete
            accessErrors     = $measure.accessErrors
            elapsedMs        = $measure.elapsedMs
            robocopyExitCode = $measure.exitCode
            error            = $measure.error
        }
        $folderResults.Add($record)

        $logLevel = 'INFO'
        if (-not $measure.ok) { $logLevel = 'WARN' }
        Write-Log -Level $logLevel -Event 'folder_measured' -Data @{
            path         = [string]$directory.FullName
            sizeBytes    = $(if ($null -eq $measure.sizeBytes) { -1 } else { [int64]$measure.sizeBytes })
            fileCount    = $(if ($null -eq $measure.fileCount) { -1 } else { [int64]$measure.fileCount })
            sizeComplete = [bool]$measure.sizeIsComplete
            exitCode     = $measure.exitCode
            elapsedMs    = $measure.elapsedMs
        }
    }

    $successfulFolders = @($folderResults | Where-Object { $null -ne $_.sizeBytes } | Sort-Object { $_.sizeBytes } -Descending)
    $failedFolders = @($folderResults | Where-Object { $null -eq $_.sizeBytes })
    $failedFolderList = New-Object System.Collections.Generic.List[object]
    foreach ($folder in $failedFolders) {
        if ($null -ne $folder) { $failedFolderList.Add($folder) }
    }
    $topFolders = New-Object System.Collections.Generic.List[object]
    foreach ($folder in @($successfulFolders | Select-Object -First $Top)) {
        $topFolders.Add($folder)
    }

    $measuredFolderBytes = [int64]0
    foreach ($folder in $successfulFolders) {
        $measuredFolderBytes += [int64]$folder.sizeBytes
    }

    $folderRows = New-Object System.Collections.Generic.List[object]
    $rank = 0
    foreach ($folder in $topFolders) {
        $rank++
        $folderRows.Add([pscustomobject]@{
            '#'        = $rank
            Tamano     = (Format-ByteSize -Bytes ([int64]$folder.sizeBytes))
            Archivos   = $folder.fileCount
            Completo   = $folder.sizeIsComplete
            Porcentaje = $(if ($null -eq $folder.percentOfVolume) { '' } else { '{0:N2} %' -f $folder.percentOfVolume })
            Carpeta    = $folder.path
        })
    }
    Write-ConsoleTable -Title ("Carpetas mas grandes (top {0})" -f $Top) -Rows $folderRows

    if ($failedFolders.Count -gt 0) {
        $failedRows = New-Object System.Collections.Generic.List[object]
        foreach ($folder in $failedFolders) {
            $failedRows.Add([pscustomobject]@{
                Carpeta = $folder.path
                Codigo  = $folder.robocopyExitCode
                Error   = $folder.error
            })
        }
        Write-ConsoleTable -Title 'Carpetas no medidas' -Rows $failedRows
    }

    Write-Host ''
    Write-Host 'Leyendo archivos sueltos en la raiz del escaneo...'
    $rootFiles = Get-DirectFiles -Root $ScanRoot -Limit $Top
    $rootFileBytes = [int64]0
    if ($rootFiles.accessible -and $null -ne $rootFiles.totalBytes) {
        $rootFileBytes = [int64]$rootFiles.totalBytes
    }
    $rootRows = New-Object System.Collections.Generic.List[object]
    foreach ($file in $rootFiles.largest) {
        $rootRows.Add([pscustomobject]@{
            Tamano  = (Format-ByteSize -Bytes ([int64]$file.sizeBytes))
            Archivo = $file.path
        })
    }
    Write-ConsoleTable -Title 'Archivos mas grandes en la raiz (no recursivo)' -Rows $rootRows

    Write-Host ''
    Write-Host 'Inventariando archivos temporales...'
    $tempLocations = Get-TempLocations
    if ($null -eq $tempLocations) { $tempLocations = @() }
    $tempResults = New-Object System.Collections.Generic.List[object]
    $tempFilePool = New-Object System.Collections.Generic.List[object]
    $tempBytes = [int64]0
    $tempCount = [int64]0

    foreach ($location in $tempLocations) {
        $locationRecord = [ordered]@{
            path           = $location
            exists         = $false
            accessible     = $false
            sizeBytes      = $null
            fileCount      = $null
            sizeIsComplete = $false
            accessErrors   = 0
            error          = $null
        }

        if (-not (Test-Path -LiteralPath $location -PathType Container)) {
            $locationRecord.error = 'La ruta no existe.'
            $tempResults.Add($locationRecord)
            Write-Log -Level WARN -Event 'temp_location_missing' -Data @{ path = $location }
            continue
        }

        $locationRecord.exists = $true
        $measure = Invoke-RobocopyMeasure -Path $location -Depth 0
        $locationRecord.accessible = [bool]$measure.ok
        $locationRecord.sizeBytes = $measure.sizeBytes
        $locationRecord.fileCount = $measure.fileCount
        $locationRecord.sizeIsComplete = [bool]$measure.sizeIsComplete
        $locationRecord.accessErrors = $measure.accessErrors
        $locationRecord.error = $measure.error
        $tempResults.Add($locationRecord)

        if ($measure.ok) {
            $tempBytes += [int64]$measure.sizeBytes
            $tempCount += [int64]$measure.fileCount
            $largest = Get-LargestFiles -Root $location -Limit $TopTempFiles
            foreach ($file in $largest.largest) {
                $tempFilePool.Add([ordered]@{
                    path             = $file.path
                    sizeBytes        = $file.sizeBytes
                    lastWriteTimeUtc = $file.lastWriteTimeUtc
                    location         = $location
                })
            }
            if ($largest.accessErrors -gt 0) {
                $locationRecord.accessErrors = [int64]$locationRecord.accessErrors + [int64]$largest.accessErrors
            }
        }

        $level = 'INFO'
        if (-not $measure.ok) { $level = 'WARN' }
        Write-Log -Level $level -Event 'temp_location_measured' -Data @{
            path      = $location
            sizeBytes = $(if ($null -eq $measure.sizeBytes) { -1 } else { [int64]$measure.sizeBytes })
            fileCount = $(if ($null -eq $measure.fileCount) { -1 } else { [int64]$measure.fileCount })
            exitCode  = $measure.exitCode
        }
    }

    $largestTempFiles = New-Object System.Collections.Generic.List[object]
    foreach ($file in @($tempFilePool | Sort-Object { $_.sizeBytes } -Descending | Select-Object -First $TopTempFiles)) {
        $largestTempFiles.Add($file)
    }

    $tempRows = New-Object System.Collections.Generic.List[object]
    foreach ($location in $tempResults) {
        $sizeLabel = ''
        if ($null -ne $location.sizeBytes) { $sizeLabel = Format-ByteSize -Bytes ([int64]$location.sizeBytes) }
        $tempRows.Add([pscustomobject]@{
            Ruta       = $location.path
            Existe     = $location.exists
            Accesible  = $location.accessible
            Archivos   = $location.fileCount
            Tamano     = $sizeLabel
            Completo   = $location.sizeIsComplete
        })
    }
    Write-ConsoleTable -Title 'Ubicaciones temporales' -Rows $tempRows

    $tempFileRows = New-Object System.Collections.Generic.List[object]
    foreach ($file in $largestTempFiles) {
        $tempFileRows.Add([pscustomobject]@{
            Tamano  = (Format-ByteSize -Bytes ([int64]$file.sizeBytes))
            Archivo = $file.path
        })
    }
    Write-ConsoleTable -Title ("Temporales mas grandes (top {0})" -f $TopTempFiles) -Rows $tempFileRows

    $accountedBytes = $measuredFolderBytes + $rootFileBytes
    $unaccountedBytes = $null
    $scanIsDriveRoot = ($ScanRoot -match '^[A-Za-z]:\\$')
    if ($scanIsDriveRoot -and $scanVolume) {
        $unaccountedBytes = [int64]$scanVolume.usedBytes - $accountedBytes
    }

    $recommendations = New-Object System.Collections.Generic.List[string]
    foreach ($volume in $volumes) {
        if ($volume.status -eq 'Critical') {
            $recommendations.Add(("El volumen {0} esta al {1}% de uso. Supera el umbral critico de {2}%." -f $volume.deviceId, $volume.usedPercent, $script:CriticalUsedPercent))
        }
        elseif ($volume.status -eq 'Warning') {
            $recommendations.Add(("El volumen {0} esta al {1}% de uso. Supera el umbral de aviso de {2}%." -f $volume.deviceId, $volume.usedPercent, $script:WarningUsedPercent))
        }
    }

    if ($tempBytes -ge $script:TempAttentionBytes) {
        $recommendations.Add(("Las ubicaciones temporales medidas suman {0}. Este script no las vacia." -f (Format-ByteSize -Bytes $tempBytes)))
    }

    $largestFolder = $null
    if ($topFolders.Count -gt 0) { $largestFolder = $topFolders[0] }
    if ($largestFolder -and [int64]$largestFolder.sizeBytes -ge 5GB) {
        $recommendations.Add(("La carpeta medida mas grande es {0} ({1})." -f $largestFolder.path, (Format-ByteSize -Bytes ([int64]$largestFolder.sizeBytes))))
    }

    if ($failedFolders.Count -gt 0) {
        $recommendations.Add(("{0} carpeta(s) no pudieron medirse por permisos u otro error de lectura." -f $failedFolders.Count))
    }

    $recommendations.Add('Analisis de solo lectura completado. No se modifico ningun archivo de usuario.')
    if ((Test-DryRun)) {
        $recommendations.Add('Dry-run: el reporte JSON no se escribio en disco.')
    }

    $driveParameter = New-Object System.Collections.Generic.List[string]
    if ($Drive) {
        foreach ($driveId in $Drive) { $driveParameter.Add([string]$driveId) }
    }

    $statuses = @($volumes | ForEach-Object { [string]$_.status })
    $overallStatus = Get-WorstStatus -Statuses $statuses
    $stopwatch.Stop()

    $osCaption = [System.Environment]::OSVersion.VersionString
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        if ($os.Caption) { $osCaption = [string]$os.Caption }
    }
    catch {
        Write-Log -Level WARN -Event 'os_caption_unavailable' -Data @{
            message = $_.Exception.Message
        }
    }

    $junctionList = New-Object System.Collections.Generic.List[object]
    foreach ($junction in $childInfo.junctions) {
        $junctionList.Add($junction)
    }

    $report = [ordered]@{
        schemaVersion = '1.0'
        generatedAt   = (Get-UtcTimestamp)
        dryRun        = [bool](Test-DryRun)
        hostname      = [string]$env:COMPUTERNAME
        userName      = [string]$env:USERNAME
        operatingSystem = $osCaption
        powershell    = $PSVersionTable.PSVersion.ToString()
        runId         = (Get-RunId)
        thresholds    = [ordered]@{
            warningUsedPercent  = $script:WarningUsedPercent
            criticalUsedPercent = $script:CriticalUsedPercent
            tempAttentionBytes  = [int64]$script:TempAttentionBytes
        }
        parameters    = [ordered]@{
            drive        = $driveParameter
            scanRoot     = $ScanRoot
            top          = $Top
            maxDepth     = $MaxDepth
            topTempFiles = $TopTempFiles
            reportPath   = $ReportPath
            logPath      = (Get-SessionLogPath)
            dryRun       = [bool](Test-DryRun)
        }
        volumes       = $volumes
        folderScan    = [ordered]@{
            root                 = $ScanRoot
            maxDepth             = $MaxDepth
            foldersConsidered    = $folderTotal
            junctionsSkipped     = $junctionList
            measuredFolderBytes  = $measuredFolderBytes
            rootFileBytes        = $rootFileBytes
            rootFileCount        = $rootFiles.totalCount
            rootFilesAccessible  = [bool]$rootFiles.accessible
            accountedBytes       = $accountedBytes
            unaccountedBytes     = $unaccountedBytes
            largestFolders       = $topFolders
            failedFolders        = $failedFolderList
            largestRootFiles     = $rootFiles.largest
        }
        temporaryFiles = [ordered]@{
            totalBytes   = $tempBytes
            totalCount   = $tempCount
            locations    = $tempResults
            largestFiles = $largestTempFiles
        }
        summary       = [ordered]@{
            overallStatus    = $overallStatus
            volumeCount      = $volumes.Count
            foldersScanned   = $folderTotal
            foldersReported  = $topFolders.Count
            junctionsSkipped = $childInfo.junctions.Count
            tempBytes        = $tempBytes
            tempFileCount    = $tempCount
            durationMs       = [int64]$stopwatch.ElapsedMilliseconds
            recommendations  = $recommendations
        }
    }

    $json = ConvertTo-ReportJson -Report $report

    if ((Test-DryRun)) {
        Write-Log -Level INFO -Event 'report_skipped' -Data @{
            reportPath = $ReportPath
            reason     = 'DryRun'
        }
        Write-Host ''
        Write-Host ("DRY-RUN: no se escribio {0}" -f $ReportPath)
    }
    else {
        [void](Write-JsonReport -Path $ReportPath -Json $json)
        $writtenBytes = [int64](Get-Item -LiteralPath $ReportPath).Length
        Write-Log -Level INFO -Event 'report_written' -Data @{
            reportPath = $ReportPath
            bytes      = $writtenBytes
        }
        Write-Host ''
        Write-Host ("Reporte escrito: {0}" -f $ReportPath)
    }

    Write-Log -Level INFO -Event 'run_completed' -Data @{
        overallStatus = $overallStatus
        durationMs    = [int64]$stopwatch.ElapsedMilliseconds
        tempBytes     = $tempBytes
        accountedBytes = $accountedBytes
    }

    Write-Host ''
    Write-Host ("Estado: {0}. Duracion: {1:N1} s." -f $overallStatus, ($stopwatch.ElapsedMilliseconds / 1000))
    Write-Output $json
    $script:ExitCode = Get-ReportExitCode -Status $overallStatus
}
catch {
    Write-ScriptError -ErrorRecord $_
    $script:ExitCode = (Get-FatalExitCode)
}

Write-Host ("Codigo de salida: {0}" -f $script:ExitCode)
if ($MyInvocation.InvocationName -ne '.') {
    exit $script:ExitCode
}
