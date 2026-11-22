#Requires -Version 5.1
<#
.SYNOPSIS
    Genera un paquete de diagnostico de Windows 11 listo para soporte.

.DESCRIPTION
    Recolecta, en solo lectura, las secciones pedidas y las deja en una carpeta
    con marca de tiempo dentro de reports\. Cada seccion tiene su propio
    manejo de errores: si una falla, las demas se generan igual y el manifiesto
    lo indica.

    Secciones:
      SystemInfo   systeminfo.exe y ficha CIM del equipo
      Disk         volumenes, discos fisicos y el reporte de DiskHealthAnalyzer
      Events       eventos Critical y Error filtrados por horas
      Services     estado y tipo de inicio de los servicios
      Network      adaptadores, IP, DNS, puerta de enlace e ipconfig
      Software     programas del Registro de desinstalacion

    No instala software, no borra archivos y no guarda contrasenas Wi-Fi.
    No usa Win32_Product (dispara reparaciones MSI).

    Sin -Include se generan todas las secciones. -Exclude las quita.

.PARAMETER Include
    Secciones a generar. Valores: SystemInfo, Disk, Events, Services,
    Network, Software. Si se omite, se generan todas.

.PARAMETER Exclude
    Secciones que no deben generarse. Tiene prioridad sobre -Include.

.PARAMETER EventHours
    Horas hacia atras para los eventos. De 1 a 168. Predeterminado: 48.

.PARAMETER MaxEvents
    Maximo de eventos que entran al paquete, los mas recientes.
    De 1 a 1000. Predeterminado: 100.

.PARAMETER EventLog
    Registros a consultar. Predeterminado: System y Application.
    Security suele exigir administrador.

.PARAMETER SkipDiskFolders
    No ejecuta DiskHealthAnalyzer. Conserva volumenes y discos fisicos.

.PARAMETER BundlePath
    Carpeta del paquete. Si se omite, se crea
    reports\SupportBundle-yyyyMMdd-HHmmss.

.PARAMETER LogPath
    JSONL de auditoria. Predeterminado: logs\SupportBundle.jsonl.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\SupportBundle.ps1

    Paquete completo con eventos de las ultimas 48 horas.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\SupportBundle.ps1 -Exclude Software -EventHours 24

    Todas las secciones menos el software, con eventos de 24 horas.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\SupportBundle.ps1 -Include Network,Services,Events -MaxEvents 50

.NOTES
    Autor: LocalAgents
    Requiere: Windows PowerShell 5.1
    Codigos de salida (Common.psm1):
      0  el paquete termino con estado Ok
      1  el paquete termino con estado Warning
      2  el paquete termino con estado Critical o Error
      3  el script fallo antes de cerrar el paquete
    El estado de cada seccion viaja dentro de manifest.json.
    El codigo de salida sigue el peor estado de las secciones ejecutadas.
    Helpers compartidos: scripts\Common.psm1.

    Recoleccion a peticion (no es una tarea diaria; systeminfo y el
    analisis de carpetas pueden tardar unos minutos):
      $script = 'C:\ProyectosDeSoftware\LocalAgents\scripts\SupportBundle.ps1'
      $arg = "-NoProfile -ExecutionPolicy Bypass -File `"$script`" -EventHours 48"
      $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arg
      $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 8:00am
      Register-ScheduledTask -TaskName 'LocalAgents-SupportBundle' -Action $action -Trigger $trigger -Description 'Paquete de diagnostico semanal'
#>
[CmdletBinding()]
param(
    [ValidateSet('SystemInfo', 'Disk', 'Events', 'Services', 'Network', 'Software')]
    [string[]]$Include,

    [ValidateSet('SystemInfo', 'Disk', 'Events', 'Services', 'Network', 'Software')]
    [string[]]$Exclude,

    [ValidateRange(1, 168)]
    [int]$EventHours = 48,

    [ValidateRange(1, 1000)]
    [int]$MaxEvents = 100,

    [string[]]$EventLog = @('System', 'Application'),

    [switch]$SkipDiskFolders,

    [string]$BundlePath,

    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -Force -DisableNameChecking

$script:ExitCode = 0
$script:BundleRoot = $null
$script:WarningDiskPercent = 80
$script:CriticalDiskPercent = 90


function Save-BundleJson {
    param(
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [Parameter(Mandatory = $true)]$Data
    )
    $full = Join-Path $script:BundleRoot $RelativePath
    Ensure-Directory -Path (Split-Path -Parent $full)
    [void](Write-JsonReport -Path $full -Report $Data)
    return $RelativePath
}

function Save-BundleText {
    param(
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [AllowEmptyString()][string]$Text
    )
    $full = Join-Path $script:BundleRoot $RelativePath
    Ensure-Directory -Path (Split-Path -Parent $full)
    if ($null -eq $Text) { $Text = '' }
    [System.IO.File]::WriteAllText($full, $Text.TrimEnd() + [Environment]::NewLine, (Get-Utf8BomEncoding))
    return $RelativePath
}

function Save-BundleCsv {
    param(
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [Parameter(Mandatory = $true)][string[]]$Headers,
        $Rows
    )
    $full = Join-Path $script:BundleRoot $RelativePath
    Ensure-Directory -Path (Split-Path -Parent $full)
    $table = New-Object System.Collections.Generic.List[object]
    if ($null -ne $Rows) {
        foreach ($row in $Rows) {
            if ($null -eq $row) { continue }
            $obj = [ordered]@{}
            foreach ($header in $Headers) {
                $value = $null
                $dictionary = $row -as [System.Collections.IDictionary]
                if ($null -ne $dictionary -and (Test-DictionaryKey -Dictionary $dictionary -Name $header)) {
                    $value = $dictionary[$header]
                }
                $obj[$header] = $value
            }
            $table.Add([pscustomobject]$obj)
        }
    }
    if ((Get-ItemCount -Value $table) -eq 0) {
        $headerLine = ($Headers -join ',') + [Environment]::NewLine
        [System.IO.File]::WriteAllText($full, $headerLine, (Get-Utf8BomEncoding))
    }
    else {
        $table.ToArray() | Export-Csv -Path $full -NoTypeInformation -Encoding UTF8
    }
    return $RelativePath
}

function Invoke-NativeText {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string]$Arguments,
        [int]$TimeoutSeconds = 90
    )
    if (-not (Test-Path -LiteralPath $FilePath)) {
        throw "No se encontro el ejecutable: $FilePath"
    }
    $oem = [System.Text.Encoding]::GetEncoding((Get-Culture).TextInfo.OEMCodePage)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $FilePath
    if ($null -eq $Arguments) { $Arguments = '' }
    $psi.Arguments = $Arguments
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.CreateNoWindow = $true
    $psi.StandardOutputEncoding = $oem
    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()
    $read = $proc.StandardOutput.ReadToEndAsync()
    $finished = $proc.WaitForExit($TimeoutSeconds * 1000)
    if (-not $finished) {
        try { $proc.Kill() } catch { }
        try { [void]$proc.WaitForExit(5000) } catch { }
        throw ("Tiempo de espera agotado ({0} s): {1}" -f $TimeoutSeconds, $FilePath)
    }
    [void]$proc.WaitForExit()
    return [string]$read.Result
}

function Get-RecentHotfixes {
    param([int]$Limit)
    $items = New-Object System.Collections.Generic.List[object]
    $fixes = @(Get-CimInstance -ClassName Win32_QuickFixEngineering -ErrorAction SilentlyContinue)
    $sorted = @($fixes | Sort-Object { $_.InstalledOn } -Descending | Select-Object -First $Limit)
    foreach ($fix in $sorted) {
        if ($null -eq $fix) { continue }
        $installed = $null
        try { if ($fix.InstalledOn) { $installed = Format-Utc -Value ([DateTime]$fix.InstalledOn) } } catch { }
        $items.Add([ordered]@{
            hotfixId    = [string]$fix.HotFixID
            description = [string]$fix.Description
            installedBy = [string]$fix.InstalledBy
            installedOn = $installed
        })
    }
    return ,$items
}

function Get-SystemInfoSection {
    $files = New-Object System.Collections.Generic.List[string]
    $exeError = $null
    $text = $null
    try {
        $systemInfo = Join-Path $env:SystemRoot 'System32\systeminfo.exe'
        $text = Invoke-NativeText -FilePath $systemInfo -Arguments '/FO LIST' -TimeoutSeconds 90
        $files.Add((Save-BundleText -RelativePath 'system\systeminfo.txt' -Text $text))
    }
    catch {
        $exeError = $_.Exception.Message
    }

    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
    $bios = $null
    $biosVersion = $null
    try {
        $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop
        $biosVersion = [string]$bios.SMBIOSBIOSVersion
    }
    catch { }

    $boot = $null
    $uptimeText = $null
    try {
        $boot = [DateTime]$os.LastBootUpTime
        $span = [DateTime]::Now - $boot
        if ($span.TotalSeconds -lt 0) { $span = [TimeSpan]::Zero }
        $days = [int][Math]::Floor($span.TotalDays)
        if ($days -gt 0) {
            $uptimeText = ('{0} d {1} h {2} min' -f $days, $span.Hours, $span.Minutes)
        }
        else {
            $uptimeText = ('{0} h {1} min' -f $span.Hours, $span.Minutes)
        }
    }
    catch { }

    $totalBytes = [int64]$os.TotalVisibleMemorySize * 1KB
    $freeBytes = [int64]$os.FreePhysicalMemory * 1KB
    $hotfixes = Get-RecentHotfixes -Limit 20
    $payload = [ordered]@{
        hostname          = [string]$env:COMPUTERNAME
        userName          = [string]$env:USERNAME
        elevated          = (Test-IsElevated)
        caption           = [string]$os.Caption
        version           = [string]$os.Version
        build             = [string]$os.BuildNumber
        architecture      = [string]$os.OSArchitecture
        installDateUtc    = (Format-Utc -Value $os.InstallDate)
        lastBootUtc       = (Format-Utc -Value $boot)
        uptimeText        = $uptimeText
        manufacturer      = [string]$cs.Manufacturer
        model             = [string]$cs.Model
        domain            = [string]$cs.Domain
        partOfDomain      = [bool]$cs.PartOfDomain
        biosVersion       = $biosVersion
        logicalProcessors = [int]$cs.NumberOfLogicalProcessors
        totalMemoryBytes  = $totalBytes
        freeMemoryBytes   = $freeBytes
        timeZoneId        = [string][TimeZoneInfo]::Local.Id
        powerShell        = [string]$PSVersionTable.PSVersion
        recentHotfixes    = $hotfixes
        systemInfoError   = $exeError
    }
    $files.Add((Save-BundleJson -RelativePath 'system\os.json' -Data $payload))

    $status = 'Ok'
    if ($exeError) { $status = 'Warning' }
    $headline = ("{0} {1}; encendido {2}." -f $payload.caption, $payload.build, $uptimeText)
    if ($exeError) { $headline = $headline + ' systeminfo.exe no quedo completo.' }

    return [ordered]@{
        status     = $status
        enabled    = $true
        files      = $files
        headline   = $headline
        error      = $exeError
    }
}

function Get-DiskHealthSnapshot {
    $logical = New-Object System.Collections.Generic.List[object]
    $statuses = New-Object System.Collections.Generic.List[string]
    $disks = @(Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop)
    foreach ($disk in $disks) {
        $size = [int64]$disk.Size
        $free = [int64]$disk.FreeSpace
        $used = $size - $free
        $percent = 0.0
        if ($size -gt 0) { $percent = [math]::Round((100.0 * $used) / $size, 2) }
        $status = 'Ok'
        if ($percent -ge $script:CriticalDiskPercent) { $status = 'Critical' }
        elseif ($percent -ge $script:WarningDiskPercent) { $status = 'Warning' }
        $statuses.Add($status)
        $logical.Add([ordered]@{
            deviceId    = [string]$disk.DeviceID
            volumeName  = [string]$disk.VolumeName
            fileSystem  = [string]$disk.FileSystem
            sizeBytes   = $size
            freeBytes   = $free
            usedBytes   = $used
            usedPercent = $percent
            status      = $status
        })
    }

    $physical = New-Object System.Collections.Generic.List[object]
    $drives = @(Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue)
    foreach ($drive in $drives) {
        $driveStatus = [string]$drive.Status
        $health = 'Ok'
        if ($driveStatus -and $driveStatus -ne 'OK') { $health = 'Warning' }
        $statuses.Add($health)
        $sizeBytes = $null
        if ($null -ne $drive.Size) { $sizeBytes = [int64]$drive.Size }
        $physical.Add([ordered]@{
            model      = [string]$drive.Model
            interface  = [string]$drive.InterfaceType
            status     = $driveStatus
            health     = $health
            sizeBytes  = $sizeBytes
            partitions = $(if ($null -eq $drive.Partitions) { $null } else { [int]$drive.Partitions })
        })
    }

    $storage = New-Object System.Collections.Generic.List[object]
    $storageError = $null
    try {
        $physicalDisks = @(Get-PhysicalDisk -ErrorAction Stop)
        foreach ($item in $physicalDisks) {
            $healthText = [string]$item.HealthStatus
            $health = 'Ok'
            if ($healthText -and $healthText -ne 'Healthy') { $health = 'Warning' }
            $statuses.Add($health)
            $sizeBytes = $null
            if ($null -ne $item.Size) { $sizeBytes = [int64]$item.Size }
            $storage.Add([ordered]@{
                friendlyName      = [string]$item.FriendlyName
                mediaType         = [string]$item.MediaType
                busType           = [string]$item.BusType
                healthStatus      = $healthText
                operationalStatus = [string]$item.OperationalStatus
                health            = $health
                sizeBytes         = $sizeBytes
            })
        }
    }
    catch {
        $storageError = $_.Exception.Message
    }

    $overall = Get-WorstStatus -Statuses ($statuses.ToArray())
    return [ordered]@{
        status       = $overall
        logical      = $logical
        physical     = $physical
        storage      = $storage
        storageError = $storageError
    }
}

function Invoke-DiskFolderReport {
    param(
        [Parameter(Mandatory = $true)][string]$ReportFile,
        [Parameter(Mandatory = $true)][string]$AnalyzerLog
    )
    $analyzer = Join-Path (Join-Path (Get-RepositoryRoot) 'scripts') 'DiskHealthAnalyzer.ps1'
    if (-not (Test-Path -LiteralPath $analyzer)) {
        return [ordered]@{
            ok       = $false
            exitCode = $null
            error    = 'No se encontro scripts\DiskHealthAnalyzer.ps1.'
        }
    }
    $output = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $analyzer -ReportPath $ReportFile -LogPath $AnalyzerLog 2>&1
    $code = $LASTEXITCODE
    # 0, 1 y 2 son el reporte cerrado (Ok, Warning, Critical). 3 es fallo.
    $finished = ($null -ne $code -and $code -ge 0 -and $code -le 2)
    $ok = ($finished -and (Test-Path -LiteralPath $ReportFile))
    $errorText = $null
    if (-not $ok) {
        $text = ($output | Out-String).Trim()
        if ($text.Length -gt 1500) { $text = $text.Substring($text.Length - 1500) }
        if ([string]::IsNullOrWhiteSpace($text)) { $text = "DiskHealthAnalyzer termino con codigo $code." }
        $errorText = $text
    }
    return [ordered]@{
        ok       = $ok
        exitCode = $code
        error    = $errorText
    }
}

function Get-DiskSection {
    $files = New-Object System.Collections.Generic.List[string]
    $snapshot = Get-DiskHealthSnapshot
    $files.Add((Save-BundleJson -RelativePath 'disk\volumes.json' -Data $snapshot))

    $folderError = $null
    $folderOk = $false
    if (-not $SkipDiskFolders) {
        $reportFile = Join-Path $script:BundleRoot 'disk\DiskHealth.json'
        $analyzerLog = Join-Path $script:BundleRoot 'disk\DiskHealth.jsonl'
        $folder = Invoke-DiskFolderReport -ReportFile $reportFile -AnalyzerLog $analyzerLog
        $folderOk = [bool]$folder.ok
        if ($folderOk) {
            $files.Add('disk\DiskHealth.json')
            if (Test-Path -LiteralPath $analyzerLog) { $files.Add('disk\DiskHealth.jsonl') }
        }
        else {
            $folderError = [string]$folder.error
        }
    }

    $status = [string]$snapshot.status
    if ($folderError) {
        if ($status -eq 'Ok') { $status = 'Warning' }
    }
    $volumeCount = Get-ItemCount -Value $snapshot.logical
    $headline = ("{0} volumen(es) fijo(s)." -f $volumeCount)
    if ($SkipDiskFolders) { $headline = $headline + ' Analisis de carpetas omitido.' }
    elseif ($folderOk) { $headline = $headline + ' Carpetas de primer nivel incluidas.' }
    else { $headline = $headline + ' El analisis de carpetas no quedo completo.' }

    return [ordered]@{
        status   = $status
        enabled  = $true
        files    = $files
        headline = $headline
        error    = $folderError
    }
}

function Get-EventSection {
    $start = (Get-Date).AddHours(-1 * $EventHours)
    $found = New-Object System.Collections.Generic.List[object]
    $logErrors = New-Object System.Collections.Generic.List[string]

    foreach ($logName in $EventLog) {
        if ([string]::IsNullOrWhiteSpace($logName)) { continue }
        try {
            $events = @(Get-WinEvent -FilterHashtable @{
                LogName   = $logName
                Level     = 1, 2
                StartTime = $start
            } -MaxEvents $MaxEvents -ErrorAction Stop)
            foreach ($event in $events) {
                $message = [string]$event.Message
                if ($message) {
                    $message = ($message -replace '\s+', ' ').Trim()
                    if ($message.Length -gt 500) { $message = $message.Substring(0, 500) }
                }
                $levelName = 'Error'
                if ([int]$event.Level -eq 1) { $levelName = 'Critical' }
                $created = $null
                try { $created = [DateTime]$event.TimeCreated } catch { }
                $found.Add([ordered]@{
                    timeCreatedUtc = (Format-Utc -Value $created)
                    logName        = [string]$event.LogName
                    level          = $levelName
                    eventId        = [int]$event.Id
                    provider       = [string]$event.ProviderName
                    message        = $message
                })
            }
        }
        catch {
            $errorId = [string]$_.FullyQualifiedErrorId
            if ($errorId -like 'NoMatchingEventsFound*') { continue }
            $logErrors.Add(('{0}: {1}' -f $logName, $_.Exception.Message))
        }
    }

    $selected = New-Object System.Collections.Generic.List[object]
    $ordered = @($found.ToArray() | Sort-Object { $_.timeCreatedUtc } -Descending | Select-Object -First $MaxEvents)
    $criticalEvents = 0
    foreach ($item in $ordered) {
        if ($null -eq $item) { continue }
        $selected.Add($item)
        if ($item.level -eq 'Critical') { $criticalEvents++ }
    }

    $status = 'Ok'
    if ((Get-ItemCount -Value $logErrors) -gt 0) { $status = 'Warning' }
    if ((Get-ItemCount -Value $selected) -gt 0 -and $status -eq 'Ok') { $status = 'Warning' }
    if ($criticalEvents -gt 0) { $status = 'Critical' }

    $payload = [ordered]@{
        hours         = $EventHours
        maxEvents     = $MaxEvents
        startLocal    = $start.ToString('yyyy-MM-ddTHH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
        logs          = $EventLog
        count         = (Get-ItemCount -Value $selected)
        criticalCount = $criticalEvents
        logErrors     = $logErrors
        items         = $selected
    }
    $files = New-Object System.Collections.Generic.List[string]
    $files.Add((Save-BundleJson -RelativePath 'events\errors.json' -Data $payload))
    $files.Add((Save-BundleCsv -RelativePath 'events\errors.csv' -Headers @('timeCreatedUtc', 'logName', 'level', 'eventId', 'provider', 'message') -Rows $selected))

    $errorText = $null
    if ((Get-ItemCount -Value $logErrors) -gt 0) {
        $errorText = [string]$logErrors[0]
    }
    return [ordered]@{
        status   = $status
        enabled  = $true
        files    = $files
        headline = ("{0} eventos Critical o Error en {1} h." -f (Get-ItemCount -Value $selected), $EventHours)
        error    = $errorText
    }
}

function Get-ServiceSection {
    $items = New-Object System.Collections.Generic.List[object]
    $attention = New-Object System.Collections.Generic.List[object]
    $services = @(Get-CimInstance -ClassName Win32_Service -ErrorAction Stop)
    foreach ($service in $services) {
        $delayed = $false
        try { if ($null -ne $service.DelayedAutoStart) { $delayed = [bool]$service.DelayedAutoStart } } catch { }
        $processId = $null
        try { if ($null -ne $service.ProcessId) { $processId = [int]$service.ProcessId } } catch { }
        $record = [ordered]@{
            name             = [string]$service.Name
            displayName      = [string]$service.DisplayName
            state            = [string]$service.State
            startMode        = [string]$service.StartMode
            delayedAutoStart = $delayed
            startName        = [string]$service.StartName
            processId        = $processId
        }
        $items.Add($record)
        $running = ($record.state -eq 'Running')
        if (($record.startMode -eq 'Auto') -and -not $running) {
            $attention.Add($record)
        }
    }

    $sorted = @($items.ToArray() | Sort-Object { $_.name })
    $sortedItems = New-Object System.Collections.Generic.List[object]
    foreach ($item in $sorted) {
        if ($null -ne $item) { $sortedItems.Add($item) }
    }
    $sortedAttention = @($attention.ToArray() | Sort-Object { $_.name })
    $attentionItems = New-Object System.Collections.Generic.List[object]
    foreach ($item in $sortedAttention) {
        if ($null -ne $item) { $attentionItems.Add($item) }
    }

    $attentionCount = Get-ItemCount -Value $attentionItems
    $status = 'Ok'
    if ($attentionCount -gt 0) { $status = 'Warning' }
    $payload = [ordered]@{
        count                    = (Get-ItemCount -Value $sortedItems)
        automaticNotRunningCount = $attentionCount
        automaticNotRunning      = $attentionItems
        items                    = $sortedItems
    }
    $files = New-Object System.Collections.Generic.List[string]
    $files.Add((Save-BundleJson -RelativePath 'services\services.json' -Data $payload))
    $files.Add((Save-BundleCsv -RelativePath 'services\services.csv' -Headers @('name', 'displayName', 'state', 'startMode', 'delayedAutoStart', 'startName', 'processId') -Rows $sortedItems))

    return [ordered]@{
        status   = $status
        enabled  = $true
        files    = $files
        headline = ("{0} servicios. {1} automaticos detenidos." -f (Get-ItemCount -Value $sortedItems), $attentionCount)
        error    = $null
    }
}

function Get-NetworkSection {
    $errors = New-Object System.Collections.Generic.List[string]
    $adapters = New-Object System.Collections.Generic.List[object]
    $addresses = New-Object System.Collections.Generic.List[object]
    $gateways = New-Object System.Collections.Generic.List[object]
    $dns = New-Object System.Collections.Generic.List[object]
    $profiles = New-Object System.Collections.Generic.List[object]
    $upCount = 0

    try {
        $items = @(Get-NetAdapter -ErrorAction Stop)
        foreach ($item in $items) {
            $status = [string]$item.Status
            if ($status -eq 'Up') { $upCount++ }
            $adapters.Add([ordered]@{
                name        = [string]$item.Name
                description = [string]$item.InterfaceDescription
                status      = $status
                macAddress  = [string]$item.MacAddress
                linkSpeed   = [string]$item.LinkSpeed
                mediaType   = [string]$item.MediaType
            })
        }
    }
    catch {
        $errors.Add(('Get-NetAdapter: {0}' -f $_.Exception.Message))
    }

    try {
        $items = @(Get-NetIPAddress -ErrorAction Stop)
        foreach ($item in $items) {
            $addresses.Add([ordered]@{
                alias        = [string]$item.InterfaceAlias
                family       = [string]$item.AddressFamily
                address      = [string]$item.IPAddress
                prefixLength = $(if ($null -eq $item.PrefixLength) { $null } else { [int]$item.PrefixLength })
            })
        }
    }
    catch {
        $errors.Add(('Get-NetIPAddress: {0}' -f $_.Exception.Message))
    }

    try {
        $items = @(Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction Stop)
        foreach ($item in $items) {
            $gateways.Add([ordered]@{
                alias    = [string]$item.InterfaceAlias
                nextHop  = [string]$item.NextHop
                metric   = $(if ($null -eq $item.RouteMetric) { $null } else { [int]$item.RouteMetric })
            })
        }
    }
    catch {
        $errors.Add(('Get-NetRoute: {0}' -f $_.Exception.Message))
    }

    try {
        $items = @(Get-DnsClientServerAddress -ErrorAction Stop)
        foreach ($item in $items) {
            $servers = New-Object System.Collections.Generic.List[string]
            if ($null -ne $item.ServerAddresses) {
                foreach ($server in @($item.ServerAddresses)) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$server)) { $servers.Add([string]$server) }
                }
            }
            if ((Get-ItemCount -Value $servers) -eq 0) { continue }
            $dns.Add([ordered]@{
                alias   = [string]$item.InterfaceAlias
                family  = [string]$item.AddressFamily
                servers = $servers
            })
        }
    }
    catch {
        $errors.Add(('Get-DnsClientServerAddress: {0}' -f $_.Exception.Message))
    }

    try {
        $items = @(Get-NetConnectionProfile -ErrorAction Stop)
        foreach ($item in $items) {
            $profiles.Add([ordered]@{
                name         = [string]$item.Name
                alias        = [string]$item.InterfaceAlias
                category     = [string]$item.NetworkCategory
                connectivity = [string]$item.IPv4Connectivity
            })
        }
    }
    catch {
        $errors.Add(('Get-NetConnectionProfile: {0}' -f $_.Exception.Message))
    }

    $ipconfigError = $null
    $files = New-Object System.Collections.Generic.List[string]
    try {
        $ipconfig = Join-Path $env:SystemRoot 'System32\ipconfig.exe'
        $text = Invoke-NativeText -FilePath $ipconfig -Arguments '/all' -TimeoutSeconds 30
        $files.Add((Save-BundleText -RelativePath 'network\ipconfig.txt' -Text $text))
    }
    catch {
        $ipconfigError = $_.Exception.Message
        $errors.Add(('ipconfig: {0}' -f $ipconfigError))
    }

    $payload = [ordered]@{
        adaptersUp = $upCount
        adapters   = $adapters
        addresses  = $addresses
        gateways   = $gateways
        dns        = $dns
        profiles   = $profiles
        errors     = $errors
    }
    $files.Add((Save-BundleJson -RelativePath 'network\network.json' -Data $payload))

    $status = 'Ok'
    if ($upCount -eq 0 -or (Get-ItemCount -Value $errors) -gt 0) { $status = 'Warning' }
    if ((Get-ItemCount -Value $adapters) -eq 0 -and (Get-ItemCount -Value $addresses) -eq 0) { $status = 'Error' }
    $errorText = $null
    if ((Get-ItemCount -Value $errors) -gt 0) { $errorText = [string]$errors[0] }

    return [ordered]@{
        status   = $status
        enabled  = $true
        files    = $files
        headline = ("{0} adaptador(es), {1} en Up." -f (Get-ItemCount -Value $adapters), $upCount)
        error    = $errorText
    }
}

function Read-UninstallHive {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Scope
    )
    $items = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $Root)) { return ,$items }
    $key = $null
    try {
        $key = Get-Item -LiteralPath $Root -ErrorAction Stop
        foreach ($name in $key.GetSubKeyNames()) {
            $sub = $null
            try {
                $sub = $key.OpenSubKey($name)
                if ($null -eq $sub) { continue }
                $display = [string]$sub.GetValue('DisplayName')
                if ([string]::IsNullOrWhiteSpace($display)) { continue }
                $systemComponent = 0
                try { $systemComponent = [int]$sub.GetValue('SystemComponent') } catch { $systemComponent = 0 }
                if ($systemComponent -eq 1) { continue }
                $release = [string]$sub.GetValue('ReleaseType')
                if ($release -eq 'Security Update' -or $release -eq 'Update' -or $release -eq 'Hotfix') { continue }
                if ($display -match '^KB\d+') { continue }
                $sizeKb = $null
                $sizeRaw = $sub.GetValue('EstimatedSize')
                if ($null -ne $sizeRaw) {
                    try { $sizeKb = [int64]$sizeRaw } catch { }
                }
                $items.Add([ordered]@{
                    name        = $display.Trim()
                    version     = [string]$sub.GetValue('DisplayVersion')
                    publisher   = [string]$sub.GetValue('Publisher')
                    installDate = [string]$sub.GetValue('InstallDate')
                    sizeKb      = $sizeKb
                    scope       = $Scope
                })
            }
            catch { }
            finally {
                if ($sub) { $sub.Close() }
            }
        }
    }
    finally {
        if ($key) { $key.Close() }
    }
    return ,$items
}

function Get-SoftwareSection {
    $items = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    $hives = @(
        @{ Root = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'; Scope = 'HKLM' },
        @{ Root = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'; Scope = 'HKLM32' },
        @{ Root = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'; Scope = 'HKCU' }
    )
    foreach ($hive in $hives) {
        $read = Read-UninstallHive -Root ([string]$hive.Root) -Scope ([string]$hive.Scope)
        foreach ($app in $read) {
            $key = ('{0}|{1}|{2}' -f [string]$app.name, [string]$app.version, [string]$app.publisher).ToLowerInvariant()
            if ($seen.ContainsKey($key)) { continue }
            $seen[$key] = $true
            $items.Add($app)
        }
    }

    $sorted = @($items.ToArray() | Sort-Object { $_.name }, { $_.version })
    $sortedItems = New-Object System.Collections.Generic.List[object]
    foreach ($item in $sorted) {
        if ($null -ne $item) { $sortedItems.Add($item) }
    }
    $count = Get-ItemCount -Value $sortedItems
    $payload = [ordered]@{
        count = $count
        items = $sortedItems
        note  = 'Se omiten componentes de sistema, revisiones KB y Win32_Product.'
    }
    $files = New-Object System.Collections.Generic.List[string]
    $files.Add((Save-BundleJson -RelativePath 'software\installed.json' -Data $payload))
    $files.Add((Save-BundleCsv -RelativePath 'software\installed.csv' -Headers @('name', 'version', 'publisher', 'installDate', 'sizeKb', 'scope') -Rows $sortedItems))

    $status = 'Ok'
    if ($count -eq 0) { $status = 'Warning' }
    return [ordered]@{
        status   = $status
        enabled  = $true
        files    = $files
        headline = ("{0} programas de usuario." -f $count)
        error    = $null
    }
}

function New-SkippedSection {
    return [ordered]@{
        status   = 'Skipped'
        enabled  = $false
        files    = (New-Object System.Collections.Generic.List[string])
        headline = 'Seccion excluida.'
        error    = $null
    }
}

function Invoke-BundleSection {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][scriptblock]$Action
    )
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $value = & $Action
        $sw.Stop()
        $dictionary = $value -as [System.Collections.IDictionary]
        if ($null -ne $dictionary) {
            $dictionary['durationMs'] = [int64]$sw.ElapsedMilliseconds
        }
        $status = $null
        if ($null -ne $dictionary -and $dictionary.Contains('status')) { $status = [string]$dictionary['status'] }
        Write-Log -Level INFO -Event 'section_completed' -Data @{
            section    = $Name
            status     = [string]$status
            durationMs = [int64]$sw.ElapsedMilliseconds
        }
        return $value
    }
    catch {
        $sw.Stop()
        $message = $_.Exception.Message
        Write-Log -Level ERROR -Event 'section_failed' -Data @{
            section = $Name
            message = $message
        }
        Write-Warning ("{0}: {1}" -f $Name, $message)
        return [ordered]@{
            status     = 'Error'
            enabled    = $true
            files      = (New-Object System.Collections.Generic.List[string])
            headline   = 'La seccion fallo.'
            error      = $message
            durationMs = [int64]$sw.ElapsedMilliseconds
        }
    }
}

function Write-BundleTree {
    param([Parameter(Mandatory = $true)][string]$Root)
    $leaf = Split-Path -Leaf $Root
    Write-Host $leaf
    Write-TreeLevel -DirectoryPath $Root -Prefix ''
}

function Write-TreeLevel {
    param(
        [Parameter(Mandatory = $true)][string]$DirectoryPath,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Prefix
    )
    $children = @(Get-ChildItem -LiteralPath $DirectoryPath -Force | Sort-Object @{ Expression = { $_.PSIsContainer }; Descending = $true }, @{ Expression = { $_.Name } })
    $total = $children.Count
    $index = 0
    foreach ($child in $children) {
        $index++
        $last = ($index -eq $total)
        $branch = '+-- '
        if ($last) { $branch = '\-- ' }
        if ($child.PSIsContainer) {
            Write-Host ($Prefix + $branch + $child.Name + '\')
            $next = $Prefix + '|   '
            if ($last) { $next = $Prefix + '    ' }
            Write-TreeLevel -DirectoryPath $child.FullName -Prefix $next
        }
        else {
            Write-Host ($Prefix + $branch + $child.Name + '  (' + (Format-ByteSize -Bytes ([int64]$child.Length)) + ')')
        }
    }
}

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }

    $catalog = @('SystemInfo', 'Disk', 'Events', 'Services', 'Network', 'Software')
    $selected = New-Object System.Collections.Generic.List[string]
    if ($Include -and @($Include).Count -gt 0) {
        foreach ($name in $Include) {
            if (-not $selected.Contains($name)) { $selected.Add($name) }
        }
    }
    else {
        foreach ($name in $catalog) { $selected.Add($name) }
    }
    if ($Exclude) {
        foreach ($name in $Exclude) { [void]$selected.Remove($name) }
    }
    if ((Get-ItemCount -Value $selected) -eq 0) {
        throw 'No quedo ninguna seccion seleccionada. Revise -Include y -Exclude.'
    }

    $repoRoot = Get-RepositoryRoot
    $stamp = [DateTime]::Now.ToString('yyyyMMdd-HHmmss')
    if ([string]::IsNullOrWhiteSpace($BundlePath)) {
        $BundlePath = Join-Path (Join-Path $repoRoot 'reports') ("SupportBundle-{0}" -f $stamp)
    }
    if ([string]::IsNullOrWhiteSpace($LogPath)) {
        $LogPath = Join-Path (Join-Path $repoRoot 'logs') 'SupportBundle.jsonl'
    }
    $BundlePath = [System.IO.Path]::GetFullPath($BundlePath)
    if (Test-Path -LiteralPath $BundlePath) {
        throw "La carpeta del paquete ya existe: $BundlePath"
    }
    Initialize-CommonSession -ScriptName 'SupportBundle' -LogPath ([System.IO.Path]::GetFullPath($LogPath)) -DryRun $false
    Ensure-Directory -Path $BundlePath
    $script:BundleRoot = $BundlePath

    $do = @{
        SystemInfo = $selected.Contains('SystemInfo')
        Disk       = $selected.Contains('Disk')
        Events     = $selected.Contains('Events')
        Services   = $selected.Contains('Services')
        Network    = $selected.Contains('Network')
        Software   = $selected.Contains('Software')
    }

    Write-Host 'SupportBundle 1.0'
    Write-Host ("Carpeta: {0}" -f $BundlePath)
    Write-Host ("Secciones: {0}" -f ($selected -join ', '))
    Write-Host ("Elevado: {0}" -f (Test-IsElevated))

    Write-Log -Level INFO -Event 'run_start' -Data @{
        bundlePath       = $BundlePath
        sections         = ($selected -join ',')
        eventHours       = $EventHours
        maxEvents        = $MaxEvents
        skipDiskFolders  = [bool]$SkipDiskFolders
        elevated         = (Test-IsElevated)
    }

    $sections = [ordered]@{}
    if ($do.SystemInfo) {
        $sections.SystemInfo = Invoke-BundleSection -Name 'SystemInfo' -Action { Get-SystemInfoSection }
    }
    else { $sections.SystemInfo = New-SkippedSection }

    if ($do.Disk) {
        $sections.Disk = Invoke-BundleSection -Name 'Disk' -Action { Get-DiskSection }
    }
    else { $sections.Disk = New-SkippedSection }

    if ($do.Events) {
        $sections.Events = Invoke-BundleSection -Name 'Events' -Action { Get-EventSection }
    }
    else { $sections.Events = New-SkippedSection }

    if ($do.Services) {
        $sections.Services = Invoke-BundleSection -Name 'Services' -Action { Get-ServiceSection }
    }
    else { $sections.Services = New-SkippedSection }

    if ($do.Network) {
        $sections.Network = Invoke-BundleSection -Name 'Network' -Action { Get-NetworkSection }
    }
    else { $sections.Network = New-SkippedSection }

    if ($do.Software) {
        $sections.Software = Invoke-BundleSection -Name 'Software' -Action { Get-SoftwareSection }
    }
    else { $sections.Software = New-SkippedSection }

    $statuses = New-Object System.Collections.Generic.List[string]
    $findings = New-Object System.Collections.Generic.List[string]
    foreach ($name in $catalog) {
        $section = $sections[$name]
        $status = 'Unknown'
        $headline = ''
        $dictionary = $section -as [System.Collections.IDictionary]
        if ($null -ne $dictionary) {
            if ($dictionary.Contains('status')) { $status = [string]$dictionary['status'] }
            if ($dictionary.Contains('headline')) { $headline = [string]$dictionary['headline'] }
        }
        if ($status -and $status -ne 'Skipped') { $statuses.Add($status) }
        if ($headline) { $findings.Add(('{0}: {1}' -f $name, $headline)) }
    }
    $overall = Get-WorstStatus -Statuses ($statuses.ToArray())
    $stopwatch.Stop()

    $readme = New-Object System.Text.StringBuilder
    [void]$readme.AppendLine('Paquete de diagnostico LocalAgents')
    [void]$readme.AppendLine(('Generado: {0}' -f (Get-UtcTimestamp)))
    [void]$readme.AppendLine(('Equipo: {0}' -f $env:COMPUTERNAME))
    [void]$readme.AppendLine(('Usuario: {0}' -f $env:USERNAME))
    [void]$readme.AppendLine(('Elevado: {0}' -f (Test-IsElevated)))
    [void]$readme.AppendLine(('Estado global: {0}' -f $overall))
    [void]$readme.AppendLine('')
    [void]$readme.AppendLine('Contenido:')
    [void]$readme.AppendLine('  manifest.json          indice del paquete y estado de cada seccion')
    [void]$readme.AppendLine('  system\systeminfo.txt  salida de systeminfo.exe')
    [void]$readme.AppendLine('  system\os.json         ficha del sistema operativo')
    [void]$readme.AppendLine('  disk\volumes.json      volumenes y discos fisicos')
    [void]$readme.AppendLine('  disk\DiskHealth.json   carpetas de primer nivel, si el analisis termino')
    [void]$readme.AppendLine('  events\errors.json     eventos Critical y Error')
    [void]$readme.AppendLine('  services\services.csv  servicios y tipo de inicio')
    [void]$readme.AppendLine('  network\ipconfig.txt   ipconfig /all')
    [void]$readme.AppendLine('  network\network.json   adaptadores, IP, DNS y perfiles')
    [void]$readme.AppendLine('  software\installed.csv programas visibles en el Registro')
    [void]$readme.AppendLine('')
    [void]$readme.AppendLine('Este paquete no incluye contrasenas Wi-Fi ni ejecuta Win32_Product.')
    [void]$readme.AppendLine('Las secciones ausentes fueron excluidas o fallaron; el detalle esta en manifest.json.')
    Save-BundleText -RelativePath 'README.txt' -Text $readme.ToString() | Out-Null

    $manifest = [ordered]@{
        schemaVersion = '1.0'
        generatedAt   = (Get-UtcTimestamp)
        hostname      = [string]$env:COMPUTERNAME
        userName      = [string]$env:USERNAME
        runId         = (Get-RunId)
        elevated      = (Test-IsElevated)
        bundlePath    = $BundlePath
        parameters    = [ordered]@{
            eventHours      = $EventHours
            maxEvents       = $MaxEvents
            eventLog        = $EventLog
            skipDiskFolders = [bool]$SkipDiskFolders
            sections        = $selected
        }
        sections      = $sections
        summary       = [ordered]@{
            overallStatus = $overall
            findings      = $findings
            durationMs    = [int64]$stopwatch.ElapsedMilliseconds
        }
    }
    Save-BundleJson -RelativePath 'manifest.json' -Data $manifest | Out-Null

    Write-Host ''
    Write-Host ("Estado global: {0}" -f $overall)
    foreach ($finding in $findings) {
        Write-Host ("- {0}" -f $finding)
    }
    Write-Host ''
    try {
        Write-BundleTree -Root $BundlePath
    }
    catch {
        Write-Warning ("No se pudo imprimir el arbol: {0}" -f $_.Exception.Message)
    }
    Write-Host ''
    Write-Host ("Duracion: {0:N1} s" -f ($stopwatch.ElapsedMilliseconds / 1000.0))

    Write-Log -Level INFO -Event 'run_completed' -Data @{
        overallStatus = $overall
        bundlePath    = $BundlePath
        durationMs    = [int64]$stopwatch.ElapsedMilliseconds
    }
    $script:ExitCode = Get-ReportExitCode -Status $overall
}
catch {
    Write-ScriptError -ErrorRecord $_
    $script:ExitCode = (Get-FatalExitCode)
}

Write-Host ("Codigo de salida: {0}" -f $script:ExitCode)
if ($MyInvocation.InvocationName -ne '.') {
    exit $script:ExitCode
}
