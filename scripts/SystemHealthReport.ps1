#Requires -Version 5.1
<#
.SYNOPSIS
    Genera un reporte de salud de Windows 11 en JSON y, si se pide, en HTML.

.DESCRIPTION
    Recolecta, en solo lectura:

      - uso de CPU y memoria fisica
      - volumenes logicos y discos fisicos
      - servicios criticos e importantes
      - eventos Critical y Error de los registros indicados
      - tiempo encendido desde el ultimo arranque
      - estado de Windows Update, reinicio pendiente y actualizaciones recientes
      - dispositivos con error de controlador

    Cada seccion tiene su propio manejo de errores. Si una fuente falla, el
    reporte se genera igual y esa seccion queda marcada como Error. En el
    estado global eso cuenta como Warning.

    Critical, y por tanto el codigo de salida 2, solo sale de un hallazgo real:
      - servicio Required ausente, deshabilitado, o Automatico sin retraso y detenido
      - memoria fisica al 95% o mas, o un volumen al 90% o mas
      - evento de nivel 1 (Critical) en la ventana pedida
      - dispositivo con codigo 10, 12, 18, 24, 28 o 31, si no es firmware ni UEFI
      - el servicio wuauserv no existe
    No es Critical: NlaSvc u otro Required en Manual y detenido, un Automatico
    con DelayedAutoStart detenido, ni un firmware o dispositivo UEFI. Esos
    quedan en Warning. La CPU alta tambien queda en Warning.

.PARAMETER EventHours
    Horas hacia atras para leer eventos. De 1 a 168. Predeterminado: 24.
    Use 48 para cubrir dos dias.

.PARAMETER MaxEvents
    Maximo de eventos que entran al reporte, los mas recientes. De 1 a 500.
    Predeterminado: 40.

.PARAMETER EventLog
    Registros de Windows a consultar. Predeterminado: System y Application.
    Security suele exigir administrador.

.PARAMETER ReportPath
    Ruta del JSON. Si se omite, se crea
    reports\SystemHealth-yyyyMMdd-HHmmss.json.

.PARAMETER Html
    Ademas del JSON, escribe un HTML de la misma fotografia.

.PARAMETER HtmlPath
    Ruta del HTML. Si se usa -Html y se omite, se crea junto al JSON.

.PARAMETER LogPath
    JSONL de auditoria. Predeterminado: logs\SystemHealthReport.jsonl.

.PARAMETER SkipPendingUpdates
    No lee reports\WindowsUpdateSnapshot.json. Este script no abre el
    catalogo: UpdateChecker es la unica consulta de la cadena diaria.
    El resto del estado de Windows Update se conserva.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\SystemHealthReport.ps1

    Reporte JSON con eventos de las ultimas 24 horas.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\SystemHealthReport.ps1 -EventHours 48 -Html

    Reporte JSON y HTML con eventos de las ultimas 48 horas.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\SystemHealthReport.ps1 -EventHours 36 -MaxEvents 80 -EventLog System,Application

.NOTES
    Autor: LocalAgents
    Requiere: Windows PowerShell 5.1
    Codigos de salida:
      0  todas las secciones medidas quedaron Ok
      1  hay Warning, o alguna seccion no se pudo leer, y no hay Critical
      2  el reporte contiene al menos un Critical real
      3  el script fallo antes de cerrar el reporte
    El catalogo de Windows Update no se consulta aqui. Si UpdateChecker dejo
    reports\WindowsUpdateSnapshot.json con menos de 20 horas, se reutiliza.
    Helpers compartidos: scripts\Common.psm1.

    Programador de tareas (cada dia a las 07:15, JSON y HTML, 48 horas):
      $script = 'C:\ProyectosDeSoftware\LocalAgents\scripts\SystemHealthReport.ps1'
      $arg = "-NoProfile -ExecutionPolicy Bypass -File `"$script`" -EventHours 48 -Html"
      $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arg
      $trigger = New-ScheduledTaskTrigger -Daily -At 7:15am
      Register-ScheduledTask -TaskName 'LocalAgents-SystemHealth' -Action $action -Trigger $trigger -Description 'Reporte de salud de Windows'
#>
[CmdletBinding()]
param(
    [ValidateRange(1, 168)]
    [int]$EventHours = 24,

    [ValidateRange(1, 500)]
    [int]$MaxEvents = 40,

    [string[]]$EventLog = @('System', 'Application'),

    [string]$ReportPath,

    [switch]$Html,

    [string]$HtmlPath,

    [string]$LogPath,

    [switch]$SkipPendingUpdates
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -Force -DisableNameChecking

$script:ExitCode = 0
$script:WarningDiskPercent = 80
$script:CriticalDiskPercent = 90
$script:WarningMemoryPercent = 90
$script:CriticalMemoryPercent = 95
$script:WarningCpuPercent = 90


function Format-DurationText {
    param([TimeSpan]$Span)
    $days = [int][Math]::Floor($Span.TotalDays)
    if ($days -gt 0) {
        return ('{0} d {1} h {2} min' -f $days, $Span.Hours, $Span.Minutes)
    }
    return ('{0} h {1} min' -f $Span.Hours, $Span.Minutes)
}


function ConvertTo-HtmlText {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Get-SectionError {
    param([string]$Message)
    return [ordered]@{
        status = 'Error'
        error  = $Message
    }
}

function Get-CpuReport {
    $processors = @(Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop)
    $names = New-Object System.Collections.Generic.List[string]
    $cores = 0
    $logical = 0
    $loadValues = New-Object System.Collections.Generic.List[int]
    foreach ($cpu in $processors) {
        if ($cpu.Name) { $names.Add(([string]$cpu.Name).Trim()) }
        $cores += [int]$cpu.NumberOfCores
        $logical += [int]$cpu.NumberOfLogicalProcessors
        if ($null -ne $cpu.LoadPercentage) {
            $loadValues.Add([int]$cpu.LoadPercentage)
        }
    }
    $load = $null
    if ($loadValues.Count -gt 0) {
        $sum = 0
        foreach ($value in $loadValues) { $sum += $value }
        $load = [math]::Round($sum / $loadValues.Count, 1)
    }
    $status = 'Unknown'
    if ($null -ne $load) {
        if ($load -ge $script:WarningCpuPercent) { $status = 'Warning' } else { $status = 'Ok' }
    }
    return [ordered]@{
        status              = $status
        name                = ($names -join ' | ')
        sockets             = $processors.Count
        cores               = $cores
        logicalProcessors   = $logical
        loadPercent         = $load
        error               = $null
    }
}

function Get-MemoryReport {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $total = [int64]$os.TotalVisibleMemorySize * 1KB
    $free = [int64]$os.FreePhysicalMemory * 1KB
    $used = $total - $free
    $percent = 0.0
    if ($total -gt 0) {
        $percent = [math]::Round((100.0 * $used) / $total, 2)
    }
    $status = 'Ok'
    if ($percent -ge $script:CriticalMemoryPercent) { $status = 'Critical' }
    elseif ($percent -ge $script:WarningMemoryPercent) { $status = 'Warning' }
    return [ordered]@{
        status      = $status
        totalBytes  = $total
        freeBytes   = $free
        usedBytes   = $used
        usedPercent = $percent
        error       = $null
    }
}

function Get-UptimeReport {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $boot = [DateTime]$os.LastBootUpTime
    $span = [DateTime]::Now - $boot
    if ($span.TotalSeconds -lt 0) { $span = [TimeSpan]::Zero }
    return [ordered]@{
        status         = 'Ok'
        lastBootLocal  = $boot.ToString('yyyy-MM-ddTHH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
        lastBootUtc    = (Format-Utc -Value $boot)
        uptimeSeconds  = [int64][Math]::Floor($span.TotalSeconds)
        uptimeText     = (Format-DurationText -Span $span)
        error          = $null
    }
}

function Get-DiskReport {
    $logical = New-Object System.Collections.Generic.List[object]
    $physical = New-Object System.Collections.Generic.List[object]
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

    $drives = @(Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue)
    foreach ($drive in $drives) {
        $driveStatus = [string]$drive.Status
        $health = 'Ok'
        if ($driveStatus -and $driveStatus -ne 'OK') { $health = 'Warning' }
        $statuses.Add($health)
        $physical.Add([ordered]@{
            model        = [string]$drive.Model
            interface    = [string]$drive.InterfaceType
            status       = $driveStatus
            health       = $health
            sizeBytes    = $(if ($null -eq $drive.Size) { $null } else { [int64]$drive.Size })
            partitions   = $(if ($null -eq $drive.Partitions) { $null } else { [int]$drive.Partitions })
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
            $operational = [string]$item.OperationalStatus
            $storage.Add([ordered]@{
                friendlyName      = [string]$item.FriendlyName
                mediaType         = [string]$item.MediaType
                healthStatus      = $healthText
                operationalStatus = $operational
                health            = $health
                sizeBytes         = $(if ($null -eq $item.Size) { $null } else { [int64]$item.Size })
            })
        }
    }
    catch {
        $storageError = $_.Exception.Message
    }

    return [ordered]@{
        status        = (Get-WorstStatus -Statuses ($statuses.ToArray()))
        logical       = $logical
        physical      = $physical
        storage       = $storage
        storageError  = $storageError
        error         = $null
    }
}

function Get-ServiceCatalog {
    return @(
        @{ Name = 'RpcSs'; Role = 'Required' },
        @{ Name = 'DcomLaunch'; Role = 'Required' },
        @{ Name = 'EventLog'; Role = 'Required' },
        @{ Name = 'Winmgmt'; Role = 'Required' },
        @{ Name = 'Schedule'; Role = 'Required' },
        @{ Name = 'SamSs'; Role = 'Required' },
        @{ Name = 'CryptSvc'; Role = 'Required' },
        @{ Name = 'BFE'; Role = 'Required' },
        @{ Name = 'nsi'; Role = 'Required' },
        @{ Name = 'Power'; Role = 'Required' },
        @{ Name = 'LanmanWorkstation'; Role = 'Required' },
        @{ Name = 'NlaSvc'; Role = 'Required' },
        @{ Name = 'Dnscache'; Role = 'Important' },
        @{ Name = 'Dhcp'; Role = 'Important' },
        @{ Name = 'MpsSvc'; Role = 'Important' },
        @{ Name = 'WinDefend'; Role = 'Important' },
        @{ Name = 'Wcmsvc'; Role = 'Important' },
        @{ Name = 'wuauserv'; Role = 'Important' },
        @{ Name = 'BITS'; Role = 'Important' },
        @{ Name = 'W32Time'; Role = 'Important' },
        @{ Name = 'Audiosrv'; Role = 'Important' },
        @{ Name = 'AudioEndpointBuilder'; Role = 'Important' },
        @{ Name = 'TrustedInstaller'; Role = 'Important' }
    )
}

function Get-CatalogService {
    param([Parameter(Mandatory = $true)][string]$Name)

    $result = [ordered]@{
        found            = $false
        displayName      = $null
        startType        = $null
        delayedAutoStart = $false
        state            = $null
    }
    try {
        $service = Get-CimInstance -ClassName Win32_Service -Filter ("Name='{0}'" -f $Name) -ErrorAction Stop
        if ($null -eq $service) { return $result }
        $result.found = $true
        $result.displayName = [string]$service.DisplayName
        $result.state = [string]$service.State
        $mode = [string]$service.StartMode
        if ($mode -eq 'Auto') { $result.startType = 'Automatic' }
        elseif ($mode -eq 'Manual') { $result.startType = 'Manual' }
        elseif ($mode -eq 'Disabled') { $result.startType = 'Disabled' }
        else { $result.startType = $mode }
        try { $result.delayedAutoStart = [bool]$service.DelayedAutoStart } catch { }
        return $result
    }
    catch {
        try {
            $service = Get-Service -Name $Name -ErrorAction Stop
            $result.found = $true
            $result.displayName = [string]$service.DisplayName
            $result.state = [string]$service.Status
            $result.startType = [string]$service.StartType
        }
        catch { }
        return $result
    }
}

function Get-ServiceHealthStatus {
    param(
        [string]$Role,
        [string]$StartType,
        [bool]$DelayedAutoStart,
        [string]$State,
        [bool]$Found
    )

    if (-not $Found) {
        if ($Role -eq 'Required') { return 'Critical' }
        return 'Warning'
    }
    if ($State -eq 'Running') { return 'Ok' }
    if ($StartType -eq 'Disabled' -and $Role -eq 'Required') { return 'Critical' }
    if ($StartType -eq 'Disabled') { return 'Warning' }
    # Solo el Automatico retrasado baja de Critical a Warning. Un Manual con
    # el flag, como BITS, sigue la regla de su tipo de inicio.
    if ($StartType -eq 'Automatic' -and $DelayedAutoStart) { return 'Warning' }
    # Manual incluye servicios por disparador. NlaSvc en Windows 11 es el caso habitual.
    if ($StartType -eq 'Automatic' -and $Role -eq 'Required') { return 'Critical' }
    if ($StartType -eq 'Automatic') { return 'Warning' }
    if ($Role -eq 'Required') { return 'Warning' }
    return 'Ok'
}

function Get-ServiceStatusNote {
    param(
        [string]$Status,
        [string]$StartType,
        [bool]$DelayedAutoStart,
        [bool]$Found
    )

    if (-not $Found -and $Status -eq 'Critical') { return 'Servicio Required ausente.' }
    if ($StartType -eq 'Disabled' -and $Status -eq 'Critical') { return 'Servicio Required deshabilitado.' }
    if ($StartType -eq 'Automatic' -and $DelayedAutoStart -and $Status -ne 'Ok') { return 'Inicio automatico retrasado. Detenido no es Critical.' }
    if ($StartType -eq 'Manual' -and $Status -eq 'Warning') { return 'Inicio manual o por disparador. Detenido no es Critical.' }
    if ($StartType -eq 'Automatic' -and $Status -eq 'Critical') { return 'Automatico, sin retraso, y detenido.' }
    return $null
}

function Get-ServiceReport {
    $items = New-Object System.Collections.Generic.List[object]
    $problems = New-Object System.Collections.Generic.List[object]
    $statuses = New-Object System.Collections.Generic.List[string]

    foreach ($spec in (Get-ServiceCatalog)) {
        $service = Get-CatalogService -Name ([string]$spec.Name)
        $status = Get-ServiceHealthStatus -Role ([string]$spec.Role) -StartType ([string]$service.startType) -DelayedAutoStart ([bool]$service.delayedAutoStart) -State ([string]$service.state) -Found ([bool]$service.found)
        $statuses.Add($status)
        $startLabel = [string]$service.startType
        if ([bool]$service.delayedAutoStart -and $startLabel -eq 'Automatic') {
            $startLabel = 'Automatic, retrasado'
        }
        $state = [string]$service.state
        if (-not [bool]$service.found) { $state = 'NotFound' }
        $record = [ordered]@{
            name             = [string]$spec.Name
            displayName      = $service.displayName
            role             = [string]$spec.Role
            startType        = $startLabel
            delayedAutoStart = [bool]$service.delayedAutoStart
            state            = $state
            status           = $status
            note             = (Get-ServiceStatusNote -Status $status -StartType ([string]$service.startType) -DelayedAutoStart ([bool]$service.delayedAutoStart) -Found ([bool]$service.found))
        }
        $items.Add($record)
        if ($status -ne 'Ok') { $problems.Add($record) }
    }

    return [ordered]@{
        status   = (Get-WorstStatus -Statuses ($statuses.ToArray()))
        checked  = $items.Count
        problems = $problems
        items    = $items
        error    = $null
    }
}

function Get-EventReport {
    param(
        [int]$Hours,
        [int]$Limit,
        [string[]]$Logs
    )

    $start = (Get-Date).AddHours(-1 * $Hours)
    $found = New-Object System.Collections.Generic.List[object]
    $logErrors = New-Object System.Collections.Generic.List[string]

    foreach ($logName in $Logs) {
        if ([string]::IsNullOrWhiteSpace($logName)) { continue }
        try {
            $events = @(Get-WinEvent -FilterHashtable @{
                LogName   = $logName
                Level     = 1, 2
                StartTime = $start
            } -MaxEvents $Limit -ErrorAction Stop)
            foreach ($event in $events) {
                $message = [string]$event.Message
                if ($message) {
                    $message = ($message -replace '\s+', ' ').Trim()
                    if ($message.Length -gt 400) { $message = $message.Substring(0, 400) }
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
    $ordered = @($found.ToArray() | Sort-Object { $_.timeCreatedUtc } -Descending | Select-Object -First $Limit)
    foreach ($item in $ordered) {
        if ($null -ne $item) { $selected.Add($item) }
    }

    $status = 'Ok'
    if ($logErrors.Count -gt 0) { $status = 'Warning' }
    if ($selected.Count -gt 0 -and $status -eq 'Ok') { $status = 'Warning' }
    $criticalEvents = 0
    foreach ($item in $selected) {
        if ($item.level -eq 'Critical') { $criticalEvents++ }
    }
    if ($criticalEvents -gt 0) { $status = 'Critical' }

    return [ordered]@{
        status       = $status
        hours        = $Hours
        maxEvents    = $Limit
        startLocal   = $start.ToString('yyyy-MM-ddTHH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
        logs         = $Logs
        count        = $selected.Count
        criticalCount = $criticalEvents
        logErrors    = $logErrors
        items        = $selected
        error        = $null
    }
}


function Get-RecentHotfixes {
    param([int]$Limit)
    $items = New-Object System.Collections.Generic.List[object]
    $fixes = @(Get-CimInstance -ClassName Win32_QuickFixEngineering -ErrorAction Stop)
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



function Get-WindowsUpdateReport {
    param([switch]$SkipPending)

    $serviceState = $null
    $serviceStart = $null
    try {
        $service = Get-Service -Name 'wuauserv' -ErrorAction Stop
        $serviceState = [string]$service.Status
        $serviceStart = [string]$service.StartType
    }
    catch {
        $serviceState = 'NotFound'
    }

    $notification = $null
    $serviceEnabled = $null
    $lastSearch = $null
    $lastInstall = $null
    $autoError = $null
    try {
        $auto = New-Object -ComObject Microsoft.Update.AutoUpdate
        try { $serviceEnabled = [bool]$auto.ServiceEnabled } catch { }
        try { $notification = Get-NotificationLevelText -Level $auto.Settings.NotificationLevel } catch { }
        try { $lastSearch = Format-Utc -Value $auto.Results.LastSearchSuccessDate } catch { }
        try { $lastInstall = Format-Utc -Value $auto.Results.LastInstallationSuccessDate } catch { }
    }
    catch {
        $autoError = $_.Exception.Message
    }

    $reboot = Get-RebootPendingReport
    $hotfixes = Get-RecentHotfixes -Limit 8
    $pending = [ordered]@{
        queried = $false
        error   = $null
        updates = (New-Object System.Collections.Generic.List[object])
    }
    $pendingSource = 'skipped'
    if (-not $SkipPending) {
        $pendingSource = 'missing'
        $snapshot = Read-WindowsUpdateSnapshot -MaxAgeHours 20
        if ($null -ne $snapshot) {
            $pendingSource = 'WindowsUpdateSnapshot.json'
            $pending.queried = $true
            $nodes = Get-JsonField -Object $snapshot -Name 'updates'
            if ($null -ne $nodes) {
                foreach ($update in @($nodes)) {
                    if ($null -eq $update) { continue }
                    $pending.updates.Add([ordered]@{
                        title      = (Get-JsonText -Object $update -Name 'title')
                        kb         = (Get-JsonText -Object $update -Name 'kb')
                        severity   = (Get-JsonText -Object $update -Name 'severity')
                        downloaded = (Get-JsonBool -Object $update -Name 'downloaded')
                    })
                }
            }
        }
        else {
            $pending.error = 'Sin consulta propia. No hay WindowsUpdateSnapshot.json vigente. UpdateChecker es quien abre el catalogo.'
        }
    }
    else {
        $pending.error = 'Omitido por -SkipPendingUpdates.'
    }

    $status = 'Ok'
    if ($reboot.pending) { $status = 'Warning' }
    $pendingCount = Get-ItemCount -Value (Get-ReportField -Object $pending -Name 'updates')
    if ($pendingCount -gt 0) { $status = 'Warning' }
    if ($autoError) { $status = 'Warning' }
    if (-not $SkipPending -and -not $pending.queried) { $status = 'Warning' }
    if ($serviceState -eq 'NotFound') { $status = 'Critical' }

    return [ordered]@{
        status              = $status
        serviceState        = $serviceState
        serviceStartType    = $serviceStart
        serviceEnabled      = $serviceEnabled
        notificationLevel   = $notification
        lastSearchSuccess   = $lastSearch
        lastInstallSuccess  = $lastInstall
        rebootPending       = [bool]$reboot.pending
        rebootReasons       = $reboot.reasons
        recentHotfixes      = $hotfixes
        pendingQueried      = [bool]$pending.queried
        pendingSource       = $pendingSource
        pendingError        = $pending.error
        pendingCount        = $pendingCount
        pendingUpdates      = $pending.updates
        error               = $autoError
    }
}

function Get-DriverProblemText {
    param([int]$Code)
    switch ($Code) {
        1 { return 'El dispositivo no esta configurado correctamente.' }
        10 { return 'El dispositivo no puede iniciarse.' }
        12 { return 'No hay recursos suficientes para el dispositivo.' }
        14 { return 'Hace falta reiniciar el equipo.' }
        18 { return 'Reinstale los controladores.' }
        22 { return 'El dispositivo esta deshabilitado.' }
        24 { return 'El dispositivo no esta presente o no funciona.' }
        28 { return 'No hay controladores instalados.' }
        31 { return 'El dispositivo no funciona correctamente.' }
        default { return ("Codigo del administrador de dispositivos: {0}." -f $Code) }
    }
}

function Test-BenignFirmwareDevice {
    param($Device)
    $className = ''
    $deviceName = ''
    $deviceId = ''
    try { $className = [string]$Device.PNPClass } catch { }
    try { $deviceName = [string]$Device.Name } catch { }
    try { $deviceId = [string]$Device.DeviceID } catch { }
    if ($className -eq 'Firmware') { return $true }
    if ($deviceId.StartsWith('UEFI\')) { return $true }
    if ($deviceName.IndexOf('UEFI', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
    return $false
}

function Get-DriverReport {
    $items = New-Object System.Collections.Generic.List[object]
    $statuses = New-Object System.Collections.Generic.List[string]
    $devices = @(Get-CimInstance -ClassName Win32_PnPEntity -ErrorAction Stop)
    foreach ($device in $devices) {
        $code = 0
        try { $code = [int]$device.ConfigManagerErrorCode } catch { continue }
        if ($code -eq 0) { continue }
        $benignFirmware = Test-BenignFirmwareDevice -Device $device
        $status = 'Warning'
        $note = $null
        if ($benignFirmware) {
            $note = 'Firmware o UEFI. El codigo del administrador no se trata como Critical.'
        }
        elseif ($code -in 10, 12, 18, 24, 28, 31) {
            $status = 'Critical'
        }
        $statuses.Add($status)
        $items.Add([ordered]@{
            name         = [string]$device.Name
            deviceId     = [string]$device.DeviceID
            pnpClass     = [string]$device.PNPClass
            errorCode    = $code
            errorText    = (Get-DriverProblemText -Code $code)
            manufacturer = [string]$device.Manufacturer
            status       = $status
            note         = $note
        })
    }

    $overall = 'Ok'
    if ($items.Count -gt 0) { $overall = Get-WorstStatus -Statuses ($statuses.ToArray()) }
    return [ordered]@{
        status       = $overall
        problemCount = $items.Count
        items        = $items
        error        = $null
    }
}

function Get-ObservedOverallStatus {
    param([string[]]$Statuses)
    $critical = $false
    $other = $false
    if ($Statuses) {
        foreach ($status in $Statuses) {
            if ([string]::IsNullOrWhiteSpace($status) -or $status -eq 'Ok') { continue }
            if ($status -eq 'Critical') { $critical = $true }
            else { $other = $true }
        }
    }
    if ($critical) { return 'Critical' }
    if ($other) { return 'Warning' }
    return 'Ok'
}

function Get-CriticalRuleText {
    $rules = New-Object System.Collections.Generic.List[string]
    [void]$rules.Add('Critical: servicio Required ausente, deshabilitado, o Automatico sin DelayedAutoStart y detenido.')
    [void]$rules.Add('No es Critical: Required en Manual y detenido, como NlaSvc. En Windows 11 arranca por disparador. Queda en Warning.')
    [void]$rules.Add('No es Critical: Automatico con DelayedAutoStart y detenido. Queda en Warning.')
    [void]$rules.Add('Critical: memoria al 95% o mas, o un volumen al 90% o mas. La CPU al 90% o mas es Warning.')
    [void]$rules.Add('Critical: al menos un evento de nivel 1 en la ventana. Los eventos de nivel 2 son Warning.')
    [void]$rules.Add('Critical: dispositivo con codigo 10, 12, 18, 24, 28 o 31. Firmware, clase Firmware o id UEFI\ queda en Warning.')
    [void]$rules.Add('Critical: wuauserv no existe. Un reinicio pendiente o actualizaciones pendientes son Warning.')
    [void]$rules.Add('Una seccion que no se pudo leer queda Error en su bloque y Warning en el estado global. El codigo 2 exige un Critical real.')
    return ,$rules
}

function Get-OperatingSystemReport {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
    return [ordered]@{
        status       = 'Ok'
        caption      = [string]$os.Caption
        version      = [string]$os.Version
        build        = [string]$os.BuildNumber
        architecture = [string]$os.OSArchitecture
        manufacturer = [string]$cs.Manufacturer
        model        = [string]$cs.Model
        error        = $null
    }
}

function New-HtmlRows {
    param(
        [string[]]$Headers,
        $Records,
        [scriptblock]$CellScript
    )
    $head = New-Object System.Text.StringBuilder
    [void]$head.Append('<tr>')
    foreach ($header in $Headers) {
        [void]$head.Append('<th>')
        [void]$head.Append((ConvertTo-HtmlText -Text $header))
        [void]$head.Append('</th>')
    }
    [void]$head.Append('</tr>')
    $body = New-Object System.Text.StringBuilder
    $count = 0
    if ($null -ne $Records) {
        foreach ($record in $Records) {
            $count++
            $cells = & $CellScript $record
            [void]$body.Append('<tr>')
            foreach ($cell in $cells) {
                [void]$body.Append('<td>')
                [void]$body.Append($cell)
                [void]$body.Append('</td>')
            }
            [void]$body.AppendLine('</tr>')
        }
    }
    if ($count -eq 0) {
        [void]$body.Append('<tr><td colspan="')
        [void]$body.Append($Headers.Count)
        [void]$body.Append('">Sin datos</td></tr>')
    }
    return [ordered]@{ head = $head.ToString(); body = $body.ToString() }
}

function ConvertTo-HealthHtml {
    param($Report)

    $summary = Get-ReportField -Object $Report -Name 'summary'
    $status = [string](Get-ReportField -Object $summary -Name 'overallStatus')
    $findingItems = Get-ReportField -Object $summary -Name 'findings'
    $hostname = [string](Get-ReportField -Object $Report -Name 'hostname')
    $generatedAt = [string](Get-ReportField -Object $Report -Name 'generatedAt')
    $osCaption = [string](Get-ReportField -Object (Get-ReportField -Object $Report -Name 'operatingSystem') -Name 'caption')
    $uptimeText = [string](Get-ReportField -Object (Get-ReportField -Object $Report -Name 'uptime') -Name 'uptimeText')
    $eventCountText = [string](Get-ReportField -Object (Get-ReportField -Object $Report -Name 'events') -Name 'count')
    $diskRecords = Get-ReportField -Object (Get-ReportField -Object $Report -Name 'disks') -Name 'logical'
    $serviceRecords = Get-ReportField -Object (Get-ReportField -Object $Report -Name 'services') -Name 'problems'
    $eventRecords = Get-ReportField -Object (Get-ReportField -Object $Report -Name 'events') -Name 'items'
    $driverRecords = Get-ReportField -Object (Get-ReportField -Object $Report -Name 'drivers') -Name 'items'
    $updateReport = Get-ReportField -Object $Report -Name 'windowsUpdate'
    $updateRecords = Get-ReportField -Object $updateReport -Name 'pendingUpdates'
    $wuService = [string](Get-ReportField -Object $updateReport -Name 'serviceState')
    $wuRebootText = [string](Get-ReportField -Object $updateReport -Name 'rebootPending')
    $wuInstall = [string](Get-ReportField -Object $updateReport -Name 'lastInstallSuccess')
    $memoryReport = Get-ReportField -Object $Report -Name 'memory'
    $cpuLoadValue = Get-ReportField -Object (Get-ReportField -Object $Report -Name 'cpu') -Name 'loadPercent'

    $findings = New-Object System.Text.StringBuilder
    if ($findingItems) {
    foreach ($finding in $findingItems) {
        [void]$findings.Append('<li>')
        [void]$findings.Append((ConvertTo-HtmlText -Text ([string]$finding)))
        [void]$findings.Append('</li>')
    }
    }

    $diskRows = New-HtmlRows -Headers @('Unidad', 'Sistema', 'Tamano', 'Libre', 'Uso', 'Estado') -Records $diskRecords -CellScript {
        param($row)
        @(
            (ConvertTo-HtmlText -Text ([string]$row.deviceId)),
            (ConvertTo-HtmlText -Text ([string]$row.fileSystem)),
            (ConvertTo-HtmlText -Text (Format-ByteSize -Bytes ([int64]$row.sizeBytes))),
            (ConvertTo-HtmlText -Text (Format-ByteSize -Bytes ([int64]$row.freeBytes))),
            (ConvertTo-HtmlText -Text ('{0:N2} %' -f $row.usedPercent)),
            (ConvertTo-HtmlText -Text ([string]$row.status))
        )
    }

    $serviceRows = New-HtmlRows -Headers @('Servicio', 'Rol', 'Inicio', 'Estado', 'Salud') -Records $serviceRecords -CellScript {
        param($row)
        @(
            (ConvertTo-HtmlText -Text ([string]$row.name)),
            (ConvertTo-HtmlText -Text ([string]$row.role)),
            (ConvertTo-HtmlText -Text ([string]$row.startType)),
            (ConvertTo-HtmlText -Text ([string]$row.state)),
            (ConvertTo-HtmlText -Text ([string]$row.status))
        )
    }

    $eventRows = New-HtmlRows -Headers @('Hora UTC', 'Registro', 'Nivel', 'Id', 'Origen', 'Mensaje') -Records $eventRecords -CellScript {
        param($row)
        @(
            (ConvertTo-HtmlText -Text ([string]$row.timeCreatedUtc)),
            (ConvertTo-HtmlText -Text ([string]$row.logName)),
            (ConvertTo-HtmlText -Text ([string]$row.level)),
            (ConvertTo-HtmlText -Text ([string]$row.eventId)),
            (ConvertTo-HtmlText -Text ([string]$row.provider)),
            (ConvertTo-HtmlText -Text ([string]$row.message))
        )
    }

    $driverRows = New-HtmlRows -Headers @('Dispositivo', 'Clase', 'Codigo', 'Detalle', 'Salud') -Records $driverRecords -CellScript {
        param($row)
        @(
            (ConvertTo-HtmlText -Text ([string]$row.name)),
            (ConvertTo-HtmlText -Text ([string]$row.pnpClass)),
            (ConvertTo-HtmlText -Text ([string]$row.errorCode)),
            (ConvertTo-HtmlText -Text ([string]$row.errorText)),
            (ConvertTo-HtmlText -Text ([string]$row.status))
        )
    }

    $updateRows = New-HtmlRows -Headers @('Titulo', 'KB', 'Severidad', 'Descargada') -Records $updateRecords -CellScript {
        param($row)
        $downloaded = 'No'
        if ($row.downloaded) { $downloaded = 'Si' }
        @(
            (ConvertTo-HtmlText -Text ([string]$row.title)),
            (ConvertTo-HtmlText -Text ([string]$row.kb)),
            (ConvertTo-HtmlText -Text ([string]$row.severity)),
            (ConvertTo-HtmlText -Text $downloaded)
        )
    }

    $cpuLoad = 'n/d'
    if ($null -ne $cpuLoadValue) { $cpuLoad = ('{0:N1} %' -f $cpuLoadValue) }
    $memoryText = 'n/d'
    $memoryUsedHtml = Get-ReportField -Object $memoryReport -Name 'usedPercent'
    if ($null -ne $memoryUsedHtml) {
        $freeBytesHtml = Get-ReportField -Object $memoryReport -Name 'freeBytes'
        $totalBytesHtml = Get-ReportField -Object $memoryReport -Name 'totalBytes'
        if ($null -eq $freeBytesHtml) { $freeBytesHtml = 0 }
        if ($null -eq $totalBytesHtml) { $totalBytesHtml = 0 }
        $memoryText = ('{0:N2} % ({1} libres de {2})' -f $memoryUsedHtml, (Format-ByteSize -Bytes ([int64]$freeBytesHtml)), (Format-ByteSize -Bytes ([int64]$totalBytesHtml)))
    }

    $html = @"
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Salud del sistema - $(ConvertTo-HtmlText -Text $hostname)</title>
<style>
body { font-family: Segoe UI, sans-serif; margin: 32px; color: #1f2328; background: #f6f8fa; }
h1 { margin-bottom: 4px; }
h2 { margin-top: 28px; font-size: 18px; }
.meta { color: #57606a; margin-bottom: 16px; }
.badge { display: inline-block; padding: 4px 10px; border-radius: 999px; color: #fff; font-weight: 600; }
.Ok { background: #1a7f37; }
.Warning { background: #9a6700; }
.Critical, .Error { background: #cf222e; }
.Unknown { background: #6e7781; }
.cards { display: flex; flex-wrap: wrap; gap: 12px; }
.card { background: #fff; border: 1px solid #d0d7de; border-radius: 8px; padding: 12px 14px; min-width: 180px; }
.card span { display: block; color: #57606a; font-size: 12px; }
table { border-collapse: collapse; width: 100%; background: #fff; }
th, td { border: 1px solid #d0d7de; padding: 6px 8px; text-align: left; vertical-align: top; font-size: 13px; }
th { background: #eaeef2; }
ul { background: #fff; border: 1px solid #d0d7de; padding: 12px 12px 12px 32px; }
</style>
</head>
<body>
<h1>Salud del sistema <span class="badge $status">$(ConvertTo-HtmlText -Text $status)</span></h1>
<p class="meta">$(ConvertTo-HtmlText -Text $hostname) · $(ConvertTo-HtmlText -Text $osCaption) · $(ConvertTo-HtmlText -Text $generatedAt)</p>
<div class="cards">
  <div class="card"><span>CPU</span><strong>$(ConvertTo-HtmlText -Text $cpuLoad)</strong></div>
  <div class="card"><span>Memoria en uso</span><strong>$(ConvertTo-HtmlText -Text $memoryText)</strong></div>
  <div class="card"><span>Encendido</span><strong>$(ConvertTo-HtmlText -Text $uptimeText)</strong></div>
  <div class="card"><span>Eventos</span><strong>$(ConvertTo-HtmlText -Text $eventCountText)</strong></div>
</div>
<h2>Hallazgos</h2>
<ul>
$($findings.ToString())
</ul>
<h2>Discos</h2>
<table><thead>$($diskRows.head)</thead><tbody>$($diskRows.body)</tbody></table>
<h2>Servicios que requieren atencion</h2>
<table><thead>$($serviceRows.head)</thead><tbody>$($serviceRows.body)</tbody></table>
<h2>Windows Update</h2>
<p>Servicio: $(ConvertTo-HtmlText -Text $wuService). Reinicio pendiente: $(ConvertTo-HtmlText -Text $wuRebootText). Ultima instalacion correcta: $(ConvertTo-HtmlText -Text $wuInstall).</p>
<table><thead>$($updateRows.head)</thead><tbody>$($updateRows.body)</tbody></table>
<h2>Controladores con problema</h2>
<table><thead>$($driverRows.head)</thead><tbody>$($driverRows.body)</tbody></table>
<h2>Eventos Critical y Error</h2>
<table><thead>$($eventRows.head)</thead><tbody>$($eventRows.body)</tbody></table>
</body>
</html>
"@
    return $html
}

function Invoke-HealthSection {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][scriptblock]$Action
    )
    try {
        $value = & $Action
        $sectionStatus = Get-ReportField -Object $value -Name 'status'
        Write-Log -Level INFO -Event 'section_completed' -Data @{
            section = $Name
            status  = [string]$sectionStatus
        }
        return $value
    }
    catch {
        $message = $_.Exception.Message
        Write-Log -Level ERROR -Event 'section_failed' -Data @{
            section = $Name
            message = $message
        }
        Write-Warning ("{0}: {1}" -f $Name, $message)
        $failed = Get-SectionError -Message $message
        return $failed
    }
}

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }

    $repoRoot = Get-RepositoryRoot
    $stamp = [DateTime]::Now.ToString('yyyyMMdd-HHmmss')
    if ([string]::IsNullOrWhiteSpace($ReportPath)) {
        $ReportPath = Join-Path (Join-Path $repoRoot 'reports') ("SystemHealth-{0}.json" -f $stamp)
    }
    if ([string]::IsNullOrWhiteSpace($LogPath)) {
        $LogPath = Join-Path (Join-Path $repoRoot 'logs') 'SystemHealthReport.jsonl'
    }
    $ReportPath = [System.IO.Path]::GetFullPath($ReportPath)
    if ($Html -and [string]::IsNullOrWhiteSpace($HtmlPath)) {
        $HtmlPath = [System.IO.Path]::ChangeExtension($ReportPath, '.html')
    }
    if (-not [string]::IsNullOrWhiteSpace($HtmlPath)) {
        $HtmlPath = [System.IO.Path]::GetFullPath($HtmlPath)
    }

    Initialize-CommonSession -ScriptName 'SystemHealthReport' -LogPath ([System.IO.Path]::GetFullPath($LogPath)) -DryRun $false
    Ensure-Directory -Path (Split-Path -Parent $ReportPath)

    Write-Host 'SystemHealthReport 1.0'
    Write-Host ("Eventos: ultimas {0} h, maximo {1}, registros {2}" -f $EventHours, $MaxEvents, ($EventLog -join ', '))
    Write-Host ("JSON: {0}" -f $ReportPath)
    if ($HtmlPath) { Write-Host ("HTML: {0}" -f $HtmlPath) }

    Write-Log -Level INFO -Event 'run_start' -Data @{
        eventHours          = $EventHours
        maxEvents           = $MaxEvents
        eventLog            = ($EventLog -join ',')
        skipPendingUpdates  = [bool]$SkipPendingUpdates
        html                = [bool]$Html
        reportPath          = $ReportPath
    }

    $operatingSystem = Invoke-HealthSection -Name 'operatingSystem' -Action { Get-OperatingSystemReport }
    $uptime = Invoke-HealthSection -Name 'uptime' -Action { Get-UptimeReport }
    $cpu = Invoke-HealthSection -Name 'cpu' -Action { Get-CpuReport }
    $memory = Invoke-HealthSection -Name 'memory' -Action { Get-MemoryReport }
    $disks = Invoke-HealthSection -Name 'disks' -Action { Get-DiskReport }
    $services = Invoke-HealthSection -Name 'services' -Action { Get-ServiceReport }
    $events = Invoke-HealthSection -Name 'events' -Action {
        Get-EventReport -Hours $EventHours -Limit $MaxEvents -Logs $EventLog
    }
    $windowsUpdate = Invoke-HealthSection -Name 'windowsUpdate' -Action {
        Get-WindowsUpdateReport -SkipPending:$SkipPendingUpdates
    }
    $drivers = Invoke-HealthSection -Name 'drivers' -Action { Get-DriverReport }

    $sectionStatuses = New-Object System.Collections.Generic.List[string]
    foreach ($section in @($cpu, $memory, $disks, $services, $events, $windowsUpdate, $drivers)) {
        $oneStatus = Get-ReportField -Object $section -Name 'status'
        if ($oneStatus) { $sectionStatuses.Add([string]$oneStatus) }
    }
    $overall = Get-ObservedOverallStatus -Statuses ($sectionStatuses.ToArray())

    $cpuStatus = Get-ReportField -Object $cpu -Name 'status'
    $cpuLoadValue = Get-ReportField -Object $cpu -Name 'loadPercent'
    $memoryStatus = Get-ReportField -Object $memory -Name 'status'
    $memoryUsed = Get-ReportField -Object $memory -Name 'usedPercent'
    $uptimeText = Get-ReportField -Object $uptime -Name 'uptimeText'
    $logicalDisks = Get-ReportField -Object $disks -Name 'logical'
    $serviceProblems = Get-ReportField -Object $services -Name 'problems'
    $servicesChecked = Get-ReportField -Object $services -Name 'checked'
    $eventCount = Get-ReportField -Object $events -Name 'count'
    $eventStatus = Get-ReportField -Object $events -Name 'status'
    $eventHoursValue = Get-ReportField -Object $events -Name 'hours'
    $wuStatus = Get-ReportField -Object $windowsUpdate -Name 'status'
    $wuReboot = Get-ReportField -Object $windowsUpdate -Name 'rebootPending'
    $wuPendingQueried = Get-ReportField -Object $windowsUpdate -Name 'pendingQueried'
    $wuPendingCount = Get-ReportField -Object $windowsUpdate -Name 'pendingCount'
    $wuPendingError = Get-ReportField -Object $windowsUpdate -Name 'pendingError'
    $wuServiceState = Get-ReportField -Object $windowsUpdate -Name 'serviceState'
    $driverStatus = Get-ReportField -Object $drivers -Name 'status'
    $driverProblemCount = Get-ReportField -Object $drivers -Name 'problemCount'

    $findings = New-Object System.Collections.Generic.List[string]
    if ($cpuStatus -eq 'Error') { $findings.Add('No se pudo leer la CPU.') }
    elseif ($null -ne $cpuLoadValue) { $findings.Add(("CPU al {0:N1}% ({1})." -f $cpuLoadValue, $cpuStatus)) }
    if ($memoryStatus -eq 'Error') { $findings.Add('No se pudo leer la memoria.') }
    elseif ($null -ne $memoryUsed) { $findings.Add(("Memoria en uso al {0:N2}% ({1})." -f $memoryUsed, $memoryStatus)) }
    if ($uptimeText) { $findings.Add(("Equipo encendido desde hace {0}." -f $uptimeText)) }
    if ($logicalDisks) {
        foreach ($volume in $logicalDisks) {
            $findings.Add(("Volumen {0} al {1:N2}% ({2})." -f (Get-ReportField -Object $volume -Name 'deviceId'), (Get-ReportField -Object $volume -Name 'usedPercent'), (Get-ReportField -Object $volume -Name 'status')))
        }
    }
    $serviceProblemCount = Get-ItemCount -Value $serviceProblems
    if ($serviceProblemCount -gt 0) {
        $findings.Add(("{0} servicio(s) requieren atencion de {1} revisados." -f $serviceProblemCount, $servicesChecked))
    }
    if ($null -ne $eventCount -and $eventStatus -ne 'Error') {
        $findings.Add(("{0} eventos Critical o Error en las ultimas {1} horas." -f $eventCount, $eventHoursValue))
    }
    if ($wuStatus -eq 'Error') { $findings.Add('No se pudo leer Windows Update.') }
    else {
        if ($wuReboot) { $findings.Add('Hay un reinicio pendiente.') }
        if ($wuPendingQueried) {
            $findings.Add(("{0} actualizaciones de software pendientes." -f $wuPendingCount))
        }
        elseif ($wuPendingError) {
            $findings.Add(("Busqueda de actualizaciones pendientes: {0}" -f $wuPendingError))
        }
    }
    if ($driverStatus -eq 'Error') { $findings.Add('No se pudieron leer los dispositivos.') }
    elseif ($null -ne $driverProblemCount) { $findings.Add(("{0} dispositivos con codigo de error de controlador." -f $driverProblemCount)) }
    if ($findings.Count -eq 0) { $findings.Add('No hubo hallazgos.') }

    $stopwatch.Stop()
    $report = [ordered]@{
        schemaVersion   = '1.0'
        generatedAt     = (Get-UtcTimestamp)
        hostname        = [string]$env:COMPUTERNAME
        userName        = [string]$env:USERNAME
        runId           = (Get-RunId)
        thresholds      = [ordered]@{
            warningCpuPercent      = $script:WarningCpuPercent
            warningMemoryPercent   = $script:WarningMemoryPercent
            criticalMemoryPercent  = $script:CriticalMemoryPercent
            warningDiskPercent     = $script:WarningDiskPercent
            criticalDiskPercent    = $script:CriticalDiskPercent
        }
        parameters      = [ordered]@{
            eventHours         = $EventHours
            maxEvents          = $MaxEvents
            eventLog           = $EventLog
            skipPendingUpdates = [bool]$SkipPendingUpdates
            html               = [bool]$Html
        }
        operatingSystem = $operatingSystem
        uptime          = $uptime
        cpu             = $cpu
        memory          = $memory
        disks           = $disks
        services        = $services
        events          = $events
        windowsUpdate   = $windowsUpdate
        drivers         = $drivers
        summary         = [ordered]@{
            overallStatus = $overall
            findings      = $findings
            criticalRules = (Get-CriticalRuleText)
            durationMs    = [int64]$stopwatch.ElapsedMilliseconds
        }
    }

    $json = Write-JsonReport -Path $ReportPath -Report $report
    $utf8Bom = Get-Utf8BomEncoding
    Write-Log -Level INFO -Event 'report_written' -Data @{
        format = 'json'
        path   = $ReportPath
    }

    if ($HtmlPath) {
        Ensure-Directory -Path (Split-Path -Parent $HtmlPath)
        $htmlText = ConvertTo-HealthHtml -Report $report
        [System.IO.File]::WriteAllText($HtmlPath, $htmlText, $utf8Bom)
        Write-Log -Level INFO -Event 'report_written' -Data @{
            format = 'html'
            path   = $HtmlPath
        }
    }

    Write-Host ''
    Write-Host ("Estado global: {0}" -f $overall)
    if ($null -ne $cpuLoadValue) {
        Write-Host ("CPU: {0:N1}%  |  RAM: {1:N2}%  |  Uptime: {2}" -f $cpuLoadValue, $memoryUsed, $uptimeText)
    }
    $diskRows = New-Object System.Collections.Generic.List[object]
    if ($logicalDisks) {
        foreach ($volume in $logicalDisks) {
            $freeBytes = Get-ReportField -Object $volume -Name 'freeBytes'
            if ($null -eq $freeBytes) { $freeBytes = 0 }
            $diskRows.Add([pscustomobject]@{
                Unidad = (Get-ReportField -Object $volume -Name 'deviceId')
                Uso    = ('{0:N2} %' -f (Get-ReportField -Object $volume -Name 'usedPercent'))
                Libre  = (Format-ByteSize -Bytes ([int64]$freeBytes))
                Estado = (Get-ReportField -Object $volume -Name 'status')
            })
        }
    }
    Write-ConsoleTable -Title 'Discos' -Rows $diskRows

    $problemRows = New-Object System.Collections.Generic.List[object]
    if ($serviceProblems) {
        foreach ($service in $serviceProblems) {
            $problemRows.Add([pscustomobject]@{
                Servicio = (Get-ReportField -Object $service -Name 'name')
                Inicio   = (Get-ReportField -Object $service -Name 'startType')
                Estado   = (Get-ReportField -Object $service -Name 'state')
                Salud    = (Get-ReportField -Object $service -Name 'status')
            })
        }
    }
    Write-ConsoleTable -Title 'Servicios que requieren atencion' -Rows $problemRows

    Write-Host ''
    Write-Host ("Eventos Critical/Error: {0} en {1} h." -f $eventCount, $eventHoursValue)
    Write-Host ("Windows Update: servicio {0}; reinicio pendiente: {1}; pendientes consultadas: {2}." -f $wuServiceState, $wuReboot, $wuPendingCount)
    Write-Host ("Controladores con problema: {0}." -f $driverProblemCount)
    Write-Host ("Reporte JSON: {0}" -f $ReportPath)
    if ($HtmlPath) { Write-Host ("Reporte HTML: {0}" -f $HtmlPath) }

    Write-Log -Level INFO -Event 'run_completed' -Data @{
        overallStatus = $overall
        durationMs    = [int64]$stopwatch.ElapsedMilliseconds
        reportPath    = $ReportPath
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
