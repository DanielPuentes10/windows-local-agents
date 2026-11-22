#Requires -Version 5.1
<#
.SYNOPSIS
    Revisa actualizaciones pendientes de Windows Update y de winget.

.DESCRIPTION
    Consulta los dos origenes y escribe un reporte priorizado:

      Critical      solo un hallazgo real: actualizacion obligatoria,
                    severidad MSRC Critical o Important, o categoria
                    Critical Updates. Security Updates sin esa severidad,
                    definiciones, controladores y paquetes winget no son
                    Critical.
      Recommended   actualizaciones normales de Windows y paquetes winget
                    cuya version instalada y disponible se conocen.
      Optional      controladores, BrowseOnly, feature packs, herramientas,
                    versiones Preview y paquetes winget con version desconocida.

    Sin -Apply el script solo informa: no descarga y no instala.
    Es el unico script de la cadena diaria que abre el catalogo de
    Windows Update. El resultado queda en reports\WindowsUpdateSnapshot.json
    y, si tiene menos de -UpdateCacheHours, la siguiente ejecucion lo reutiliza.
    Con -Apply instala unicamente lo no critico que entre en -ApplyScope.
    Nunca instala la prioridad Critical y nunca reinicia el equipo.

    Si ya hay un reinicio pendiente, -Apply no instala nada y el reporte
    lo explica. -Force solo omite la confirmacion SI; sin -Apply no cambia
    el equipo. -WhatIf obliga el modo dry-run.

    Windows Update se busca en un proceso aparte con tiempo limite.
    winget se ejecuta sin --all y sin instalar, salvo en el tramo -Apply.
    No usa Win32_Product.

.PARAMETER WindowsUpdate
    Consulta solo Windows Update. Si no se indica origen, se consultan los dos.

.PARAMETER Winget
    Consulta solo winget. Si no se indica origen, se consultan los dos.

.PARAMETER Apply
    Instala las actualizaciones no criticas del alcance. Sin este
    modificador el script solo informa.

.PARAMETER ApplyScope
    Que no criticas instalaria -Apply.
    NonCritical (predeterminado) incluye Recommended y Optional.
    Recommended u Optional limitan a esa prioridad.
    Critical nunca entra en el alcance.

.PARAMETER Force
    Con -Apply, no pide la confirmacion SI. Pensado para el Programador
    de tareas. Sin -Apply no cambia nada.

.PARAMETER WindowsTimeoutSeconds
    Tiempo maximo de la busqueda de Windows Update. De 15 a 300.
    Predeterminado: 75.

.PARAMETER WingetTimeoutSeconds
    Tiempo maximo del listado de winget. De 15 a 300.
    Predeterminado: 120.

.PARAMETER InstallTimeoutSeconds
    Tiempo maximo de la instalacion de Windows Update cuando hay -Apply.
    De 60 a 1800. Predeterminado: 600.

.PARAMETER PackageTimeoutSeconds
    Tiempo maximo de cada paquete winget cuando hay -Apply.
    De 30 a 600. Predeterminado: 180.

.PARAMETER MaxUpdates
    Maximo de actualizaciones que se guardan por origen. De 1 a 200.
    Predeterminado: 80.

.PARAMETER ReportPath
    JSON del reporte. Si se omite, se crea
    reports\UpdateChecker-yyyyMMdd-HHmmss.json.
    Se escribe tambien en dry-run.

.PARAMETER UpdateCacheHours
    Si ya hay un snapshot completo mas reciente que estas horas, no se
    vuelve a abrir el catalogo. De 1 a 168. Predeterminado: 20.

.PARAMETER RefreshWindowsUpdate
    Ignora el snapshot y consulta el catalogo otra vez. No lo use en la
    cadena diaria: esa cadena debe consultar Windows Update una sola vez.

.PARAMETER LogPath
    JSONL de auditoria. Predeterminado: logs\UpdateChecker.jsonl.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\UpdateChecker.ps1

    Dry-run de Windows Update y winget. No instala nada.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\UpdateChecker.ps1 -Winget -MaxUpdates 40

    Solo lista winget y guarda como mucho 40 paquetes.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\UpdateChecker.ps1 -Apply -ApplyScope Recommended

    Pide confirmacion SI e instala solo la prioridad Recommended.
    Las criticas quedan en el reporte y no se tocan.

.NOTES
    Autor: LocalAgents
    Requiere: Windows PowerShell 5.1 y, para el origen winget, winget.exe.
    La instalacion de Windows Update necesita una consola elevada.
    DailyOptimizer y SystemHealthReport leen el snapshot y no abren el catalogo.
    Codigos de salida:
      0  no hay actualizaciones pendientes ni avisos
      1  hay Warning (recomendadas, opcionales, reinicio pendiente
         o un origen que fallo) y ninguna actualizacion Critical
      2  el catalogo contiene al menos una actualizacion Critical
      3  el script fallo antes de cerrar el reporte
    Helpers compartidos: scripts\Common.psm1.

    Programador de tareas, solo lectura, cada dia a las 08:15.
    No usa -Apply:
      $script = 'C:\ProyectosDeSoftware\LocalAgents\scripts\UpdateChecker.ps1'
      $arg = "-NoProfile -ExecutionPolicy Bypass -File `"$script`""
      $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arg
      $trigger = New-ScheduledTaskTrigger -Daily -At 8:15am
      Register-ScheduledTask -TaskName 'LocalAgents-UpdateChecker' -Action $action -Trigger $trigger -Description 'Revision de actualizaciones en dry-run'

    Para instalar no criticas hace falta decidirlo aparte, con elevacion,
    y conviene un reinicio ya resuelto. Ejemplo, no registrado:
      -Apply -ApplyScope Recommended -Force
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$WindowsUpdate,
    [switch]$Winget,

    [switch]$Apply,
    [switch]$Force,

    [ValidateSet('NonCritical', 'Recommended', 'Optional')]
    [string]$ApplyScope = 'NonCritical',

    [ValidateRange(15, 300)]
    [int]$WindowsTimeoutSeconds = 75,

    [ValidateRange(15, 300)]
    [int]$WingetTimeoutSeconds = 120,

    [ValidateRange(60, 1800)]
    [int]$InstallTimeoutSeconds = 600,

    [ValidateRange(30, 600)]
    [int]$PackageTimeoutSeconds = 180,

    [ValidateRange(1, 200)]
    [int]$MaxUpdates = 80,

    [ValidateRange(1, 168)]
    [int]$UpdateCacheHours = 20,

    [switch]$RefreshWindowsUpdate,

    [string]$ReportPath,

    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -Force -DisableNameChecking

$script:ExitCode = 0
$script:WindowsJob = $null
$script:WindowsTemp = $null
$script:WindowsDeadline = $null
$script:AllowTemp = $null
$script:WindowsSnapshot = $null


function Test-CategoryName {
    param(
        [AllowEmptyString()][string]$Categories,
        [Parameter(Mandatory = $true)][string]$Name
    )
    if ([string]::IsNullOrWhiteSpace($Categories)) { return $false }
    return ($Categories.IndexOf($Name, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
}

function Get-WindowsUpdatePriority {
    param(
        [AllowEmptyString()][string]$Severity,
        [AllowEmptyString()][string]$Categories,
        [bool]$Mandatory,
        [bool]$BrowseOnly,
        [AllowEmptyString()][string]$UpdateType,
        [AllowEmptyString()][string]$Title
    )

    $severityKey = ''
    if (-not [string]::IsNullOrWhiteSpace($Severity)) {
        $severityKey = $Severity.Trim().ToLowerInvariant()
    }

    if ((Test-CategoryName -Categories $Categories -Name 'Definition Updates') -and -not $Mandatory) {
        return [ordered]@{ priority = 'Recommended'; reason = 'Definicion, por ejemplo Defender.' }
    }
    if ($Mandatory) {
        return [ordered]@{ priority = 'Critical'; reason = 'Actualizacion obligatoria.' }
    }
    if ($severityKey -eq 'critical') {
        return [ordered]@{ priority = 'Critical'; reason = 'Severidad MSRC Critical.' }
    }
    if ($severityKey -eq 'important') {
        return [ordered]@{ priority = 'Critical'; reason = 'Severidad MSRC Important.' }
    }
    if (Test-CategoryName -Categories $Categories -Name 'Critical Updates') {
        return [ordered]@{ priority = 'Critical'; reason = 'Categoria Critical Updates.' }
    }
    if (Test-CategoryName -Categories $Categories -Name 'Security Updates') {
        if ($severityKey -eq 'moderate' -or $severityKey -eq 'low' -or [string]::IsNullOrWhiteSpace($severityKey)) {
            return [ordered]@{ priority = 'Recommended'; reason = 'Security Updates sin severidad MSRC alta.' }
        }
        return [ordered]@{ priority = 'Critical'; reason = 'Categoria Security Updates con severidad alta.' }
    }
    if ($BrowseOnly) {
        return [ordered]@{ priority = 'Optional'; reason = 'Marcada como solo exploracion (BrowseOnly).' }
    }
    if ($UpdateType -eq 'Driver') {
        return [ordered]@{ priority = 'Optional'; reason = 'Controlador.' }
    }
    foreach ($name in @('Feature Packs', 'Upgrades', 'Tools', 'Language Packs', 'Guidance')) {
        if (Test-CategoryName -Categories $Categories -Name $name) {
            return [ordered]@{ priority = 'Optional'; reason = ("Categoria {0}." -f $name) }
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($Title) -and $Title.IndexOf('Preview', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
        return [ordered]@{ priority = 'Optional'; reason = 'El titulo indica una version Preview.' }
    }
    return [ordered]@{ priority = 'Recommended'; reason = 'Actualizacion normal de Windows.' }
}

function Get-WingetPriority {
    param(
        [AllowEmptyString()][string]$Version,
        [AllowEmptyString()][string]$Available
    )
    $versionText = ''
    $availableText = ''
    if (-not [string]::IsNullOrWhiteSpace($Version)) { $versionText = $Version.Trim().ToLowerInvariant() }
    if (-not [string]::IsNullOrWhiteSpace($Available)) { $availableText = $Available.Trim().ToLowerInvariant() }
    $unknown = @('unknown', 'desconocido', '')
    $inexact = $false
    if ($versionText.StartsWith('>') -or $versionText.StartsWith('<')) { $inexact = $true }
    if ($availableText.StartsWith('>') -or $availableText.StartsWith('<')) { $inexact = $true }
    if ($inexact -or ($unknown -contains $versionText) -or ($unknown -contains $availableText)) {
        return [ordered]@{ priority = 'Optional'; reason = 'winget no pudo determinar una version exacta.'; includeUnknown = $true }
    }
    return [ordered]@{ priority = 'Recommended'; reason = 'Actualizacion de aplicacion publicada en winget.'; includeUnknown = $false }
}

function Test-ScopeEligible {
    param([Parameter(Mandatory = $true)][string]$Priority)
    if ($Priority -eq 'Critical') { return $false }
    if ($ApplyScope -eq 'Recommended') { return ($Priority -eq 'Recommended') }
    if ($ApplyScope -eq 'Optional') { return ($Priority -eq 'Optional') }
    return ($Priority -eq 'Recommended' -or $Priority -eq 'Optional')
}

function Get-ApplyDecision {
    param(
        [Parameter(Mandatory = $true)][string]$Priority,
        [bool]$IdentifierSafe = $true
    )
    if (-not $IdentifierSafe) {
        return [ordered]@{
            eligible = $false
            reason   = 'Identificador truncado o no seguro; no se instala.'
        }
    }
    if ($Priority -eq 'Critical') {
        return [ordered]@{
            eligible = $false
            reason   = 'Este script no instala actualizaciones criticas.'
        }
    }
    if (-not (Test-ScopeEligible -Priority $Priority)) {
        return [ordered]@{
            eligible = $false
            reason   = 'Fuera del alcance -ApplyScope.'
        }
    }
    return [ordered]@{
        eligible = $true
        reason   = $null
    }
}

function New-UpdateRecord {
    param(
        [Parameter(Mandatory = $true)][string]$Priority,
        [Parameter(Mandatory = $true)][string]$PriorityReason,
        [Parameter(Mandatory = $true)][string]$Source,
        [AllowEmptyString()][string]$Title,
        [AllowEmptyString()][string]$Identifier,
        [AllowEmptyString()][string]$UpdateId,
        [AllowEmptyString()][string]$PackageId,
        [AllowEmptyString()][string]$InstalledVersion,
        [AllowEmptyString()][string]$AvailableVersion,
        [AllowEmptyString()][string]$Severity,
        [AllowEmptyString()][string]$Categories,
        [bool]$RebootRequired,
        [bool]$Downloaded,
        $SizeBytes,
        [bool]$Mandatory,
        [bool]$BrowseOnly,
        [AllowEmptyString()][string]$UpdateType,
        [AllowEmptyString()][string]$SourceName,
        [bool]$IncludeUnknown,
        [bool]$ApplyEligible,
        [AllowEmptyString()][string]$ApplyBlockReason
    )
    $block = $null
    if (-not [string]::IsNullOrWhiteSpace($ApplyBlockReason)) { $block = $ApplyBlockReason }
    return [ordered]@{
        priority         = $Priority
        priorityReason   = $PriorityReason
        source           = $Source
        title            = $Title
        identifier       = $Identifier
        updateId         = $UpdateId
        packageId        = $PackageId
        installedVersion = $InstalledVersion
        availableVersion = $AvailableVersion
        severity         = $Severity
        categories       = $Categories
        rebootRequired   = $RebootRequired
        downloaded       = $Downloaded
        sizeBytes        = $SizeBytes
        mandatory        = $Mandatory
        browseOnly       = $BrowseOnly
        updateType       = $UpdateType
        sourceName       = $SourceName
        includeUnknown   = $IncludeUnknown
        applyEligible    = $ApplyEligible
        applyBlockReason = $block
        applyResult      = $null
        applyError       = $null
    }
}

function Get-PriorityCounts {
    param($Items)
    $critical = 0
    $recommended = 0
    $optional = 0
    $eligible = 0
    if ($null -ne $Items) {
        foreach ($item in $Items) {
            if ($null -eq $item) { continue }
            $priority = [string](Get-DictionaryValue -Object $item -Name 'priority')
            if ($priority -eq 'Critical') { $critical++ }
            elseif ($priority -eq 'Recommended') { $recommended++ }
            elseif ($priority -eq 'Optional') { $optional++ }
            $canApply = Get-DictionaryValue -Object $item -Name 'applyEligible'
            if ($canApply) { $eligible++ }
        }
    }
    return [ordered]@{
        critical     = $critical
        recommended  = $recommended
        optional     = $optional
        applyEligible = $eligible
        total        = ($critical + $recommended + $optional)
    }
}

function New-SkippedSource {
    param([string]$Note)
    return [ordered]@{
        meta  = [ordered]@{
            status   = 'Skipped'
            enabled  = $false
            queried  = $false
            error    = $null
            headline = $Note
            note     = $Note
        }
        items = (New-Object System.Collections.Generic.List[object])
    }
}

function Invoke-Check {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][scriptblock]$Action
    )
    try {
        $value = & $Action
        $meta = Get-DictionaryValue -Object $value -Name 'meta'
        $status = Get-DictionaryValue -Object $meta -Name 'status'
        if ([string]::IsNullOrWhiteSpace([string]$status)) {
            $status = Get-DictionaryValue -Object $value -Name 'status'
        }
        Write-Log -Level INFO -Event 'source_completed' -Data @{
            source = $Name
            status = [string]$status
        }
        return $value
    }
    catch {
        $message = $_.Exception.Message
        Write-Log -Level ERROR -Event 'source_failed' -Data @{
            source  = $Name
            message = $message
        }
        Write-Warning ("{0}: {1}" -f $Name, $message)
        return [ordered]@{
            meta  = [ordered]@{
                status   = 'Error'
                enabled  = $true
                queried  = $false
                error    = $message
                headline = $message
            }
            items = (New-Object System.Collections.Generic.List[object])
        }
    }
}

function Clear-WindowsUpdateJob {
    if ($script:WindowsJob) {
        try { Stop-Job -Job $script:WindowsJob -ErrorAction SilentlyContinue } catch { }
        try { Remove-Job -Job $script:WindowsJob -Force -ErrorAction SilentlyContinue } catch { }
        $script:WindowsJob = $null
    }
    if ($script:WindowsTemp -and (Test-Path -LiteralPath $script:WindowsTemp)) {
        try { [System.IO.File]::Delete($script:WindowsTemp) } catch { }
    }
    $script:WindowsTemp = $null
    if ($script:AllowTemp -and (Test-Path -LiteralPath $script:AllowTemp)) {
        try { [System.IO.File]::Delete($script:AllowTemp) } catch { }
    }
    $script:AllowTemp = $null
}

function Start-WindowsUpdateSearch {
    param(
        [int]$TimeoutSeconds,
        [int]$MaxCount
    )
    $script:WindowsSnapshot = $null
    if (-not $RefreshWindowsUpdate) {
        $cached = Read-WindowsUpdateSnapshot -MaxAgeHours $UpdateCacheHours
        if ($null -ne $cached) {
            $script:WindowsSnapshot = $cached
            Write-Log -Level INFO -Event 'update_snapshot_used' -Data @{
                generatedAt  = (Get-JsonText -Object $cached -Name 'generatedAt')
                maxAgeHours  = $UpdateCacheHours
            }
            return
        }
    }
    $temp = Join-Path $env:TEMP ('UpdateChecker-WU-{0}.json' -f ([guid]::NewGuid().ToString('n')))
    $job = Start-Job -ArgumentList $temp, $MaxCount -ScriptBlock {
        param($OutputPath, $Take)
        Set-StrictMode -Version Latest
        $ErrorActionPreference = 'Stop'
        $clock = [System.Diagnostics.Stopwatch]::StartNew()
        $payload = @{
            ok         = $false
            error      = $null
            searchMode = $null
            truncated  = $false
            durationMs = 0
            updates    = @()
        }
        try {
            $session = New-Object -ComObject Microsoft.Update.Session
            try { $session.ClientApplicationID = 'LocalAgents.UpdateChecker' } catch { }
            $searcher = $session.CreateUpdateSearcher()
            $criteria = 'IsInstalled=0 and IsHidden=0'
            $result = $null
            $mode = 'Online'
            try {
                $searcher.Online = $true
                $result = $searcher.Search($criteria)
            }
            catch {
                $mode = 'Offline'
                $searcher.Online = $false
                $result = $searcher.Search($criteria)
            }
            $payload.searchMode = $mode
            $updates = New-Object System.Collections.Generic.List[object]
            $kept = 0
            foreach ($update in $result.Updates) {
                if ($kept -ge $Take) {
                    $payload.truncated = $true
                    break
                }
                $title = ''
                try { $title = [string]$update.Title } catch { }
                if ([string]::IsNullOrWhiteSpace($title)) { continue }
                $kbList = New-Object System.Collections.Generic.List[string]
                try {
                    foreach ($article in $update.KBArticleIDs) {
                        if ($null -eq $article) { continue }
                        $kbList.Add([string]$article)
                    }
                }
                catch { }
                $categoryList = New-Object System.Collections.Generic.List[string]
                try {
                    if ($null -ne $update.Categories) {
                        foreach ($category in $update.Categories) {
                            if ($null -eq $category) { continue }
                            try { $categoryList.Add([string]$category.Name) } catch { }
                        }
                    }
                }
                catch { }
                $updateId = ''
                try { $updateId = [string]$update.Identity.UpdateID } catch { }
                $downloaded = $false
                $mandatory = $false
                $rebootRequired = $false
                $browseOnly = $false
                $autoSelect = $false
                try { $downloaded = [bool]$update.IsDownloaded } catch { }
                try { $mandatory = [bool]$update.IsMandatory } catch { }
                try { $rebootRequired = [bool]$update.RebootRequired } catch { }
                try { $browseOnly = [bool]$update.BrowseOnly } catch { }
                try { $autoSelect = [bool]$update.AutoSelectOnWebSites } catch { }
                $severity = ''
                try { $severity = [string]$update.MsrcSeverity } catch { }
                $typeText = 'Unknown'
                try {
                    $typeCode = [int]$update.Type
                    if ($typeCode -eq 1) { $typeText = 'Software' }
                    elseif ($typeCode -eq 2) { $typeText = 'Driver' }
                }
                catch { }
                $sizeValue = 0
                $sizeKnown = $false
                try {
                    $rawSize = [int64]$update.MaxDownloadSize
                    if ($rawSize -gt 0 -and $rawSize -le [int]::MaxValue) {
                        $sizeValue = [int]$rawSize
                        $sizeKnown = $true
                    }
                }
                catch { }
                $kept++
                $updates.Add(@{
                    title          = $title
                    kb             = ($kbList.ToArray() -join ',')
                    severity       = $severity
                    categories     = ($categoryList.ToArray() -join ' | ')
                    downloaded     = $downloaded
                    mandatory      = $mandatory
                    rebootRequired = $rebootRequired
                    browseOnly     = $browseOnly
                    autoSelect     = $autoSelect
                    updateType     = $typeText
                    sizeBytes      = $sizeValue
                    sizeKnown      = $sizeKnown
                    updateId       = $updateId
                })
            }
            $payload.ok = $true
            $payload.updates = $updates.ToArray()
        }
        catch {
            $payload.error = $_.Exception.Message
            $payload.updates = @()
        }
        $clock.Stop()
        $payload.durationMs = [int]$clock.ElapsedMilliseconds

        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($item in $payload.updates) {
            if ($null -eq $item) { continue }
            $parts.Add((ConvertTo-Json -InputObject $item -Depth 4 -Compress))
        }
        $updatesJson = '[' + ($parts -join ',') + ']'
        $errorJson = 'null'
        if ($null -ne $payload.error) {
            $errorJson = ConvertTo-Json -InputObject ([string]$payload.error) -Compress
        }
        $modeJson = 'null'
        if ($null -ne $payload.searchMode) {
            $modeJson = ConvertTo-Json -InputObject ([string]$payload.searchMode) -Compress
        }
        $okJson = 'false'
        if ($payload.ok) { $okJson = 'true' }
        $truncatedJson = 'false'
        if ($payload.truncated) { $truncatedJson = 'true' }
        $durationJson = ([int]$payload.durationMs).ToString([System.Globalization.CultureInfo]::InvariantCulture)
        $json = '{"ok":' + $okJson + ',"error":' + $errorJson + ',"searchMode":' + $modeJson + ',"truncated":' + $truncatedJson + ',"durationMs":' + $durationJson + ',"updates":' + $updatesJson + '}'
        $utf8 = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllText($OutputPath, $json, $utf8)
    }
    $script:WindowsJob = $job
    $script:WindowsTemp = $temp
    $script:WindowsDeadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
}

function Complete-WindowsUpdateSearch {
    param([int]$TimeoutSeconds)

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

    $notification = $null
    $autoError = $null
    try {
        $auto = New-Object -ComObject Microsoft.Update.AutoUpdate
        try { $notification = Get-NotificationLevelText -Level $auto.Settings.NotificationLevel } catch { }
    }
    catch {
        $autoError = $_.Exception.Message
    }

    $reboot = Get-RebootPending
    $rawItems = New-Object System.Collections.Generic.List[object]
    $queried = $false
    $errorText = $null
    $truncated = $false
    $searchMode = $null
    $searchMs = $null

    if ($null -ne $script:WindowsSnapshot) {
        $queried = $true
        $searchMode = 'Cache'
        $truncated = Get-JsonBool -Object $script:WindowsSnapshot -Name 'truncated'
        $searchMs = 0
        $cachedError = Get-JsonText -Object $script:WindowsSnapshot -Name 'error'
        if (-not [string]::IsNullOrWhiteSpace($cachedError)) { $errorText = $cachedError }
        $updateNodes = Get-JsonField -Object $script:WindowsSnapshot -Name 'updates'
        if ($null -ne $updateNodes) {
            foreach ($update in @($updateNodes)) {
                if ($null -eq $update) { continue }
                $rawItems.Add([ordered]@{
                    title          = (Get-JsonText -Object $update -Name 'title')
                    kb             = (Get-JsonText -Object $update -Name 'kb')
                    severity       = (Get-JsonText -Object $update -Name 'severity')
                    categories     = (Get-JsonText -Object $update -Name 'categories')
                    downloaded     = (Get-JsonBool -Object $update -Name 'downloaded')
                    mandatory      = (Get-JsonBool -Object $update -Name 'mandatory')
                    rebootRequired = (Get-JsonBool -Object $update -Name 'rebootRequired')
                    browseOnly     = (Get-JsonBool -Object $update -Name 'browseOnly')
                    autoSelect     = (Get-JsonBool -Object $update -Name 'autoSelect')
                    updateType     = (Get-JsonText -Object $update -Name 'updateType')
                    sizeBytes      = (Get-JsonField -Object $update -Name 'sizeBytes')
                    sizeKnown      = (Get-JsonBool -Object $update -Name 'sizeKnown')
                    updateId       = (Get-JsonText -Object $update -Name 'updateId')
                })
            }
        }
    }
    elseif ($script:WindowsJob) {
        $remaining = [int]($script:WindowsDeadline - [DateTime]::UtcNow).TotalSeconds
        if ($remaining -lt 1) { $remaining = 1 }
        $finished = Wait-Job -Job $script:WindowsJob -Timeout $remaining
        if (-not $finished) {
            $errorText = ("Tiempo de espera agotado ({0} s) al buscar Windows Update." -f $TimeoutSeconds)
        }
        else {
            Receive-Job -Job $script:WindowsJob -ErrorAction SilentlyContinue | Out-Null
            if ($script:WindowsTemp -and (Test-Path -LiteralPath $script:WindowsTemp)) {
                $raw = [System.IO.File]::ReadAllText($script:WindowsTemp)
                $parsed = (ConvertFrom-DataJson -Json $raw)
                $searchMode = Get-JsonText -Object $parsed -Name 'searchMode'
                $truncated = Get-JsonBool -Object $parsed -Name 'truncated'
                $durationValue = Get-JsonField -Object $parsed -Name 'durationMs'
                if ($null -ne $durationValue) {
                    try { $searchMs = [int]$durationValue } catch { }
                }
                $updateNodes = Get-JsonField -Object $parsed -Name 'updates'
                if ($null -ne $updateNodes) {
                    foreach ($update in $updateNodes) {
                        if ($null -eq $update) { continue }
                        $rawItems.Add([ordered]@{
                            title          = (Get-JsonText -Object $update -Name 'title')
                            kb             = (Get-JsonText -Object $update -Name 'kb')
                            severity       = (Get-JsonText -Object $update -Name 'severity')
                            categories     = (Get-JsonText -Object $update -Name 'categories')
                            downloaded     = (Get-JsonBool -Object $update -Name 'downloaded')
                            mandatory      = (Get-JsonBool -Object $update -Name 'mandatory')
                            rebootRequired = (Get-JsonBool -Object $update -Name 'rebootRequired')
                            browseOnly     = (Get-JsonBool -Object $update -Name 'browseOnly')
                            autoSelect     = (Get-JsonBool -Object $update -Name 'autoSelect')
                            updateType     = (Get-JsonText -Object $update -Name 'updateType')
                            sizeBytes      = (Get-JsonField -Object $update -Name 'sizeBytes')
                            sizeKnown      = (Get-JsonBool -Object $update -Name 'sizeKnown')
                            updateId       = (Get-JsonText -Object $update -Name 'updateId')
                        })
                    }
                }
                if (Get-JsonBool -Object $parsed -Name 'ok') { $queried = $true }
                $parsedError = Get-JsonText -Object $parsed -Name 'error'
                if (-not [string]::IsNullOrWhiteSpace($parsedError)) { $errorText = $parsedError }
            }
            elseif (-not $errorText) {
                $errorText = 'La busqueda de Windows Update no dejo resultado.'
            }
        }
        Clear-WindowsUpdateJob
    }
    else {
        $errorText = 'La busqueda de Windows Update no llego a iniciarse.'
    }

    $items = New-Object System.Collections.Generic.List[object]
    foreach ($rawItem in $rawItems) {
        $title = [string](Get-DictionaryValue -Object $rawItem -Name 'title')
        $kb = [string](Get-DictionaryValue -Object $rawItem -Name 'kb')
        $severity = [string](Get-DictionaryValue -Object $rawItem -Name 'severity')
        $categories = [string](Get-DictionaryValue -Object $rawItem -Name 'categories')
        $updateType = [string](Get-DictionaryValue -Object $rawItem -Name 'updateType')
        $mandatory = [bool](Get-DictionaryValue -Object $rawItem -Name 'mandatory')
        $browseOnly = [bool](Get-DictionaryValue -Object $rawItem -Name 'browseOnly')
        $ranked = Get-WindowsUpdatePriority -Severity $severity -Categories $categories -Mandatory $mandatory -BrowseOnly $browseOnly -UpdateType $updateType -Title $title
        $updateId = [string](Get-DictionaryValue -Object $rawItem -Name 'updateId')
        $identifier = $kb
        if ([string]::IsNullOrWhiteSpace($identifier)) { $identifier = $updateId }
        $safeId = $true
        if ([string]::IsNullOrWhiteSpace($updateId) -and [string]::IsNullOrWhiteSpace($kb)) { $safeId = $false }
        $decision = Get-ApplyDecision -Priority ([string]$ranked.priority) -IdentifierSafe $safeId
        $sizeBytes = $null
        if (Get-DictionaryValue -Object $rawItem -Name 'sizeKnown') {
            try { $sizeBytes = [int64](Get-DictionaryValue -Object $rawItem -Name 'sizeBytes') } catch { $sizeBytes = $null }
        }
        $items.Add((New-UpdateRecord `
            -Priority ([string]$ranked.priority) `
            -PriorityReason ([string]$ranked.reason) `
            -Source 'WindowsUpdate' `
            -Title $title `
            -Identifier $identifier `
            -UpdateId $updateId `
            -Severity $severity `
            -Categories $categories `
            -RebootRequired ([bool](Get-DictionaryValue -Object $rawItem -Name 'rebootRequired')) `
            -Downloaded ([bool](Get-DictionaryValue -Object $rawItem -Name 'downloaded')) `
            -SizeBytes $sizeBytes `
            -Mandatory $mandatory `
            -BrowseOnly $browseOnly `
            -UpdateType $updateType `
            -ApplyEligible ([bool]$decision.eligible) `
            -ApplyBlockReason ([string]$decision.reason)))
    }

    $counts = Get-PriorityCounts -Items $items
    $status = 'Ok'
    if (-not $queried) { $status = 'Error' }
    elseif ([int]$counts.critical -gt 0) { $status = 'Critical' }
    elseif ($errorText -or $serviceState -eq 'NotFound' -or $reboot.pending -or [int]$counts.total -gt 0) { $status = 'Warning' }

    $headline = 'Sin actualizaciones pendientes de Windows.'
    if (-not $queried) {
        $headline = $errorText
    }
    elseif ([int]$counts.total -gt 0) {
        $headline = ("{0} criticas, {1} recomendadas, {2} opcionales." -f [int]$counts.critical, [int]$counts.recommended, [int]$counts.optional)
    }
    if ($truncated) {
        $headline = ($headline + ' Lista truncada.')
    }
    if ($searchMode -eq 'Cache') {
        $headline = ($headline + ' Reutilizado del snapshot diario.')
    }
    elseif ($queried) {
        $stored = New-Object System.Collections.Generic.List[object]
        $index = 0
        foreach ($rawItem in $rawItems) {
            $copy = [ordered]@{}
            foreach ($key in $rawItem.Keys) { $copy[$key] = $rawItem[$key] }
            if ($index -lt $items.Count) {
                $copy['priority'] = [string](Get-DictionaryValue -Object $items[$index] -Name 'priority')
                $copy['priorityReason'] = [string](Get-DictionaryValue -Object $items[$index] -Name 'priorityReason')
            }
            $stored.Add($copy)
            $index++
        }
        try {
            $snapshotPath = Write-WindowsUpdateSnapshot -Snapshot ([ordered]@{
                schemaVersion = '1.0'
                generatedAt   = (Get-UtcTimestamp)
                source        = 'UpdateChecker'
                catalog       = 'full'
                queried       = $true
                searchMode    = $searchMode
                truncated     = [bool]$truncated
                error         = $errorText
                pendingCount  = (Get-ItemCount -Value $items)
                updates       = $stored
            })
            Write-Log -Level INFO -Event 'update_snapshot_written' -Data @{
                path  = $snapshotPath
                count = (Get-ItemCount -Value $stored)
            }
        }
        catch {
            Write-Log -Level WARN -Event 'update_snapshot_write_failed' -Data @{ message = $_.Exception.Message }
        }
    }

    return [ordered]@{
        meta  = [ordered]@{
            status            = $status
            enabled           = $true
            queried           = $queried
            searchMode        = $searchMode
            truncated         = $truncated
            durationMs        = $searchMs
            serviceState      = $serviceState
            serviceStartType  = $serviceStart
            notificationLevel = $notification
            rebootPending     = [bool]$reboot.pending
            rebootReasons     = $reboot.reasons
            autoError         = $autoError
            counts            = $counts
            error             = $errorText
            headline          = $headline
        }
        items = $items
    }
}

function Get-WingetExecutable {
    $command = Get-Command -Name 'winget.exe' -CommandType Application -ErrorAction SilentlyContinue
    if ($command -and -not [string]::IsNullOrWhiteSpace([string]$command.Source)) {
        return [string]$command.Source
    }
    $local = Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\winget.exe'
    if (Test-Path -LiteralPath $local) { return $local }
    return $null
}

function Get-CleanProcessText {
    param([AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $clean = [regex]::Replace($Text, '\x1B\[[0-9;?]*[ -/]*[@-~]', '')
    $clean = $clean -replace "`r`n", "`n" -replace "`r", "`n"
    return $clean
}

function Test-WingetNoUpdateText {
    param([AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    $patterns = @(
        'No installed package found',
        'No applicable update found',
        'No newer package versions',
        'No available upgrade found',
        'No se encontr',
        'No hay actualizaciones'
    )
    foreach ($pattern in $patterns) {
        if ($Text.IndexOf($pattern, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
    }
    return $false
}

function Get-WingetColumnMap {
    param([Parameter(Mandatory = $true)][string]$Header)
    $specs = @(
        @{ Key = 'name'; Labels = @('Name', 'Nombre') },
        @{ Key = 'id'; Labels = @('Id') },
        @{ Key = 'version'; Labels = @('Version') },
        @{ Key = 'available'; Labels = @('Available', 'Disponible') },
        @{ Key = 'source'; Labels = @('Source', 'Origen') },
        @{ Key = 'pin'; Labels = @('Pin') }
    )
    $found = New-Object System.Collections.Generic.List[object]
    foreach ($spec in $specs) {
        $at = -1
        foreach ($label in $spec.Labels) {
            $index = $Header.IndexOf([string]$label)
            if ($index -ge 0) {
                $at = $index
                break
            }
        }
        if ($at -ge 0) {
            $found.Add([pscustomobject]@{ Key = [string]$spec.Key; Index = $at })
        }
    }
    $ordered = New-Object System.Collections.Generic.List[object]
    foreach ($item in ($found.ToArray() | Sort-Object -Property Index)) {
        $ordered.Add($item)
    }
    return ,$ordered.ToArray()
}

function Get-WingetCell {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Line,
        [int]$Start,
        [int]$Next
    )
    if ($Start -lt 0 -or $Start -ge $Line.Length) { return '' }
    $end = $Line.Length
    if ($Next -gt $Start -and $Next -le $Line.Length) { $end = $Next }
    return $Line.Substring($Start, $end - $Start).Trim()
}

function ConvertFrom-WingetTable {
    param([AllowEmptyString()][string]$Text)
    $rows = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($Text)) { return ,$rows }

    $lines = $Text -split "`n"
    $headerIndex = -1
    for ($index = 0; $index -lt $lines.Length; $index++) {
        $line = [string]$lines[$index]
        $hasName = ($line.IndexOf('Name') -ge 0) -or ($line.IndexOf('Nombre') -ge 0)
        $hasId = $line.IndexOf('Id') -ge 0
        $hasAvailable = ($line.IndexOf('Available') -ge 0) -or ($line.IndexOf('Disponible') -ge 0)
        if ($hasName -and $hasId -and $hasAvailable) {
            $headerIndex = $index
            break
        }
    }
    if ($headerIndex -lt 0) { return ,$rows }

    $columns = Get-WingetColumnMap -Header ([string]$lines[$headerIndex])
    $columnList = New-Object System.Collections.Generic.List[object]
    if ($null -ne $columns) {
        foreach ($column in $columns) {
            if ($null -ne $column) { $columnList.Add($column) }
        }
    }
    if ($columnList.Count -lt 3) { return ,$rows }

    for ($index = $headerIndex + 1; $index -lt $lines.Length; $index++) {
        $line = [string]$lines[$index]
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $trimmed = $line.Trim()
        if ($trimmed -match '^-{5,}$') { continue }
        $cells = @{}
        for ($columnIndex = 0; $columnIndex -lt $columnList.Count; $columnIndex++) {
            $start = [int]$columnList[$columnIndex].Index
            $next = -1
            if (($columnIndex + 1) -lt $columnList.Count) { $next = [int]$columnList[$columnIndex + 1].Index }
            $cells[[string]$columnList[$columnIndex].Key] = Get-WingetCell -Line $line -Start $start -Next $next
        }
        $packageId = ''
        if ($cells.ContainsKey('id')) { $packageId = [string]$cells['id'] }
        if ([string]::IsNullOrWhiteSpace($packageId)) { continue }
        if ($packageId -eq 'Id') { continue }
        $truncatedId = $false
        if ($packageId.IndexOf([char]0x2026) -ge 0 -or $packageId.EndsWith('...')) { $truncatedId = $true }
        elseif ($packageId -notmatch '^[A-Za-z0-9][A-Za-z0-9._+-]{0,200}$') { continue }
        $name = ''
        $version = ''
        $available = ''
        $sourceName = ''
        if ($cells.ContainsKey('name')) { $name = [string]$cells['name'] }
        if ($cells.ContainsKey('version')) { $version = [string]$cells['version'] }
        if ($cells.ContainsKey('available')) { $available = [string]$cells['available'] }
        if ($cells.ContainsKey('source')) { $sourceName = [string]$cells['source'] }
        $rows.Add([ordered]@{
            name         = $name
            id           = $packageId
            version      = $version
            available    = $available
            source       = $sourceName
            truncatedId  = $truncatedId
        })
    }
    return ,$rows
}

function Invoke-NativeProcess {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string]$Arguments,
        [int]$TimeoutSeconds
    )
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $FilePath
    $startInfo.Arguments = $Arguments
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.StandardOutputEncoding = New-Object System.Text.UTF8Encoding $false
    $startInfo.StandardErrorEncoding = New-Object System.Text.UTF8Encoding $false

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    [void]$process.Start()
    $outTask = $process.StandardOutput.ReadToEndAsync()
    $errTask = $process.StandardError.ReadToEndAsync()
    $timeoutMs = $TimeoutSeconds * 1000
    $finished = $process.WaitForExit($timeoutMs)
    $timedOut = $false
    if (-not $finished) {
        $timedOut = $true
        & taskkill.exe /PID $process.Id /T /F 2>$null | Out-Null
        try { [void]$process.WaitForExit(5000) } catch { }
    }
    [void]$outTask.Wait(10000)
    [void]$errTask.Wait(10000)
    $stdout = ''
    $stderr = ''
    if ($outTask.IsCompleted) {
        try { $stdout = [string]$outTask.Result } catch { }
    }
    if ($errTask.IsCompleted) {
        try { $stderr = [string]$errTask.Result } catch { }
    }
    $exitCode = $null
    if (-not $timedOut) {
        try { $exitCode = [int]$process.ExitCode } catch { }
    }
    $clock.Stop()
    try { $process.Dispose() } catch { }
    return [ordered]@{
        timedOut   = $timedOut
        exitCode   = $exitCode
        stdout     = (Get-CleanProcessText -Text $stdout)
        stderr     = (Get-CleanProcessText -Text $stderr)
        durationMs = [int]$clock.ElapsedMilliseconds
    }
}

function Get-WingetSection {
    param([int]$TimeoutSeconds)

    $executable = Get-WingetExecutable
    $items = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($executable)) {
        return [ordered]@{
            meta  = [ordered]@{
                status     = 'Warning'
                enabled    = $true
                queried    = $false
                executable = $null
                exitCode   = $null
                truncated  = $false
                durationMs = 0
                counts     = (Get-PriorityCounts -Items $items)
                error      = 'winget.exe no esta disponible.'
                headline   = 'winget.exe no esta disponible.'
                outputSample = $null
            }
            items = $items
        }
    }

    $arguments = 'upgrade --include-unknown --accept-source-agreements --disable-interactivity'
    $run = Invoke-NativeProcess -FilePath $executable -Arguments $arguments -TimeoutSeconds $TimeoutSeconds
    $stdout = [string]$run.stdout
    $stderr = [string]$run.stderr
    $parsed = ConvertFrom-WingetTable -Text $stdout
    $truncated = $false
    $kept = 0
    foreach ($row in $parsed) {
        if ($null -eq $row) { continue }
        $kept++
        if ($kept -gt $MaxUpdates) {
            $truncated = $true
            continue
        }
        $name = [string](Get-DictionaryValue -Object $row -Name 'name')
        $packageId = [string](Get-DictionaryValue -Object $row -Name 'id')
        $version = [string](Get-DictionaryValue -Object $row -Name 'version')
        $available = [string](Get-DictionaryValue -Object $row -Name 'available')
        $sourceName = [string](Get-DictionaryValue -Object $row -Name 'source')
        $ranked = Get-WingetPriority -Version $version -Available $available
        $safeId = -not [bool](Get-DictionaryValue -Object $row -Name 'truncatedId')
        if ($packageId -notmatch '^[A-Za-z0-9][A-Za-z0-9._+-]{0,200}$') { $safeId = $false }
        $decision = Get-ApplyDecision -Priority ([string]$ranked.priority) -IdentifierSafe $safeId
        $items.Add((New-UpdateRecord `
            -Priority ([string]$ranked.priority) `
            -PriorityReason ([string]$ranked.reason) `
            -Source 'Winget' `
            -Title $name `
            -Identifier $packageId `
            -PackageId $packageId `
            -InstalledVersion $version `
            -AvailableVersion $available `
            -SourceName $sourceName `
            -IncludeUnknown ([bool]$ranked.includeUnknown) `
            -ApplyEligible ([bool]$decision.eligible) `
            -ApplyBlockReason ([string]$decision.reason)))
    }

    $exitCode = $run.exitCode
    $emptyCode = $false
    if ($null -ne $exitCode -and [int]$exitCode -eq -1978335189) { $emptyCode = $true }
    $emptyText = (Test-WingetNoUpdateText -Text $stdout) -or (Test-WingetNoUpdateText -Text $stderr)
    $errorText = $null
    $sample = $null
    $queried = $false
    $status = 'Ok'
    $parsedCount = Get-ItemCount -Value $items

    if ($run.timedOut) {
        $status = 'Error'
        $errorText = ("Tiempo de espera agotado ({0} s) al listar winget." -f $TimeoutSeconds)
    }
    elseif ($parsedCount -gt 0) {
        $queried = $true
        if ($null -ne $exitCode -and [int]$exitCode -ne 0) {
            $status = 'Warning'
            $errorText = ("winget termino con codigo {0} y dejo un listado." -f [int]$exitCode)
        }
    }
    elseif ($emptyCode -or $emptyText -or ($null -ne $exitCode -and [int]$exitCode -eq 0 -and $stdout.Length -lt 80)) {
        $queried = $true
        $status = 'Ok'
    }
    elseif ($null -ne $exitCode -and [int]$exitCode -eq 0) {
        $queried = $false
        $status = 'Warning'
        $errorText = 'No se reconocio la tabla de winget.'
        $sample = Get-ShortText -Text $stdout -Max 800
    }
    else {
        $status = 'Error'
        $codeText = 'desconocido'
        if ($null -ne $exitCode) { $codeText = [string][int]$exitCode }
        $detail = Get-ShortText -Text ($stderr + ' ' + $stdout) -Max 240
        $errorText = ("winget fallo con codigo {0}. {1}" -f $codeText, $detail).Trim()
    }

    $counts = Get-PriorityCounts -Items $items
    if ($queried -and [int]$counts.critical -gt 0) { $status = 'Critical' }
    elseif ($queried -and [int]$counts.total -gt 0 -and $status -eq 'Ok') { $status = 'Warning' }

    $headline = 'Sin actualizaciones pendientes en winget.'
    if (-not [string]::IsNullOrWhiteSpace($errorText) -and -not $queried) {
        $headline = $errorText
    }
    elseif ([int]$counts.total -gt 0) {
        $headline = ("{0} criticas, {1} recomendadas, {2} opcionales." -f [int]$counts.critical, [int]$counts.recommended, [int]$counts.optional)
    }
    if ($truncated) { $headline = ($headline + ' Lista truncada.') }

    return [ordered]@{
        meta  = [ordered]@{
            status       = $status
            enabled      = $true
            queried      = $queried
            executable   = $executable
            exitCode     = $exitCode
            truncated    = $truncated
            durationMs   = [int]$run.durationMs
            counts       = $counts
            error        = $errorText
            headline     = $headline
            outputSample = $sample
            note         = 'El listado no incluye paquetes anclados. No usa --all.'
        }
        items = $items
    }
}

function Test-SafePackageId {
    param([AllowEmptyString()][string]$Id)
    if ([string]::IsNullOrWhiteSpace($Id)) { return $false }
    if ($Id.StartsWith('-')) { return $false }
    if ($Id.Contains('..')) { return $false }
    return $Id -match '^[A-Za-z0-9][A-Za-z0-9._+-]{0,200}$'
}

function Set-SkippedApply {
    param(
        $Items,
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$Reason
    )
    if ($null -eq $Items) { return }
    foreach ($item in $Items) {
        if ($null -eq $item) { continue }
        if ([string](Get-DictionaryValue -Object $item -Name 'source') -ne $Source) { continue }
        $eligible = Get-DictionaryValue -Object $item -Name 'applyEligible'
        if (-not $eligible) { continue }
        $item['applyResult'] = 'Omitida'
        $item['applyError'] = $Reason
    }
}

function Install-NonCriticalWindowsUpdates {
    param($Items)
    if ((Test-DryRun)) {
        throw 'Invariante rota: Windows Update no se instala en dry-run.'
    }
    $selected = New-Object System.Collections.Generic.List[object]
    foreach ($item in $Items) {
        if ($null -eq $item) { continue }
        if ([string](Get-DictionaryValue -Object $item -Name 'source') -ne 'WindowsUpdate') { continue }
        if (-not (Get-DictionaryValue -Object $item -Name 'applyEligible')) { continue }
        $selected.Add($item)
    }
    if ((Get-ItemCount -Value $selected) -eq 0) { return }

    if (-not (Test-IsElevated)) {
        $reason = 'Hace falta una consola elevada para instalar Windows Update.'
        Set-SkippedApply -Items $selected -Source 'WindowsUpdate' -Reason $reason
        Write-Log -Level WARN -Event 'install_skipped' -Data @{ source = 'WindowsUpdate'; reason = $reason }
        return
    }

    $reboot = Get-RebootPending
    if ($reboot.pending) {
        $reason = 'Ya hay un reinicio pendiente. No se instala Windows Update.'
        Set-SkippedApply -Items $selected -Source 'WindowsUpdate' -Reason $reason
        Write-Log -Level WARN -Event 'install_skipped' -Data @{ source = 'WindowsUpdate'; reason = $reason }
        return
    }

    $allow = @{
        ids  = (New-Object System.Collections.Generic.List[string])
        keys = (New-Object System.Collections.Generic.List[string])
    }
    foreach ($item in $selected) {
        $updateId = [string](Get-DictionaryValue -Object $item -Name 'updateId')
        $title = [string](Get-DictionaryValue -Object $item -Name 'title')
        $identifier = [string](Get-DictionaryValue -Object $item -Name 'identifier')
        if (-not [string]::IsNullOrWhiteSpace($updateId)) { $allow.ids.Add($updateId) }
        $allow.keys.Add(($title + '|' + $identifier))
    }
    $allowFile = Join-Path $env:TEMP ('UpdateChecker-Allow-{0}.json' -f ([guid]::NewGuid().ToString('n')))
    $allowJson = (ConvertTo-DataJson -Value $allow)
    $utf8 = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($allowFile, $allowJson, $utf8)
    $script:AllowTemp = $allowFile

    $resultFile = Join-Path $env:TEMP ('UpdateChecker-Install-{0}.json' -f ([guid]::NewGuid().ToString('n')))
    $script:WindowsTemp = $resultFile
    Write-Log -Level INFO -Event 'install_started' -Data @{
        source = 'WindowsUpdate'
        count  = (Get-ItemCount -Value $selected)
    }
    $job = Start-Job -ArgumentList $allowFile, $resultFile -ScriptBlock {
        param($AllowPath, $OutputPath)
        Set-StrictMode -Version Latest
        $ErrorActionPreference = 'Stop'

        function Test-BlockedUpdate {
            param(
                [AllowEmptyString()][string]$Severity,
                [AllowEmptyString()][string]$Categories,
                [bool]$Mandatory
            )
            if ($Mandatory) { return $true }
            $severityKey = ''
            if (-not [string]::IsNullOrWhiteSpace($Severity)) { $severityKey = $Severity.Trim().ToLowerInvariant() }
            $isDefinition = $false
            if (-not [string]::IsNullOrWhiteSpace($Categories) -and $Categories.IndexOf('Definition Updates', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $isDefinition = $true
            }
            if ($isDefinition) { return $false }
            if ($severityKey -eq 'critical' -or $severityKey -eq 'important') { return $true }
            if ([string]::IsNullOrWhiteSpace($Categories)) { return $false }
            if ($Categories.IndexOf('Critical Updates', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
            if ($Categories.IndexOf('Security Updates', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                if ($severityKey -eq 'moderate' -or $severityKey -eq 'low' -or [string]::IsNullOrWhiteSpace($severityKey)) { return $false }
                return $true
            }
            return $false
        }

        function Get-ResultName {
            param([int]$Code)
            switch ($Code) {
                2 { return 'Installed' }
                3 { return 'InstalledWithErrors' }
                4 { return 'Failed' }
                5 { return 'Aborted' }
                default { return 'Failed' }
            }
        }

        $results = New-Object System.Collections.Generic.List[object]
        $errorText = $null
        try {
            Add-Type -AssemblyName System.Web.Extensions
            $serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
            $serializer.MaxJsonLength = [int]::MaxValue
            $allowRaw = [System.IO.File]::ReadAllText($AllowPath)
            $allow = $serializer.DeserializeObject($allowRaw)
            $allowedIds = @{}
            $allowedKeys = @{}
            if ($null -ne $allow -and $allow.ContainsKey('ids') -and $null -ne $allow['ids']) {
                foreach ($id in $allow['ids']) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$id)) { $allowedIds[[string]$id] = $true }
                }
            }
            if ($null -ne $allow -and $allow.ContainsKey('keys') -and $null -ne $allow['keys']) {
                foreach ($key in $allow['keys']) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$key)) { $allowedKeys[[string]$key] = $true }
                }
            }

            $session = New-Object -ComObject Microsoft.Update.Session
            try { $session.ClientApplicationID = 'LocalAgents.UpdateChecker' } catch { }
            $searcher = $session.CreateUpdateSearcher()
            $result = $null
            try {
                $searcher.Online = $true
                $result = $searcher.Search('IsInstalled=0 and IsHidden=0')
            }
            catch {
                $searcher.Online = $false
                $result = $searcher.Search('IsInstalled=0 and IsHidden=0')
            }

            $toInstall = New-Object -ComObject Microsoft.Update.UpdateColl
            foreach ($update in $result.Updates) {
                $title = ''
                $severity = ''
                $updateId = ''
                $mandatory = $false
                try { $title = [string]$update.Title } catch { }
                try { $severity = [string]$update.MsrcSeverity } catch { }
                try { $updateId = [string]$update.Identity.UpdateID } catch { }
                try { $mandatory = [bool]$update.IsMandatory } catch { }
                $categoryList = New-Object System.Collections.Generic.List[string]
                try {
                    foreach ($category in $update.Categories) {
                        if ($null -eq $category) { continue }
                        try { $categoryList.Add([string]$category.Name) } catch { }
                    }
                }
                catch { }
                $categories = ($categoryList.ToArray() -join ' | ')
                $kbList = New-Object System.Collections.Generic.List[string]
                try {
                    foreach ($article in $update.KBArticleIDs) {
                        if ($null -eq $article) { continue }
                        $kbList.Add([string]$article)
                    }
                }
                catch { }
                $kb = ($kbList.ToArray() -join ',')
                $key = $title + '|' + $kb
                if ([string]::IsNullOrWhiteSpace($kb)) { $key = $title + '|' + $updateId }
                $allowed = $false
                if (-not [string]::IsNullOrWhiteSpace($updateId) -and $allowedIds.ContainsKey($updateId)) { $allowed = $true }
                if ($allowedKeys.ContainsKey($key)) { $allowed = $true }
                if (-not $allowed) { continue }
                if (Test-BlockedUpdate -Severity $severity -Categories $categories -Mandatory $mandatory) {
                    $results.Add(@{
                        updateId = $updateId
                        key      = $key
                        outcome  = 'Blocked'
                        hresult  = 0
                    })
                    continue
                }
                try {
                    if (-not [bool]$update.EulaAccepted) { [void]$update.AcceptEula() }
                }
                catch { }
                [void]$toInstall.Add($update)
            }

            if ($toInstall.Count -gt 0) {
                $downloader = $session.CreateUpdateDownloader()
                $downloader.Updates = $toInstall
                $downloadResult = $downloader.Download()
                $ready = New-Object -ComObject Microsoft.Update.UpdateColl
                for ($index = 0; $index -lt $toInstall.Count; $index++) {
                    $one = $downloadResult.GetUpdateResult($index)
                    $code = [int]$one.ResultCode
                    $current = $toInstall.Item($index)
                    $already = $false
                    try { $already = [bool]$current.IsDownloaded } catch { }
                    if ($code -eq 2 -or $code -eq 3 -or $already) {
                        [void]$ready.Add($current)
                    }
                    else {
                        $failedId = ''
                        try { $failedId = [string]$current.Identity.UpdateID } catch { }
                        $results.Add(@{
                            updateId = $failedId
                            key      = ''
                            outcome  = 'DownloadFailed'
                            hresult  = [int64]$one.HResult
                        })
                    }
                }
                if ($ready.Count -gt 0) {
                    $installer = $session.CreateUpdateInstaller()
                    $installer.Updates = $ready
                    $installer.AllowSourcePrompts = $false
                    try { $installer.ForceQuiet = $true } catch { }
                    $installResult = $installer.Install()
                    for ($index = 0; $index -lt $ready.Count; $index++) {
                        $one = $installResult.GetUpdateResult($index)
                        $current = $ready.Item($index)
                        $installedId = ''
                        try { $installedId = [string]$current.Identity.UpdateID } catch { }
                        $results.Add(@{
                            updateId = $installedId
                            key      = ''
                            outcome  = (Get-ResultName -Code ([int]$one.ResultCode))
                            hresult  = [int64]$one.HResult
                        })
                    }
                }
            }
        }
        catch {
            $errorText = $_.Exception.Message
        }

        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($item in $results) {
            if ($null -eq $item) { continue }
            $parts.Add((ConvertTo-Json -InputObject $item -Depth 3 -Compress))
        }
        $errorJson = 'null'
        if ($null -ne $errorText) { $errorJson = ConvertTo-Json -InputObject ([string]$errorText) -Compress }
        $json = '{"error":' + $errorJson + ',"results":[' + ($parts -join ',') + ']}'
        $utf8 = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllText($OutputPath, $json, $utf8)
    }
    $script:WindowsJob = $job
    $finished = Wait-Job -Job $job -Timeout $InstallTimeoutSeconds
    if (-not $finished) {
        $reason = ("Tiempo de espera agotado ({0} s) al instalar Windows Update." -f $InstallTimeoutSeconds)
        Set-SkippedApply -Items $selected -Source 'WindowsUpdate' -Reason $reason
        Write-Log -Level ERROR -Event 'install_failed' -Data @{ source = 'WindowsUpdate'; message = $reason }
        Clear-WindowsUpdateJob
        return
    }
    Receive-Job -Job $job -ErrorAction SilentlyContinue | Out-Null
    $parsedError = $null
    $resultNodes = $null
    if (Test-Path -LiteralPath $resultFile) {
        $raw = [System.IO.File]::ReadAllText($resultFile)
        $parsed = (ConvertFrom-DataJson -Json $raw)
        $parsedError = Get-JsonText -Object $parsed -Name 'error'
        $resultNodes = Get-JsonField -Object $parsed -Name 'results'
    }
    else {
        $parsedError = 'La instalacion no dejo resultado.'
    }

    $byId = @{}
    if ($null -ne $resultNodes) {
        foreach ($node in $resultNodes) {
            $id = Get-JsonText -Object $node -Name 'updateId'
            if (-not [string]::IsNullOrWhiteSpace($id)) { $byId[$id] = $node }
        }
    }
    foreach ($item in $selected) {
        $updateId = [string](Get-DictionaryValue -Object $item -Name 'updateId')
        if (-not [string]::IsNullOrWhiteSpace($parsedError) -and -not $byId.ContainsKey($updateId)) {
            $item['applyResult'] = 'Error'
            $item['applyError'] = $parsedError
            continue
        }
        if (-not $byId.ContainsKey($updateId)) {
            $item['applyResult'] = 'Error'
            $item['applyError'] = 'No hubo resultado para esta actualizacion.'
            continue
        }
        $node = $byId[$updateId]
        $item['applyResult'] = Get-JsonText -Object $node -Name 'outcome'
        $hresult = Get-JsonField -Object $node -Name 'hresult'
        if ($item['applyResult'] -eq 'Installed') {
            $item['applyError'] = $null
        }
        else {
            $item['applyError'] = ("HRESULT {0}" -f $hresult)
        }
    }
    Write-Log -Level INFO -Event 'install_finished' -Data @{
        source = 'WindowsUpdate'
        error  = $parsedError
    }
    Clear-WindowsUpdateJob
}

function Install-NonCriticalWingetUpdates {
    param($Items)
    if ((Test-DryRun)) {
        throw 'Invariante rota: winget no se instala en dry-run.'
    }
    $executable = Get-WingetExecutable
    if ([string]::IsNullOrWhiteSpace($executable)) {
        Set-SkippedApply -Items $Items -Source 'Winget' -Reason 'winget.exe no esta disponible.'
        return
    }
    $reboot = Get-RebootPending
    if ($reboot.pending) {
        $reason = 'Ya hay un reinicio pendiente. No se instalan paquetes winget.'
        Set-SkippedApply -Items $Items -Source 'Winget' -Reason $reason
        Write-Log -Level WARN -Event 'install_skipped' -Data @{ source = 'Winget'; reason = $reason }
        return
    }
    foreach ($item in $Items) {
        if ($null -eq $item) { continue }
        if ([string](Get-DictionaryValue -Object $item -Name 'source') -ne 'Winget') { continue }
        if (-not (Get-DictionaryValue -Object $item -Name 'applyEligible')) { continue }
        $packageId = [string](Get-DictionaryValue -Object $item -Name 'packageId')
        if (-not (Test-SafePackageId -Id $packageId)) {
            $item['applyResult'] = 'Omitida'
            $item['applyError'] = 'Identificador de paquete no seguro.'
            continue
        }
        $arguments = 'upgrade --id ' + $packageId + ' --exact --silent --accept-package-agreements --accept-source-agreements --disable-interactivity'
        if (Get-DictionaryValue -Object $item -Name 'includeUnknown') {
            $arguments = $arguments + ' --include-unknown'
        }
        Write-Log -Level INFO -Event 'install_started' -Data @{ source = 'Winget'; packageId = $packageId }
        $run = Invoke-NativeProcess -FilePath $executable -Arguments $arguments -TimeoutSeconds $PackageTimeoutSeconds
        if ($run.timedOut) {
            $item['applyResult'] = 'Error'
            $item['applyError'] = ("Tiempo de espera agotado ({0} s)." -f $PackageTimeoutSeconds)
        }
        elseif ($null -ne $run.exitCode -and [int]$run.exitCode -eq 0) {
            $item['applyResult'] = 'Installed'
            $item['applyError'] = $null
        }
        else {
            $codeText = 'desconocido'
            if ($null -ne $run.exitCode) { $codeText = [string][int]$run.exitCode }
            $item['applyResult'] = 'Error'
            $item['applyError'] = ("winget {0}. {1}" -f $codeText, (Get-ShortText -Text ([string]$run.stderr + ' ' + [string]$run.stdout) -Max 180)).Trim()
        }
        Write-Log -Level INFO -Event 'install_finished' -Data @{
            source    = 'Winget'
            packageId = $packageId
            result    = [string]$item['applyResult']
        }
    }
}

function Confirm-UpdateApply {
    param(
        [int]$WindowsRecommended,
        [int]$WindowsOptional,
        [int]$WingetRecommended,
        [int]$WingetOptional,
        [int]$CriticalCount
    )
    if ((Test-DryRun)) { return $false }
    if ($Force) {
        Write-Log -Level WARN -Event 'confirmation_forced' -Data @{
            windowsRecommended = $WindowsRecommended
            windowsOptional    = $WindowsOptional
            wingetRecommended  = $WingetRecommended
            wingetOptional     = $WingetOptional
        }
        return $true
    }
    if (-not (Test-InteractiveSession)) {
        throw 'La consola no es interactiva. No se instalo nada. Repita con -Apply -Force para una tarea programada.'
    }
    Write-Host ''
    Write-Host 'Se instalarian solo actualizaciones no criticas:'
    Write-Host ("- Windows Update recomendadas: {0}" -f $WindowsRecommended)
    Write-Host ("- Windows Update opcionales: {0}" -f $WindowsOptional)
    Write-Host ("- winget recomendadas: {0}" -f $WingetRecommended)
    Write-Host ("- winget opcionales: {0}" -f $WingetOptional)
    Write-Host ("Las criticas ({0}) no se instalan. El equipo no se reinicia solo." -f $CriticalCount)
    $answer = Read-Host 'Escriba SI en mayusculas para continuar'
    $accepted = ($answer -eq 'SI')
    Write-Log -Level INFO -Event 'confirmation_answer' -Data @{ accepted = $accepted }
    return $accepted
}

function Get-SourceMeta {
    param($Section)
    $meta = Get-DictionaryValue -Object $Section -Name 'meta'
    if ($null -eq $meta) { return $Section }
    return $meta
}

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch { }

    if ($WhatIfPreference) { $Apply = $false }
    Set-DryRunMode -DryRun (-not [bool]$Apply)

    $runAll = -not ($WindowsUpdate -or $Winget)
    $doWindows = $runAll -or [bool]$WindowsUpdate
    $doWinget = $runAll -or [bool]$Winget

    $repoRoot = Get-RepositoryRoot
    $stamp = [DateTime]::Now.ToString('yyyyMMdd-HHmmss')
    if ([string]::IsNullOrWhiteSpace($ReportPath)) {
        $ReportPath = Join-Path (Join-Path $repoRoot 'reports') ("UpdateChecker-{0}.json" -f $stamp)
    }
    if ([string]::IsNullOrWhiteSpace($LogPath)) {
        $LogPath = Join-Path (Join-Path $repoRoot 'logs') 'UpdateChecker.jsonl'
    }
    $ReportPath = [System.IO.Path]::GetFullPath($ReportPath)
    Initialize-CommonSession -ScriptName 'UpdateChecker' -LogPath ([System.IO.Path]::GetFullPath($LogPath)) -DryRun (Test-DryRun)
    Ensure-Directory -Path (Split-Path -Parent $ReportPath)

    $modeText = 'dry-run'
    if (-not (Test-DryRun)) { $modeText = 'apply' }
    $elevated = Test-IsElevated

    Write-Host 'UpdateChecker 1.0'
    Write-Host ("Modo: {0}" -f $modeText)
    Write-Host ("Origenes: Windows Update={0}; winget={1}" -f $doWindows, $doWinget)
    Write-Host ("Alcance si se usara -Apply: {0}. Elevado: {1}." -f $ApplyScope, $elevated)

    Write-Log -Level INFO -Event 'run_start' -Data @{
        dryRun                 = [bool](Test-DryRun)
        windowsUpdate          = $doWindows
        winget                 = $doWinget
        applyScope             = $ApplyScope
        windowsTimeoutSeconds  = $WindowsTimeoutSeconds
        wingetTimeoutSeconds   = $WingetTimeoutSeconds
        maxUpdates             = $MaxUpdates
        elevated               = $elevated
    }

    if ($doWindows) {
        Start-WindowsUpdateSearch -TimeoutSeconds $WindowsTimeoutSeconds -MaxCount $MaxUpdates
    }

    $wingetSection = New-SkippedSource -Note 'Origen winget no solicitado.'
    if ($doWinget) {
        $wingetSection = Invoke-Check -Name 'winget' -Action { Get-WingetSection -TimeoutSeconds $WingetTimeoutSeconds }
    }

    $windowsSection = New-SkippedSource -Note 'Origen Windows Update no solicitado.'
    if ($doWindows) {
        $windowsSection = Invoke-Check -Name 'windowsUpdate' -Action { Complete-WindowsUpdateSearch -TimeoutSeconds $WindowsTimeoutSeconds }
    }

    $updates = New-Object System.Collections.Generic.List[object]
    foreach ($bucket in @('Critical', 'Recommended', 'Optional')) {
        foreach ($section in @($windowsSection, $wingetSection)) {
            $sectionItems = Get-DictionaryValue -Object $section -Name 'items'
            if ($null -eq $sectionItems) { continue }
            foreach ($item in $sectionItems) {
                if ($null -eq $item) { continue }
                if ([string](Get-DictionaryValue -Object $item -Name 'priority') -eq $bucket) {
                    $updates.Add($item)
                }
            }
        }
    }

    $counts = Get-PriorityCounts -Items $updates
    $windowsMeta = Get-SourceMeta -Section $windowsSection
    $wingetMeta = Get-SourceMeta -Section $wingetSection
    $rebootPending = $false
    $rebootValue = Get-DictionaryValue -Object $windowsMeta -Name 'rebootPending'
    if ($rebootValue) { $rebootPending = $true }
    if (-not $doWindows) {
        $reboot = Get-RebootPending
        $rebootPending = [bool]$reboot.pending
    }

    $eligibleTotal = [int]$counts.applyEligible
    $confirmed = $false
    $cancelled = $false
    $confirmationError = $null
    if ((Test-DryRun)) {
        Write-Log -Level INFO -Event 'apply_skipped' -Data @{ reason = 'DryRun' }
    }
    elseif ($eligibleTotal -le 0) {
        Write-Log -Level INFO -Event 'apply_skipped' -Data @{ reason = 'No hay actualizaciones no criticas que instalar.' }
    }
    else {
        try {
            $windowsCounts = Get-DictionaryValue -Object $windowsMeta -Name 'counts'
            $wingetCounts = Get-DictionaryValue -Object $wingetMeta -Name 'counts'
            $confirmed = Confirm-UpdateApply `
                -WindowsRecommended ([int](Get-DictionaryValue -Object $windowsCounts -Name 'recommended')) `
                -WindowsOptional ([int](Get-DictionaryValue -Object $windowsCounts -Name 'optional')) `
                -WingetRecommended ([int](Get-DictionaryValue -Object $wingetCounts -Name 'recommended')) `
                -WingetOptional ([int](Get-DictionaryValue -Object $wingetCounts -Name 'optional')) `
                -CriticalCount ([int]$counts.critical)
        }
        catch {
            $confirmationError = $_.Exception.Message
            $cancelled = $true
            Write-Log -Level WARN -Event 'apply_cancelled' -Data @{ reason = $confirmationError }
        }
        if (-not $confirmed -and -not $confirmationError) {
            $cancelled = $true
            Write-Log -Level INFO -Event 'apply_cancelled' -Data @{ reason = 'El usuario no escribio SI.' }
        }
    }

    $appliedCount = 0
    $failedCount = 0
    if ($confirmed) {
        if ($rebootPending) {
            $reason = 'Ya hay un reinicio pendiente. No se instalo nada.'
            Set-SkippedApply -Items $updates -Source 'WindowsUpdate' -Reason $reason
            Set-SkippedApply -Items $updates -Source 'Winget' -Reason $reason
            Write-Log -Level WARN -Event 'install_skipped' -Data @{ source = 'all'; reason = $reason }
        }
        else {
            if ($doWindows) { Install-NonCriticalWindowsUpdates -Items $updates }
            if ($doWinget) { Install-NonCriticalWingetUpdates -Items $updates }
        }
        foreach ($item in $updates) {
            $resultName = [string](Get-DictionaryValue -Object $item -Name 'applyResult')
            if ($resultName -eq 'Installed' -or $resultName -eq 'InstalledWithErrors') { $appliedCount++ }
            elseif ($resultName -eq 'Error' -or $resultName -eq 'Failed' -or $resultName -eq 'DownloadFailed' -or $resultName -eq 'Aborted') { $failedCount++ }
        }
    }

    $sourceAttention = $false
    foreach ($meta in @($windowsMeta, $wingetMeta)) {
        $one = [string](Get-DictionaryValue -Object $meta -Name 'status')
        if ($one -and $one -ne 'Ok' -and $one -ne 'Skipped') { $sourceAttention = $true }
    }
    if ([int]$counts.critical -gt 0) {
        $overall = 'Critical'
    }
    elseif ($sourceAttention -or $rebootPending -or [int]$counts.recommended -gt 0 -or [int]$counts.optional -gt 0) {
        $overall = 'Warning'
    }
    else {
        $overall = 'Ok'
    }

    $findings = New-Object System.Collections.Generic.List[string]
    if ((Test-DryRun)) {
        $findings.Add('Modo dry-run: no se descargo ni se instalo nada.')
    }
    elseif ($confirmationError) {
        $findings.Add($confirmationError)
    }
    elseif ($cancelled) {
        $findings.Add('Apply cancelado: no se escribio SI. No se modifico el equipo.')
    }
    elseif ($appliedCount -eq 0 -and $failedCount -eq 0) {
        $findings.Add('Apply activo, pero no se instalo ninguna actualizacion.')
    }
    else {
        $findings.Add(("Instaladas: {0}. Con error: {1}." -f $appliedCount, $failedCount))
    }
    $findings.Add(("Criticas: {0}. Recomendadas: {1}. Opcionales: {2}. Elegibles para apply: {3}." -f [int]$counts.critical, [int]$counts.recommended, [int]$counts.optional, $eligibleTotal))
    $findings.Add('Las actualizaciones criticas se informan y este script no las instala.')
    if ($rebootPending) {
        $findings.Add('Hay un reinicio pendiente. -Apply no instalara nada hasta reiniciar.')
    }
    $windowsHeadline = [string](Get-DictionaryValue -Object $windowsMeta -Name 'headline')
    $wingetHeadline = [string](Get-DictionaryValue -Object $wingetMeta -Name 'headline')
    if ($doWindows) { $findings.Add(("Windows Update: {0}" -f $windowsHeadline)) }
    if ($doWinget) { $findings.Add(("winget: {0}" -f $wingetHeadline)) }

    $notes = New-Object System.Collections.Generic.List[string]
    $notes.Add('Critical real: obligatoria, MSRC Critical o Important, o categoria Critical Updates.')
    $notes.Add('No es Critical: definiciones, controladores, BrowseOnly, feature packs, Preview, Security Updates Moderate/Low o sin severidad, y cualquier paquete winget.')
    $notes.Add('Recommended: actualizaciones normales de Windows y paquetes winget con version conocida.')
    $notes.Add('Optional: controladores, BrowseOnly, feature packs, herramientas, Preview y winget sin version exacta.')
    $notes.Add('El codigo 2 sale solo si hay al menos una actualizacion Critical. Un origen que falla, sin ninguna Critical, deja el reporte en Warning.')
    $notes.Add('Sin -Apply no se descarga ni se instala. -Apply nunca incluye Critical y nunca reinicia el equipo.')
    $notes.Add('winget no aporta severidad MSRC. Los paquetes anclados no se listan.')
    $notes.Add('El catalogo de Windows Update se abre como mucho una vez por ventana de UpdateCacheHours. DailyOptimizer y SystemHealthReport solo leen el snapshot.')

    $stopwatch.Stop()
    $report = [ordered]@{
        schemaVersion = '1.0'
        generatedAt   = (Get-UtcTimestamp)
        hostname      = [string]$env:COMPUTERNAME
        userName      = [string]$env:USERNAME
        runId         = (Get-RunId)
        elevated      = $elevated
        dryRun        = [bool](Test-DryRun)
        parameters    = [ordered]@{
            windowsUpdate         = $doWindows
            winget                = $doWinget
            applyScope            = $ApplyScope
            windowsTimeoutSeconds = $WindowsTimeoutSeconds
            wingetTimeoutSeconds  = $WingetTimeoutSeconds
            maxUpdates            = $MaxUpdates
        }
        sources       = [ordered]@{
            windowsUpdate = $windowsMeta
            winget        = $wingetMeta
        }
        updates       = $updates
        notes         = $notes
        summary       = [ordered]@{
            overallStatus     = $overall
            counts            = $counts
            rebootPending     = $rebootPending
            cancelled         = $cancelled
            confirmationError = $confirmationError
            appliedCount      = $appliedCount
            failedCount       = $failedCount
            findings          = $findings
            durationMs        = [int64]$stopwatch.ElapsedMilliseconds
        }
    }

    $json = Write-JsonReport -Path $ReportPath -Report $report
    Write-Log -Level INFO -Event 'report_written' -Data @{ path = $ReportPath }

    $consoleRows = New-Object System.Collections.Generic.List[object]
    foreach ($item in $updates) {
        $source = [string](Get-DictionaryValue -Object $item -Name 'source')
        $detail = [string](Get-DictionaryValue -Object $item -Name 'severity')
        if ($source -eq 'Winget') {
            $installedVersion = [string](Get-DictionaryValue -Object $item -Name 'installedVersion')
            $availableVersion = [string](Get-DictionaryValue -Object $item -Name 'availableVersion')
            $detail = ("{0} -> {1}" -f $installedVersion, $availableVersion)
        }
        elseif ([string]::IsNullOrWhiteSpace($detail)) {
            $detail = Get-ShortText -Text ([string](Get-DictionaryValue -Object $item -Name 'categories')) -Max 32
        }
        $sizeText = 'n/d'
        $sizeValue = Get-DictionaryValue -Object $item -Name 'sizeBytes'
        if ($null -ne $sizeValue) { $sizeText = Format-ByteSize -Bytes ([int64]$sizeValue) }
        $eligibleText = 'no'
        if (Get-DictionaryValue -Object $item -Name 'applyEligible') { $eligibleText = 'si' }
        $consoleRows.Add([pscustomobject]@{
            Prioridad = (Get-DictionaryValue -Object $item -Name 'priority')
            Origen    = $source
            Id        = (Get-ShortText -Text ([string](Get-DictionaryValue -Object $item -Name 'identifier')) -Max 42)
            Detalle   = (Get-ShortText -Text $detail -Max 28)
            Tamano    = $sizeText
            Aplica    = $eligibleText
            Titulo    = (Get-ShortText -Text ([string](Get-DictionaryValue -Object $item -Name 'title')) -Max 64)
        })
    }
    Write-ConsoleTable -Title 'Catalogo priorizado' -Rows $consoleRows

    $summaryRows = New-Object System.Collections.Generic.List[object]
    $summaryRows.Add([pscustomobject]@{ Prioridad = 'Critical'; Cantidad = [int]$counts.critical })
    $summaryRows.Add([pscustomobject]@{ Prioridad = 'Recommended'; Cantidad = [int]$counts.recommended })
    $summaryRows.Add([pscustomobject]@{ Prioridad = 'Optional'; Cantidad = [int]$counts.optional })
    $summaryRows.Add([pscustomobject]@{ Prioridad = 'Elegibles'; Cantidad = $eligibleTotal })
    $summaryRows.Add([pscustomobject]@{ Prioridad = 'Aplicadas'; Cantidad = $appliedCount })
    Write-ConsoleTable -Title 'Resumen' -Rows $summaryRows

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
        total         = [int]$counts.total
        critical      = [int]$counts.critical
        recommended   = [int]$counts.recommended
        optional      = [int]$counts.optional
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
    Clear-WindowsUpdateJob
}

Write-Host ("Codigo de salida: {0}" -f $script:ExitCode)
if ($MyInvocation.InvocationName -ne '.') {
    exit $script:ExitCode
}
