#Requires -Version 5.1
<#
.SYNOPSIS
    Modulo comun de los scripts LocalAgents.

.DESCRIPTION
    Centraliza la sesion de cada ejecucion, el log JSONL con rotacion,
    la escritura de reportes JSON, el codigo de salida por severidad
    y los helpers de formato, dry-run y error.

    Cada script llama a Initialize-CommonSession despues de resolver
    sus rutas. Las funciones del modulo no ven las variables $script:
    del script que las llama: el estado vive aqui.

    Codigos de salida de un reporte cerrado:
      0  Ok
      1  Warning (Unknown tambien sale 1)
      2  Critical o Error
    Un fallo antes de cerrar el reporte no pasa por Get-ReportExitCode.
    Get-FatalExitCode devuelve 3 para ese caso.

    El JSONL rota cuando el archivo activo llega a MaxLogBytes
    (predeterminado 5 MB). Se conservan LogGenerations copias
    (predeterminado 3): nombre.1.jsonl es la mas reciente y la mas
    vieja se elimina al rotar. El append usa UTF-8 sin BOM.
    Los reportes JSON se escriben en UTF-8 con BOM.

.NOTES
    Autor: LocalAgents
    Requiere: Windows PowerShell 5.1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:SessionReady = $false
$script:RunId = $null
$script:ScriptName = $null
$script:LogPath = $null
$script:LogReady = $false
$script:DryRun = $true
$script:JsonSerializer = $null
$script:MaxLogBytes = 5MB
$script:LogGenerations = 3
$script:MaxReportBytes = 32MB
$script:LogProperties = @{}

function Get-UtcTimestamp {
    return ([DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss.fff', [System.Globalization.CultureInfo]::InvariantCulture) + 'Z')
}

function Format-Utc {
    param($Value)
    if ($null -eq $Value) { return $null }
    try {
        $date = [DateTime]$Value
        if ($date.Year -lt 2000) { return $null }
        return $date.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fff', [System.Globalization.CultureInfo]::InvariantCulture) + 'Z'
    }
    catch {
        return $null
    }
}

function Initialize-JsonSerializer {
    if ($null -ne $script:JsonSerializer) { return }
    Add-Type -AssemblyName System.Web.Extensions
    $serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    $serializer.MaxJsonLength = [int]::MaxValue
    $serializer.RecursionLimit = 80
    $script:JsonSerializer = $serializer
}

function Format-CompactJson {
    param([Parameter(Mandatory = $true)][string]$Json)

    $builder = New-Object System.Text.StringBuilder
    $indent = 0
    $inString = $false
    $escape = $false
    $chars = $Json.ToCharArray()

    for ($index = 0; $index -lt $chars.Length; $index++) {
        $char = $chars[$index]
        if ($inString) {
            [void]$builder.Append($char)
            if ($escape) { $escape = $false }
            elseif ($char -eq '\') { $escape = $true }
            elseif ($char -eq '"') { $inString = $false }
            continue
        }
        if ($char -eq '"') {
            $inString = $true
            [void]$builder.Append($char)
            continue
        }
        if ($char -eq '{' -or $char -eq '[') {
            $closing = '}'
            if ($char -eq '[') { $closing = ']' }
            if (($index + 1) -lt $chars.Length -and $chars[$index + 1] -eq [char]$closing) {
                [void]$builder.Append($char)
                [void]$builder.Append($closing)
                $index++
                continue
            }
            $indent++
            [void]$builder.Append($char)
            [void]$builder.AppendLine()
            [void]$builder.Append(' ' * ($indent * 2))
            continue
        }
        if ($char -eq '}' -or $char -eq ']') {
            $indent--
            if ($indent -lt 0) { $indent = 0 }
            [void]$builder.AppendLine()
            [void]$builder.Append(' ' * ($indent * 2))
            [void]$builder.Append($char)
            continue
        }
        if ($char -eq ',') {
            [void]$builder.Append($char)
            [void]$builder.AppendLine()
            [void]$builder.Append(' ' * ($indent * 2))
            continue
        }
        if ($char -eq ':') {
            [void]$builder.Append(': ')
            continue
        }
        [void]$builder.Append($char)
    }

    return ($builder.ToString().TrimEnd() + [Environment]::NewLine)
}

function ConvertTo-PlainData {
    param($Value)

    if ($null -eq $Value) { return $null }

    $base = $Value
    if ($Value -is [psobject]) {
        $unwrapped = $Value.PSObject.BaseObject
        if ($null -ne $unwrapped) { $base = $unwrapped }
    }

    if ($base -is [string]) { return [string]$base }
    if ($base -is [bool] -or $base -is [byte] -or $base -is [int16] -or $base -is [int32] -or $base -is [int64] -or $base -is [single] -or $base -is [double] -or $base -is [decimal]) {
        return $base
    }

    $dictionary = $base -as [System.Collections.IDictionary]
    if ($null -ne $dictionary) {
        $plain = [ordered]@{}
        foreach ($key in $dictionary.Keys) {
            $plain[[string]$key] = ConvertTo-PlainData $dictionary[$key]
        }
        return ,$plain
    }

    $enumerable = $base -as [System.Collections.IEnumerable]
    if ($null -ne $enumerable) {
        $items = New-Object System.Collections.Generic.List[object]
        foreach ($item in $enumerable) {
            $items.Add((ConvertTo-PlainData $item))
        }
        return ,$items.ToArray()
    }

    return [string]$base
}

function ConvertTo-DataJson {
    param([Parameter(Mandatory = $true)]$Value)
    Initialize-JsonSerializer
    $plain = ConvertTo-PlainData $Value
    return $script:JsonSerializer.Serialize($plain)
}

function ConvertFrom-DataJson {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Json)
    Initialize-JsonSerializer
    return $script:JsonSerializer.DeserializeObject($Json)
}

function ConvertTo-ReportJson {
    param([Parameter(Mandatory = $true)]$Report)
    $compact = ConvertTo-DataJson -Value $Report
    return (Format-CompactJson -Json $compact)
}

function ConvertTo-LogLevel {
    param([string]$Level)
    if ([string]::IsNullOrWhiteSpace($Level)) { return 'Info' }
    switch ($Level.Trim().ToUpperInvariant()) {
        'DEBUG' { return 'Debug' }
        'INFO' { return 'Info' }
        'WARN' { return 'Warning' }
        'WARNING' { return 'Warning' }
        'ERROR' { return 'Error' }
        'CRITICAL' { return 'Critical' }
        default { return 'Info' }
    }
}

function Invoke-LogRotation {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][long]$MaxBytes,
        [Parameter(Mandatory = $true)][int]$Generations
    )

    if ($MaxBytes -le 0 -or $Generations -lt 1) { return }
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $length = (Get-Item -LiteralPath $Path).Length
    if ($length -lt $MaxBytes) { return }

    $directory = Split-Path -Parent $Path
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $extension = [System.IO.Path]::GetExtension($Path)

    for ($index = $Generations; $index -ge 1; $index--) {
        $current = Join-Path $directory ('{0}.{1}{2}' -f $baseName, $index, $extension)
        if (-not (Test-Path -LiteralPath $current)) { continue }
        if ($index -eq $Generations) {
            [System.IO.File]::Delete($current)
            continue
        }
        $older = Join-Path $directory ('{0}.{1}{2}' -f $baseName, ($index + 1), $extension)
        if (Test-Path -LiteralPath $older) {
            [System.IO.File]::Delete($older)
        }
        [System.IO.File]::Move($current, $older)
    }

    $newest = Join-Path $directory ('{0}.1{1}' -f $baseName, $extension)
    if (Test-Path -LiteralPath $newest) {
        [System.IO.File]::Delete($newest)
    }
    [System.IO.File]::Move($Path, $newest)
}

function Write-Log {
    <#
    .SYNOPSIS
        Escribe una linea JSONL con nivel Info, Warning, Error o Critical.
    .DESCRIPTION
        Acepta tambien los niveles antiguos INFO, WARN, ERROR y DEBUG.
        WARN se guarda como Warning. Si la sesion no esta lista, no escribe.
        dryRun sale en cada linea. Set-LogProperty agrega campos fijos.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Level,
        [Parameter(Mandatory = $true)]
        [string]$Event,
        [hashtable]$Data
    )

    if (-not $script:LogReady) { return }

    $entry = [ordered]@{
        timestamp = (Get-UtcTimestamp)
        level     = (ConvertTo-LogLevel -Level $Level)
        event     = $Event
        script    = $script:ScriptName
        runId     = $script:RunId
        pid       = $PID
        dryRun    = [bool]$script:DryRun
    }
    if ($script:LogProperties) {
        foreach ($key in @($script:LogProperties.Keys)) {
            $entry[[string]$key] = $script:LogProperties[$key]
        }
    }
    if ($Data) {
        foreach ($key in @($Data.Keys)) {
            $entry[[string]$key] = $Data[$key]
        }
    }

    try {
        Invoke-LogRotation -Path $script:LogPath -MaxBytes $script:MaxLogBytes -Generations $script:LogGenerations
        $line = (ConvertTo-DataJson -Value $entry) + [Environment]::NewLine
        $encoding = New-Object System.Text.UTF8Encoding $false
        $bytes = $encoding.GetBytes($line)
        $stream = [System.IO.File]::Open(
            $script:LogPath,
            [System.IO.FileMode]::Append,
            [System.IO.FileAccess]::Write,
            [System.IO.FileShare]::Read
        )
        try {
            $stream.Write($bytes, 0, $bytes.Length)
            $stream.Flush()
        }
        finally {
            $stream.Dispose()
        }
    }
    catch {
        Write-Warning ("No se pudo escribir el log: {0}" -f $_.Exception.Message)
    }
}

function Write-StructuredLog {
    param(
        [Parameter(Mandatory = $true)][string]$Level,
        [Parameter(Mandatory = $true)][string]$Event,
        [hashtable]$Data
    )
    Write-Log -Level $Level -Event $Event -Data $Data
}

function Write-JsonReport {
    <#
    .SYNOPSIS
        Escribe un reporte JSON con UTF-8 BOM y avisa si pasa el tope.
    .DESCRIPTION
        Pase -Report para serializarlo, o -Json si el texto ya existe.
        Devuelve el texto escrito. No recorta el reporte: el tope
        (32 MB) solo deja un aviso en el log.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        $Report,
        [string]$Json
    )

    if ([string]::IsNullOrEmpty($Json)) {
        if ($null -eq $Report) { throw 'Write-JsonReport necesita -Report o -Json.' }
        $Json = ConvertTo-ReportJson -Report $Report
    }

    $directory = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($directory)) {
        Ensure-Directory -Path $directory
    }

    $sizeBytes = [int64][System.Text.Encoding]::UTF8.GetByteCount($Json)
    if ($sizeBytes -gt $script:MaxReportBytes) {
        Write-Log -Level Warning -Event 'report_size' -Data @{
            path     = $Path
            bytes    = $sizeBytes
            maxBytes = [int64]$script:MaxReportBytes
        }
    }

    $encoding = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($Path, $Json, $encoding)
    return $Json
}

function Get-Utf8BomEncoding {
    return (New-Object System.Text.UTF8Encoding $true)
}

function Format-ByteSize {
    param([int64]$Bytes)
    if ($Bytes -lt 0) { $Bytes = 0 }
    $units = @(
        @{ Name = 'TB'; Size = 1TB },
        @{ Name = 'GB'; Size = 1GB },
        @{ Name = 'MB'; Size = 1MB },
        @{ Name = 'KB'; Size = 1KB }
    )
    foreach ($unit in $units) {
        if ($Bytes -ge [int64]$unit.Size) {
            return ('{0:N2} {1}' -f ($Bytes / [double]$unit.Size), $unit.Name)
        }
    }
    return ("{0} B" -f $Bytes)
}

function Get-RepositoryRoot {
    $scriptDir = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($scriptDir)) {
        throw 'No se pudo resolver la carpeta del modulo Common.'
    }
    return (Split-Path -Parent $scriptDir)
}

function Ensure-Directory {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Get-ItemCount {
    param($Value)
    if ($null -eq $Value) { return 0 }
    # No usar .Count: con un solo elemento, PowerShell 5.1 puede devolver
    # el Count de ese elemento y no el de la coleccion.
    $total = 0
    foreach ($item in $Value) {
        if ($null -ne $item) { $total++ }
    }
    return $total
}

function Test-DictionaryKey {
    param($Dictionary, [string]$Name)
    if ($null -eq $Dictionary) { return $false }
    foreach ($key in $Dictionary.Keys) {
        if ([string]$key -eq $Name) { return $true }
    }
    return $false
}

function Get-DictionaryValue {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    $dictionary = $Object -as [System.Collections.IDictionary]
    if ($null -eq $dictionary) { return $null }
    if (-not (Test-DictionaryKey -Dictionary $dictionary -Name $Name)) { return $null }
    $value = $dictionary[$Name]
    if (($value -is [System.Collections.IList]) -and -not ($value -is [string])) {
        return ,$value
    }
    return $value
}

function Get-ReportField {
    param($Object, [string]$Name)
    $value = Get-DictionaryValue -Object $Object -Name $Name
    if (($value -is [System.Collections.IList]) -and -not ($value -is [string])) {
        return ,$value
    }
    return $value
}

function Get-JsonField {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    $dictionary = $Object -as [System.Collections.IDictionary]
    if ($null -eq $dictionary) { return $null }
    if (-not (Test-DictionaryKey -Dictionary $dictionary -Name $Name)) { return $null }
    return $dictionary[$Name]
}

function Get-JsonText {
    param($Object, [string]$Name)
    $value = Get-JsonField -Object $Object -Name $Name
    if ($null -eq $value) { return '' }
    return [string]$value
}

function Get-JsonBool {
    param($Object, [string]$Name)
    $value = Get-JsonField -Object $Object -Name $Name
    if ($null -eq $value) { return $false }
    try { return [bool]$value } catch { return $false }
}

function Get-WorstStatus {
    param([string[]]$Statuses)
    if ($Statuses -contains 'Critical') { return 'Critical' }
    if ($Statuses -contains 'Error') { return 'Error' }
    if ($Statuses -contains 'Warning') { return 'Warning' }
    if ($Statuses -contains 'Ok') { return 'Ok' }
    return 'Unknown'
}

function Get-ReportExitCode {
    <#
    .SYNOPSIS
        Traduce el estado del reporte a un codigo de salida.
    .DESCRIPTION
        Ok = 0, Warning = 1, Critical = 2, Error = 2.
        Unknown y cualquier otro texto salen 1.
        El fallo previo al reporte es 3 y lo da Get-FatalExitCode.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Status
    )
    switch ($Status) {
        'Ok' { return 0 }
        'Warning' { return 1 }
        'Critical' { return 2 }
        'Error' { return 2 }
        default { return 1 }
    }
}

function Get-FatalExitCode {
    return 3
}

function Get-ShortText {
    param(
        [AllowEmptyString()][string]$Text,
        [int]$Max = 72
    )
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $clean = ([regex]::Replace($Text, '\s+', ' ')).Trim()
    if ($clean.Length -le $Max) { return $clean }
    if ($Max -le 3) { return $clean.Substring(0, $Max) }
    return ($clean.Substring(0, $Max - 3) + '...')
}

function Write-ConsoleTable {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        $Rows
    )
    $buffer = New-Object System.Collections.ArrayList
    if ($null -ne $Rows) {
        foreach ($row in $Rows) { [void]$buffer.Add($row) }
    }
    Write-Host ''
    Write-Host $Title
    $ruleWidth = $Title.Length
    if ($ruleWidth -gt 80) { $ruleWidth = 80 }
    if ($ruleWidth -lt 1) { $ruleWidth = 1 }
    Write-Host ('-' * $ruleWidth)
    if ($buffer.Count -eq 0) {
        Write-Host '(sin datos)'
        return
    }
    $table = $buffer.ToArray() | Format-Table -AutoSize | Out-String -Width 220
    Write-Host $table.TrimEnd()
}

function Test-IsElevated {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal $identity
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Test-InteractiveSession {
    try {
        if (-not [Environment]::UserInteractive) { return $false }
        return -not [Console]::IsInputRedirected
    }
    catch {
        return $false
    }
}

function Get-RebootPending {
    $reasons = New-Object System.Collections.Generic.List[string]
    $keys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
    )
    foreach ($key in $keys) {
        if (Test-Path -LiteralPath $key) { $reasons.Add($key) }
    }
    try {
        $property = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction Stop
        if ($null -ne $property.PendingFileRenameOperations) {
            $reasons.Add('HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations')
        }
    }
    catch { }
    return [ordered]@{
        pending = ($reasons.Count -gt 0)
        reasons = $reasons
    }
}

function Get-RebootPendingReport {
    return (Get-RebootPending)
}

function Get-NotificationLevelText {
    param($Level)
    switch ([string]$Level) {
        '1' { return 'Nunca comprobar' }
        '2' { return 'Comprobar, sin descargar' }
        '3' { return 'Descargar, sin instalar' }
        '4' { return 'Instalacion automatica' }
        default { return [string]$Level }
    }
}

function Initialize-CommonSession {
    <#
    .SYNOPSIS
        Prepara runId, log JSONL, dry-run y el serializador JSON.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$ScriptName,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [string]$RunId,
        [bool]$DryRun = $true,
        [long]$MaxLogBytes = 5MB,
        [int]$LogGenerations = 3
    )

    if ([string]::IsNullOrWhiteSpace($RunId)) {
        $RunId = [guid]::NewGuid().ToString('n')
    }
    if ($LogGenerations -lt 1) { $LogGenerations = 1 }
    if ($MaxLogBytes -lt 0) { $MaxLogBytes = 0 }

    $script:ScriptName = $ScriptName
    $script:RunId = $RunId
    $script:DryRun = [bool]$DryRun
    $script:MaxLogBytes = $MaxLogBytes
    $script:LogGenerations = $LogGenerations
    $script:LogProperties = @{}
    $script:JsonSerializer = $null
    Initialize-JsonSerializer

    $fullLog = [System.IO.Path]::GetFullPath($LogPath)
    $parent = Split-Path -Parent $fullLog
    Ensure-Directory -Path $parent
    $script:LogPath = $fullLog
    $script:LogReady = $true
    $script:SessionReady = $true
}

function Set-DryRunMode {
    param([Parameter(Mandatory = $true)][bool]$DryRun)
    $script:DryRun = [bool]$DryRun
}

function Test-DryRun {
    return [bool]$script:DryRun
}

function Get-RunId {
    return $script:RunId
}

function Get-SessionLogPath {
    return $script:LogPath
}

function Test-LogReady {
    return [bool]$script:LogReady
}

function Set-LogProperty {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        $Value
    )
    if ($null -eq $script:LogProperties) {
        $script:LogProperties = @{}
    }
    $script:LogProperties[$Name] = $Value
}

function Write-ScriptError {
    <#
    .SYNOPSIS
        Registra un fallo y lo escribe en la salida de error.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $message = [string]$ErrorRecord.Exception.Message
    $position = ''
    $stack = ''
    if ($ErrorRecord.InvocationInfo) {
        $position = [string]$ErrorRecord.InvocationInfo.PositionMessage
    }
    if ($ErrorRecord.ScriptStackTrace) {
        $stack = [string]$ErrorRecord.ScriptStackTrace
    }

    if (Test-LogReady) {
        Write-Log -Level Error -Event 'run_failed' -Data @{
            message  = $message
            position = $position
            stack    = $stack
        }
    }

    [Console]::Error.WriteLine(("ERROR: {0}" -f $message))
    if (-not [string]::IsNullOrWhiteSpace($position)) {
        [Console]::Error.WriteLine($position)
    }
    if (-not [string]::IsNullOrWhiteSpace($stack)) {
        [Console]::Error.WriteLine($stack)
    }
}

function Initialize-TreeWalker {
    if ('LocalAgents.TreeWalker' -as [type]) { return }
    $source = @'
using System;
using System.Collections.Generic;
using System.IO;

namespace LocalAgents
{
    public sealed class TreeFileHit
    {
        public string Path { get; set; }
        public long Bytes { get; set; }
        public string LastWriteUtc { get; set; }
    }

    public sealed class TreeMeasureResult
    {
        public long Enumerated { get; set; }
        public long FileCount { get; set; }
        public long ByteCount { get; set; }
        public int AccessErrors { get; set; }
        public int ProtectedSkipped { get; set; }
        public string FirstError { get; set; }
        public TreeFileHit[] Largest { get; set; }
    }

    public static class TreeWalker
    {
        public static TreeMeasureResult Measure(string root, string pattern, long cutoffUtcTicks, int top, string[] protectedRoots, bool rootOnly)
        {
            TreeMeasureResult result = new TreeMeasureResult();
            List<TreeFileHit> largest = new List<TreeFileHit>();
            if (top < 1) { top = 1; }
            if (string.IsNullOrEmpty(root))
            {
                result.AccessErrors = 1;
                result.FirstError = "Ruta vacia.";
                result.Largest = new TreeFileHit[0];
                return result;
            }

            Stack<string> stack = new Stack<string>();
            stack.Push(root);
            while (stack.Count > 0)
            {
                string current = stack.Pop();
                bool atRoot = current.Equals(root, StringComparison.OrdinalIgnoreCase);
                DirectoryInfo directory;
                try
                {
                    directory = new DirectoryInfo(current);
                }
                catch (Exception ex)
                {
                    Note(result, ex);
                    continue;
                }

                IEnumerable<FileInfo> files;
                try
                {
                    if (rootOnly)
                    {
                        if (!atRoot) { continue; }
                        files = string.IsNullOrEmpty(pattern) ? directory.EnumerateFiles() : directory.EnumerateFiles(pattern);
                    }
                    else
                    {
                        files = directory.EnumerateFiles();
                    }
                }
                catch (FileNotFoundException) { continue; }
                catch (DirectoryNotFoundException) { continue; }
                catch (Exception ex)
                {
                    Note(result, ex);
                    continue;
                }

                foreach (FileInfo file in files)
                {
                    try
                    {
                        result.Enumerated++;
                        if ((file.Attributes & FileAttributes.ReparsePoint) != 0) { continue; }
                        if (cutoffUtcTicks > 0 && file.LastWriteTimeUtc.Ticks >= cutoffUtcTicks) { continue; }
                        if (IsProtected(file.FullName, protectedRoots))
                        {
                            result.ProtectedSkipped++;
                            continue;
                        }
                        long length = file.Length;
                        result.FileCount++;
                        result.ByteCount += length;
                        Consider(largest, top, file.FullName, length, file.LastWriteTimeUtc);
                    }
                    catch (FileNotFoundException) { }
                    catch (DirectoryNotFoundException) { }
                    catch (Exception ex)
                    {
                        Note(result, ex);
                    }
                }

                if (rootOnly) { continue; }

                IEnumerable<DirectoryInfo> children;
                try
                {
                    children = directory.EnumerateDirectories();
                }
                catch (FileNotFoundException) { continue; }
                catch (DirectoryNotFoundException) { continue; }
                catch (Exception ex)
                {
                    Note(result, ex);
                    continue;
                }

                foreach (DirectoryInfo child in children)
                {
                    try
                    {
                        if ((child.Attributes & FileAttributes.ReparsePoint) != 0) { continue; }
                        if (IsProtected(child.FullName, protectedRoots))
                        {
                            result.ProtectedSkipped++;
                            continue;
                        }
                        stack.Push(child.FullName);
                    }
                    catch (Exception ex)
                    {
                        Note(result, ex);
                    }
                }
            }

            largest.Sort(delegate(TreeFileHit left, TreeFileHit right)
            {
                return right.Bytes.CompareTo(left.Bytes);
            });
            result.Largest = largest.ToArray();
            return result;
        }

        static void Consider(List<TreeFileHit> largest, int top, string path, long bytes, DateTime lastWriteUtc)
        {
            TreeFileHit hit = new TreeFileHit();
            hit.Path = path;
            hit.Bytes = bytes;
            hit.LastWriteUtc = lastWriteUtc.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fff") + "Z";
            if (largest.Count < top)
            {
                largest.Add(hit);
                return;
            }
            int minIndex = 0;
            long minBytes = largest[0].Bytes;
            for (int index = 1; index < largest.Count; index++)
            {
                if (largest[index].Bytes < minBytes)
                {
                    minBytes = largest[index].Bytes;
                    minIndex = index;
                }
            }
            if (bytes > minBytes)
            {
                largest[minIndex] = hit;
            }
        }

        static bool IsProtected(string path, string[] roots)
        {
            if (roots == null || string.IsNullOrEmpty(path)) { return false; }
            for (int index = 0; index < roots.Length; index++)
            {
                string root = roots[index];
                if (string.IsNullOrEmpty(root)) { continue; }
                if (path.Equals(root, StringComparison.OrdinalIgnoreCase)) { return true; }
                string prefix = root.EndsWith("\\", StringComparison.Ordinal) ? root : root + "\\";
                if (path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) { return true; }
            }
            return false;
        }

        static void Note(TreeMeasureResult result, Exception ex)
        {
            result.AccessErrors++;
            if (string.IsNullOrEmpty(result.FirstError) && ex != null)
            {
                result.FirstError = ex.Message;
            }
        }
    }
}
'@
    Add-Type -TypeDefinition $source -Language CSharp
}

function Measure-TreeFiles {
    <#
    .SYNOPSIS
        Mide una carpeta en una sola pasada de solo lectura.
    .DESCRIPTION
        No borra, no sigue puntos de union y no entra en las raices protegidas.
        CutoffUtcTicks menor o igual a 0 no filtra por fecha. Un archivo entra
        cuando su LastWriteTimeUtc es anterior a ese corte.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [string]$Pattern,
        [int64]$CutoffUtcTicks = -1,
        [int]$Top = 15,
        [string[]]$ProtectedRoots,
        [switch]$RootOnly
    )

    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    Initialize-TreeWalker
    $roots = @()
    if ($null -ne $ProtectedRoots) { $roots = @($ProtectedRoots) }
    $native = [LocalAgents.TreeWalker]::Measure($Root, $Pattern, $CutoffUtcTicks, $Top, $roots, [bool]$RootOnly)
    $largest = New-Object System.Collections.Generic.List[object]
    if ($null -ne $native -and $null -ne $native.Largest) {
        foreach ($hit in $native.Largest) {
            if ($null -eq $hit) { continue }
            $largest.Add([ordered]@{
                path             = [string]$hit.Path
                sizeBytes        = [int64]$hit.Bytes
                lastWriteTimeUtc = [string]$hit.LastWriteUtc
            })
        }
    }
    $clock.Stop()
    $enumerated = [int64]0
    $fileCount = [int64]0
    $sizeBytes = [int64]0
    $accessErrors = 0
    $protectedSkipped = 0
    $errorText = $null
    if ($null -ne $native) {
        $enumerated = [int64]$native.Enumerated
        $fileCount = [int64]$native.FileCount
        $sizeBytes = [int64]$native.ByteCount
        $accessErrors = [int]$native.AccessErrors
        $protectedSkipped = [int]$native.ProtectedSkipped
        if (-not [string]::IsNullOrWhiteSpace([string]$native.FirstError)) {
            $errorText = [string]$native.FirstError
        }
    }
    return [ordered]@{
        enumerated       = $enumerated
        fileCount        = $fileCount
        sizeBytes        = $sizeBytes
        accessErrors     = $accessErrors
        protectedSkipped = $protectedSkipped
        error            = $errorText
        largest          = $largest
        durationMs       = [int64]$clock.ElapsedMilliseconds
    }
}

function Get-WindowsUpdateSnapshotPath {
    return (Join-Path (Join-Path (Get-RepositoryRoot) 'reports') 'WindowsUpdateSnapshot.json')
}

function Read-WindowsUpdateSnapshot {
    <#
    .SYNOPSIS
        Lee el catalogo reciente escrito por UpdateChecker.
    .DESCRIPTION
        Devuelve el diccionario si el snapshot es de catalogo completo,
        fue consultado y no supera MaxAgeHours. Si no, devuelve null.
        No abre Windows Update.
    #>
    param([ValidateRange(1, 168)][int]$MaxAgeHours = 20)

    $path = Get-WindowsUpdateSnapshotPath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
    try {
        $raw = [System.IO.File]::ReadAllText($path)
        $parsed = ConvertFrom-DataJson -Json $raw
        if ($null -eq $parsed) { return $null }
        if ((Get-JsonText -Object $parsed -Name 'catalog') -ne 'full') { return $null }
        if (-not (Get-JsonBool -Object $parsed -Name 'queried')) { return $null }
        $generated = Get-JsonText -Object $parsed -Name 'generatedAt'
        if ([string]::IsNullOrWhiteSpace($generated)) { return $null }
        $when = [DateTime]::Parse($generated, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
        $ageHours = ([DateTime]::UtcNow - $when.ToUniversalTime()).TotalHours
        if ($ageHours -lt -1) { return $null }
        if ($ageHours -gt $MaxAgeHours) { return $null }
        return $parsed
    }
    catch {
        return $null
    }
}

function Write-WindowsUpdateSnapshot {
    <#
    .SYNOPSIS
        Guarda el unico catalogo de Windows Update de la cadena diaria.
    #>
    param([Parameter(Mandatory = $true)]$Snapshot)

    $path = Get-WindowsUpdateSnapshotPath
    [void](Write-JsonReport -Path $path -Report $Snapshot)
    return $path
}

Export-ModuleMember -Function @(
    'Get-UtcTimestamp',
    'Format-Utc',
    'ConvertTo-DataJson',
    'ConvertFrom-DataJson',
    'ConvertTo-ReportJson',
    'Write-Log',
    'Write-StructuredLog',
    'Write-JsonReport',
    'Get-Utf8BomEncoding',
    'Format-ByteSize',
    'Get-RepositoryRoot',
    'Ensure-Directory',
    'Get-ItemCount',
    'Test-DictionaryKey',
    'Get-DictionaryValue',
    'Get-ReportField',
    'Get-JsonField',
    'Get-JsonText',
    'Get-JsonBool',
    'Get-WorstStatus',
    'Get-ReportExitCode',
    'Get-FatalExitCode',
    'Get-ShortText',
    'Write-ConsoleTable',
    'Test-IsElevated',
    'Test-InteractiveSession',
    'Get-RebootPending',
    'Get-RebootPendingReport',
    'Get-NotificationLevelText',
    'Initialize-CommonSession',
    'Set-DryRunMode',
    'Test-DryRun',
    'Get-RunId',
    'Get-SessionLogPath',
    'Test-LogReady',
    'Set-LogProperty',
    'Write-ScriptError',
    'Measure-TreeFiles',
    'Get-WindowsUpdateSnapshotPath',
    'Read-WindowsUpdateSnapshot',
    'Write-WindowsUpdateSnapshot'
)
