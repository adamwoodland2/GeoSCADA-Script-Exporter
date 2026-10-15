# GeoSCADA-Script-Exporter.ps1
# Copyright (c) 2026  Adam Woodland
#
# Licensed under the MIT License. You may obtain a copy of the License in the LICENSE file
# distributed with this software, or at <https://opensource.org/licenses/MIT>.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
# BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
# NONINFRINGEMENT.

<#
.SYNOPSIS
    Exports every client-side script in a Geo SCADA database - mimic scripts and Script Library
    code, in VBScript, JScript, JScript9 and JavaScript V8 - to one .txt file per object, named
    after the object's full name, using the Geo SCADA .NET Client API (read-only).

.DESCRIPTION
    The script connects with ClearScada.Client.dll, logs on, and lists the CMimic and
    CScriptLibrary objects with SQL (SELECT WITH TEMPLATES, so objects inside template
    definitions are included). For each one:

      Mimics (CMimic)
        IObjectConfiguration.GetWebDocumentContent() returns the mimic definition as XML (the
        same XML the web server serves for /db/<Mimic>?content). From it the script takes:
          - <Script lang="{GUID}">  the script language
          - the <Script> text       the mimic's own script code
          - <Library name="..."/>   the script libraries the mimic links to
          - <Action type="scriptCall">  pick actions that call a script function
        A file is written when the mimic has its own code, links to a library, or has a pick
        action that calls a script.

      Script libraries (CScriptLibrary)
        The library's own server method GetScriptSource() returns the code. The language is not
        exposed by any API call, so it is read from the 16-byte language GUID at a fixed position
        in the library's stored record (IObjectConfiguration.GetDocumentContent()).

    Template-controlled objects: an object inside an instance takes its content from its template
    unless the instance overrides it, so it cannot be edited where it is. The header's
    EditableSource gives the object to edit instead, found by following
    IObjectHierarchy.GetTemplateObject through each template level (TemplateSource lines) until
    an object that is not instance-controlled, or that overrides the content, is reached.

    Languages are identified by GUID:
        VBScript      {B54F3741-5B07-11CF-A4B0-00AA004A55E8}  (also the default when none is set)
        JScript       {F414C260-6AC0-11CF-B6D1-00AA00BBBB58}
        JScript9      {16D51579-A30B-4C8B-A276-0FF4DC41E755}
        JavaScriptV8  {16BA25BB-FB03-48BA-AECA-C53BF6ABD2A1}  (new in Geo SCADA 2025 R2)
    Any other GUID is exported with Language 'Unknown' and the GUID recorded.

    Each file starts with a header of 'Key: Value' lines inside comment markers for the script's
    language, followed by the script exactly as stored (line endings normalised to CRLF). The
    header's ScriptStartLine gives the file line on which script line 1 sits. A manifest CSV
    (_ScriptExport_Manifest.csv) lists every file written. See README.md for the file format.

    -CheckMissingLibraries: the XML omits library links whose library no longer exists (or
    cannot be resolved). With this switch the mimic's stored document
    (GetDocumentContent) is also decoded to read the full list of library links, so dangling
    links are reported as 'LibraryMissing'. The decoded script text is compared with the XML
    as a cross-check. This reads an undocumented binary layout, so it is off by default.

    Nothing in the database is changed: the script only reads.

.PARAMETER OutputFolder
    Folder to write the .txt files and the manifest to. Created if missing. Defaults to
    .\ScriptExport next to this script.

.PARAMETER Server
    Geo SCADA server host name or address. Defaults to 127.0.0.1 (localhost).

.PARAMETER Port
    Geo SCADA client port. Defaults to 5481.

.PARAMETER Credential
    PSCredential for the Geo SCADA logon. If omitted, -Username/-Password are used; if those are
    also omitted you are prompted.

.PARAMETER Username
.PARAMETER Password
    Plain-text alternative to -Credential (handy for scripted runs; prefer -Credential).

.PARAMETER Group
    Only export objects at or below these database groups, e.g. 'Sewage.Site'. Can be combined
    with -Name.

.PARAMETER Name
    Only export these objects (full names).

.PARAMETER SkipMimics
    Do not export mimic scripts.

.PARAMETER SkipLibraries
    Do not export script libraries.

.PARAMETER CheckMissingLibraries
    Also decode each mimic's stored document to find library links the XML leaves out (links
    to libraries that do not exist). Off by default; see DESCRIPTION.

.PARAMETER ClientDllPath
    Path to ClearScada.Client.dll. By default the Geo SCADA client folder is tried first, then
    the server folder.

.PARAMETER FailOnError
    Stop at the first object that cannot be exported instead of continuing. Without it the
    script finishes, warns that the export is incomplete and exits with code 2.

.PARAMETER AcceptDisclaimer
    Accept the usage disclaimer non-interactively (required in non-interactive sessions).

.EXAMPLE
    .\GeoSCADA-Script-Exporter.ps1 -OutputFolder C:\Temp\Scripts -Username operator -Password 'secret'

.EXAMPLE
    .\GeoSCADA-Script-Exporter.ps1 -Server scada01 -Credential (Get-Credential) -Group 'Sewage.Site' -OutputFolder C:\Temp\Scripts -AcceptDisclaimer

.EXAMPLE
    .\GeoSCADA-Script-Exporter.ps1 -OutputFolder C:\Temp\Scripts -SkipMimics -Credential (Get-Credential)

.NOTES
    DISCLAIMER
    This script is provided "AS IS", without warranty of any kind, express or implied. In no
    event shall the author be liable for any claim, damages or other liability arising from,
    out of or in connection with the script or its use. It connects to the Geo SCADA server
    through the .NET Client API to read mimic and script library definitions and writes text
    files. It is NOT certified for production SCADA environments. Do NOT run it against a
    production or safety-critical system without first reviewing the code and testing it on a
    representative non-production system. You run it at your own risk and are responsible for
    compliance with your own change-control and security policies.

    Requires Windows PowerShell 5.1 (ClearScada.Client.dll targets .NET Framework 4.8) and a
    Geo SCADA user with permission to view the objects. Tested against Geo SCADA 2025 R2
    (ClearScada.Client.dll 6.88.9742.1).
#>
[CmdletBinding()]
param(
    [string]$OutputFolder,
    [string]$Server = '127.0.0.1',
    [int]$Port = 5481,
    [System.Management.Automation.PSCredential]$Credential,
    [string]$Username,
    [string]$Password,
    [string[]]$Group,
    [string[]]$Name,
    [switch]$SkipMimics,
    [switch]$SkipLibraries,
    [switch]$CheckMissingLibraries,
    [string]$ClientDllPath,
    [switch]$FailOnError,
    [switch]$AcceptDisclaimer
)

$ErrorActionPreference = 'Stop'
$ExporterVersion = '1.1.0'
$FormatVersion   = 2

# Resolve the output path here (not in the param default) so it works even when
# $PSScriptRoot is empty - e.g. running a selection (F8) or dot-sourcing.
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot }
             elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath }
             else { (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($OutputFolder)) { $OutputFolder = Join-Path $scriptDir 'ScriptExport' }
$OutputFolder = [System.IO.Path]::GetFullPath($OutputFolder)

if ($SkipMimics -and $SkipLibraries) { throw 'Nothing to do: both -SkipMimics and -SkipLibraries were given.' }

# --- Usage disclaimer -----------------------------------------------------
function Confirm-Disclaimer {
    param([switch]$Accepted)

    $disclaimer = @'
------------------------------------------------------------------------------
 GeoSCADA-Script-Exporter.ps1  Copyright (c) 2026  Adam Woodland
 Licensed under the MIT License. This is free software, and you are welcome to
 redistribute it under those conditions; it comes with ABSOLUTELY NO WARRANTY.

 DISCLAIMER
 This script is provided "AS IS", WITHOUT WARRANTY OF ANY KIND, express or
 implied. The author accepts no liability for any damages arising from its use.
 It connects to the Geo SCADA server through the .NET Client API to read mimic
 and script library definitions and writes text files. It is NOT certified for
 production SCADA systems. Do NOT run it against a production or
 safety-critical system without first reviewing the code and testing on a
 representative non-production system. You run it at your own risk and remain
 responsible for your own change-control and security policies.
------------------------------------------------------------------------------
'@

    if ($Accepted) { Write-Verbose 'Disclaimer accepted via -AcceptDisclaimer.'; return $true }

    $canPrompt = [Environment]::UserInteractive
    try { if ([Console]::IsInputRedirected) { $canPrompt = $false } } catch { }

    if (-not $canPrompt) {
        Write-Host $disclaimer -ForegroundColor Yellow
        throw 'Disclaimer not accepted (non-interactive session). Re-run with -AcceptDisclaimer to confirm acceptance.'
    }

    Write-Host ''
    Write-Host $disclaimer -ForegroundColor Yellow
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription '&Yes', 'I accept the terms and have tested appropriately.'
    $no  = New-Object System.Management.Automation.Host.ChoiceDescription '&No',  'Do not run.'
    $choices = [System.Management.Automation.Host.ChoiceDescription[]]@($yes, $no)
    try {
        $decision = $Host.UI.PromptForChoice('Disclaimer', 'Do you accept these terms and confirm you have tested appropriately?', $choices, 1)
    } catch {
        throw 'Disclaimer not accepted (host could not prompt). Re-run with -AcceptDisclaimer to confirm acceptance.'
    }
    if ($decision -ne 0) { Write-Warning 'Disclaimer not accepted. Exiting without running.'; return $false }
    return $true
}

# --- Languages ------------------------------------------------------------
# Keyed by upper-case braced GUID. Comment is the line-comment token used for the file header.
$Languages = @{
    '{B54F3741-5B07-11CF-A4B0-00AA004A55E8}' = @{ Name = 'VBScript';     Comment = "'"  }
    '{F414C260-6AC0-11CF-B6D1-00AA00BBBB58}' = @{ Name = 'JScript';      Comment = '//' }
    '{16D51579-A30B-4C8B-A276-0FF4DC41E755}' = @{ Name = 'JScript9';     Comment = '//' }
    '{16BA25BB-FB03-48BA-AECA-C53BF6ABD2A1}' = @{ Name = 'JavaScriptV8'; Comment = '//' }
}
$DefaultLanguageGuid = '{B54F3741-5B07-11CF-A4B0-00AA004A55E8}'   # Geo SCADA treats "no language" as VBScript

function Get-LanguageInfo([string]$Guid) {
    <# Returns @{ Name; Comment; Guid } for a language GUID ('' / null GUID => VBScript default). #>
    if ([string]::IsNullOrWhiteSpace($Guid) -or $Guid -match '^\{?0{8}-0{4}-0{4}-0{4}-0{12}\}?$') {
        $l = $Languages[$DefaultLanguageGuid]
        return @{ Name = $l.Name; Comment = $l.Comment; Guid = $DefaultLanguageGuid; Defaulted = $true }
    }
    $g = $Guid.Trim().ToUpperInvariant()
    if (-not $g.StartsWith('{')) { $g = '{' + $g + '}' }
    if ($Languages.ContainsKey($g)) { return @{ Name = $Languages[$g].Name; Comment = $Languages[$g].Comment; Guid = $g; Defaulted = $false } }
    return @{ Name = 'Unknown'; Comment = '//'; Guid = $g; Defaulted = $false }
}

# --- Binary helpers (only for the library language GUID and -CheckMissingLibraries) ------------
if (-not ([System.Management.Automation.PSTypeName]'GeoScadaScriptDoc').Type) {
    Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Collections.Generic;
public static class GeoScadaScriptDoc {
    public static int IndexOf(byte[] hay, byte[] needle, int start) {
        if (hay == null || needle == null || needle.Length == 0) return -1;
        for (int i = Math.Max(0, start); i <= hay.Length - needle.Length; i++) {
            if (hay[i] != needle[0]) continue;
            int j = 1;
            while (j < needle.Length && hay[i + j] == needle[j]) j++;
            if (j == needle.Length) return i;
        }
        return -1;
    }
    // Stored documents start: uint16 version, uint16 class-name length (chars), class name (UTF-16).
    public static string ClassName(byte[] d) {
        if (d == null || d.Length < 4) return null;
        int n = BitConverter.ToUInt16(d, 2);
        if (4 + 2 * n > d.Length) return null;
        return Encoding.Unicode.GetString(d, 4, 2 * n);
    }
    // CScriptLibrary record: class name, 2 bytes, uint32 record length, uint16 version, 16-byte language GUID, ...
    public static string LibraryLanguageGuid(byte[] d) {
        string cls = ClassName(d);
        if (cls != "CScriptLibrary") return null;
        int p = 4 + 2 * cls.Length + 8;
        if (p + 16 > d.Length) return null;
        byte[] g = new byte[16]; Array.Copy(d, p, g, 0, 16);
        return new Guid(g).ToString("B").ToUpperInvariant();
    }
    // CMimic document: "ZLIB", uint32 uncompressed length, zlib stream (2-byte header + deflate).
    public static byte[] InflateMimic(byte[] d) {
        int p = IndexOf(d, new byte[] { 0x5A, 0x4C, 0x49, 0x42 }, 0);
        if (p < 0 || p + 10 > d.Length) return null;
        using (var ms = new MemoryStream(d, p + 10, d.Length - p - 10))
        using (var ds = new DeflateStream(ms, CompressionMode.Decompress))
        using (var o = new MemoryStream()) { ds.CopyTo(o); return o.ToArray(); }
    }
    // In the inflated mimic: language GUID, then an MFC Unicode CString (FF FE FF + length) holding
    // the script, then int32 library count and per library: EF BE AD DE, int64 byte length,
    // UTF-16 name, 4-byte trailer.
    public static bool ReadMimicScript(byte[] doc, string guid, out string script, out List<string> libs) {
        script = null; libs = new List<string>();
        if (doc == null) return false;
        byte[] g = new Guid(guid).ToByteArray();
        byte[] needle = new byte[19]; Array.Copy(g, needle, 16); needle[16] = 0xFF; needle[17] = 0xFE; needle[18] = 0xFF;
        int at = IndexOf(doc, needle, 0);
        if (at < 0) return false;
        int c = at + 19;
        if (c >= doc.Length) return false;
        long n = doc[c]; c++;
        if (n == 0xFF) { n = BitConverter.ToUInt16(doc, c); c += 2; if (n == 0xFFFF) { n = BitConverter.ToUInt32(doc, c); c += 4; } }
        if (c + 2 * n > doc.Length) return false;
        script = Encoding.Unicode.GetString(doc, c, (int)(2 * n)); c += (int)(2 * n);
        if (c + 4 > doc.Length) return true;
        int count = BitConverter.ToInt32(doc, c); c += 4;
        if (count < 0 || count > 10000) return false;
        for (int k = 0; k < count; k++) {
            if (c + 12 > doc.Length || doc[c] != 0xEF || doc[c + 1] != 0xBE || doc[c + 2] != 0xAD || doc[c + 3] != 0xDE) return false;
            long bl = BitConverter.ToInt64(doc, c + 4);
            if (bl < 0 || c + 12 + bl > doc.Length) return false;
            libs.Add(Encoding.Unicode.GetString(doc, c + 12, (int)bl));
            c += 12 + (int)bl + 4;
        }
        return true;
    }
}
"@
}

$script:ProgressClock = [System.Diagnostics.Stopwatch]::StartNew()
function Show-Progress([string]$Activity, [string]$Status, [int]$Percent = -1) {
    <# At most one Write-Progress every 2 s of work: in some hosts each call itself takes seconds. #>
    if ($script:ProgressClock.ElapsedMilliseconds -lt 2000) { return }
    if ($Percent -ge 0) { Write-Progress -Activity $Activity -Status $Status -PercentComplete $Percent }
    else { Write-Progress -Activity $Activity -Status $Status }
    $script:ProgressClock.Restart()      # after the call, so a slow call does not count as work time
}

# --- Text helpers -----------------------------------------------------------
function ConvertTo-CrLf([string]$Text) {
    if ($null -eq $Text) { return '' }
    return ($Text -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n", "`r`n")
}

function Get-LineCount([string]$Text) {
    if ([string]::IsNullOrEmpty($Text)) { return 0 }
    return ($Text -split "`r`n").Count
}

function ConvertTo-HeaderValue($Value) {
    <# Header values are single-line. Not trimmed: Geo SCADA names can end in a space. #>
    if ($null -eq $Value) { return '' }
    return (([string]$Value) -replace "[`r`n]+", ' ')
}

function Format-ConfigTime($Value) {
    if ($null -eq $Value) { return '' }
    if ($Value -is [DateTimeOffset]) { return $Value.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ') }
    if ($Value -is [DateTime])       { return $Value.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
    return [string]$Value
}

function Resolve-RelativeName {
    <# ".Scripts.X" relative to "A.B.Mimic" -> "A.Scripts.X" (each leading '.' strips one level, as ViewX does). #>
    param([string]$Ref, [string]$ParentFullName)
    $dots = 0
    while ($dots -lt $Ref.Length -and $Ref[$dots] -eq '.') { $dots++ }
    if ($dots -eq 0) { return $Ref }
    $rest  = $Ref.Substring($dots)
    $parts = $ParentFullName.Split('.')
    $keep  = $parts.Count - $dots
    if ($keep -le 0) { return $rest }
    return (($parts[0..($keep - 1)] -join '.') + '.' + $rest)
}

# --- File naming ------------------------------------------------------------
$script:UsedFileNames = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
# Windows device names, also reserved with an extension ("CON.Child.txt") or trailing spaces ("CON .x").
$ReservedNames = '^(CON|PRN|AUX|NUL|COM[0-9¹²³]|LPT[0-9¹²³])\s*(\.|$)'

function Get-ExportFileName {
    <# The object's full name made safe for Windows; unique within this run (case-insensitive). #>
    param([string]$FullName, [int]$Id)
    $n = $FullName
    foreach ($ch in [System.IO.Path]::GetInvalidFileNameChars()) { $n = $n.Replace([string]$ch, '_') }
    $n = $n.TrimEnd(' ', '.')
    if ($n -eq '') { $n = '_' }
    if ($n -match $ReservedNames) { $n = '_' + $n }
    $candidate = $n + '.txt'
    if (-not $script:UsedFileNames.Add($candidate)) {
        $candidate = "$n [$Id].txt"
        [void]$script:UsedFileNames.Add($candidate)
    }
    return $candidate
}

$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:LongPaths = 0
function Write-TextFile([string]$Path, [string]$Text) {
    # The \\?\ prefix lets .NET write paths longer than 260 characters.
    $p = $Path
    if ($Path.Length -ge 248 -and -not $Path.StartsWith('\\')) { $p = '\\?\' + $Path; $script:LongPaths++ }
    [System.IO.File]::WriteAllText($p, $Text, $Utf8NoBom)
}

function New-ExportText {
    <#
      Builds the file: header lines (each prefixed with the language's comment token) between the
      BEGIN/END markers, then the script, then one section per inline pick-action script:
          <comment> ===== GEOSCADA-SCRIPT-EXPORT PICKACTION <n> =====
          <code>
      Header entries are an ordered list of Key/Value pairs; a key may repeat (Library, PickAction ...).
      PickActionCode, ScriptStartLine and ScriptEndLine are computed here and added last.
    #>
    param([string]$Comment, [System.Collections.ArrayList]$Entries, [string]$Script, $Sections = @())
    $body      = ConvertTo-CrLf $Script
    $bodyLines = Get-LineCount $body
    $sectionCount = if ($null -ne $Sections) { $Sections.Count } else { 0 }
    $headerCount = 1 + $Entries.Count + $sectionCount + 3               # BEGIN, entries, PickActionCode, Start, End, END
    $scriptStart = $headerCount + 1
    $scriptEnd   = $scriptStart + $bodyLines - 1
    $next = $scriptEnd + 1                                               # line of the next section marker
    $codeEntries = New-Object System.Collections.Generic.List[string]
    $sectionText = New-Object System.Collections.Generic.List[string]
    foreach ($s in $Sections) {
        $code = ConvertTo-CrLf $s.Code
        $cl = Get-LineCount $code
        $codeEntries.Add("$Comment PickActionCode: $($s.Id) | $($next + 1) | $($next + $cl)")
        $sectionText.Add("$Comment ===== GEOSCADA-SCRIPT-EXPORT PICKACTION $($s.Id) =====")
        if ($cl) { $sectionText.Add($code) }
        $next += 1 + $cl
    }
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("$Comment ===== GEOSCADA-SCRIPT-EXPORT BEGIN =====")
    foreach ($e in $Entries) { $lines.Add("$Comment $($e.Key): $(ConvertTo-HeaderValue $e.Value)") }
    foreach ($c in $codeEntries) { $lines.Add($c) }
    $lines.Add("$Comment ScriptStartLine: $scriptStart")
    $lines.Add("$Comment ScriptEndLine: $scriptEnd")
    $lines.Add("$Comment ===== GEOSCADA-SCRIPT-EXPORT END =====")
    if ($bodyLines) { $lines.Add($body) }
    foreach ($t in $sectionText) { $lines.Add($t) }
    return ($lines -join "`r`n")
}

function Add-Entry([System.Collections.ArrayList]$List, [string]$Key, $Value) { [void]$List.Add([pscustomobject]@{ Key = $Key; Value = $Value }) }

# Spreadsheet formula-injection guard for the manifest. Names, users and messages come from the database;
# a text cell starting with = + - @ (or tab / CR) would be evaluated by Excel, so it is written with a
# leading ' (shown as text). GeoSCADA-Script-Analyser.ps1 removes the ' again when it reads the manifest.
function ConvertTo-SafeCsvRow {
    process {
        $row = $_; $copy = $null
        foreach ($p in $row.PSObject.Properties) {
            $v = $p.Value
            if ($v -is [string] -and $v.Length -and "=+-@`t`r".IndexOf($v[0]) -ge 0) {
                if (-not $copy) { $copy = $row.PSObject.Copy() }
                $copy.($p.Name) = "'" + $v
            }
        }
        if ($copy) { $copy } else { $row }
    }
}

# --- Geo SCADA API helpers -----------------------------------------------
function Resolve-ClientDll {
    if ($ClientDllPath) {
        if (-not (Test-Path -LiteralPath $ClientDllPath)) { throw "ClearScada.Client.dll not found at '$ClientDllPath'." }
        return (Resolve-Path -LiteralPath $ClientDllPath).Path
    }
    foreach ($dir in @("$env:ProgramFiles\Schneider Electric\ClearSCADA Client", "$env:ProgramFiles\Schneider Electric\ClearSCADA")) {
        $p = Join-Path $dir 'ClearScada.Client.dll'
        if (Test-Path -LiteralPath $p) { return $p }
    }
    throw 'ClearScada.Client.dll not found. Install the Geo SCADA client or pass -ClientDllPath.'
}

function Invoke-GeoQuery {
    <# Runs SQL through the Advanced API and returns the rows (each a ReadOnlyCollection of values). #>
    param([string]$Sql)
    $q = $script:Conn.Server.PrepareQuery($Sql, (New-Object ClearScada.Client.Advanced.QueryParseParameters))
    try {
        $r = $q.ExecuteSync((New-Object ClearScada.Client.Advanced.QueryExecuteParameters))
        if ($r.Status -ne [ClearScada.Client.Advanced.QueryStatus]::Succeeded -and
            $r.Status -ne [ClearScada.Client.Advanced.QueryStatus]::NoDataFound) { throw "Query failed ($($r.Status)): $Sql" }
        return @($r.Rows | ForEach-Object { , $_.Data })
    } finally { $q.Dispose() }
}

function Get-ObjectList([string]$Class) {
    <# WITH TEMPLATES includes objects inside template definitions, which a plain SELECT leaves out. #>
    $rows = Invoke-GeoQuery "SELECT WITH TEMPLATES Id, FullName, ConfigVersion, ConfigTime, ConfigUser, ParentTemplateId, ParentInstanceId FROM $Class ORDER BY FullName"
    $list = foreach ($d in $rows) {
        [pscustomobject]@{
            Id = [int]$d[0]; FullName = [string]$d[1]; Class = $Class; InTemplate = ([int]$d[5] -gt 0)
            ConfigVersion = $d[2]; ConfigTime = (Format-ConfigTime $d[3]); ConfigUser = [string]$d[4]
            TemplateId = [int]$d[5]; InstanceId = [int]$d[6]
        }
    }
    return @($list | Where-Object { Test-InScope $_.FullName })
}

function Test-InScope([string]$FullName) {
    if (-not $Group -and -not $Name) { return $true }
    foreach ($n in $Name)  { if ([string]::Equals($FullName, $n.Trim(), [StringComparison]::OrdinalIgnoreCase)) { return $true } }
    foreach ($g in $Group) {
        $g = $g.Trim().TrimEnd('.')
        if ([string]::Equals($FullName, $g, [StringComparison]::OrdinalIgnoreCase) -or
            $FullName.StartsWith($g + '.', [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Get-Document([int]$Id, [switch]$Web) {
    <# Returns @{ Bytes; Details } - the document content and the server's ObjectDetails for the object. #>
    $oid = New-Object ClearScada.Client.ObjectId($Id)
    $ms = New-Object System.IO.MemoryStream
    try {
        $details = if ($Web) { $script:Conn.Server.GetWebDocumentContent($oid, $null, $ms) }
                   else      { $script:Conn.Server.GetDocumentContent($oid, $null, $ms) }
        return @{ Bytes = $ms.ToArray(); Details = $details }
    } finally { $ms.Dispose() }
}

function ConvertTo-MimicXml([byte[]]$Bytes) {
    <# GetWebDocumentContent: int32 character count, then the XML as UTF-16. #>
    if ($Bytes.Length -lt 4) { throw 'Server returned no mimic content.' }
    $n = [BitConverter]::ToInt32($Bytes, 0)
    if ($n -le 0 -or 4 + 2 * $n -gt $Bytes.Length) { throw "Unexpected mimic content framing ($($Bytes.Length) bytes)." }
    $text = [System.Text.Encoding]::Unicode.GetString($Bytes, 4, 2 * $n).TrimEnd([char]0)
    $xml = New-Object System.Xml.XmlDocument
    $xml.PreserveWhitespace = $true     # keep the script's own whitespace and CRLFs
    $xml.XmlResolver = $null
    $xml.LoadXml($text)
    return $xml
}

$script:TemplateObjectCache = @{}
function Resolve-EditableSource($Details) {
    <#
      Where the object's content is actually maintained. An object inside an instance takes its
      content from the matching object in the instance's template unless the instance overrides
      it. That template object may itself sit inside an instance (a template that uses another
      template), so the chain is followed with IObjectHierarchy.GetTemplateObject until an object
      that is not controlled by an instance, or that overrides the content, is reached.
      Returns @{ Controlled; Override; Chain (template full names, in order); Source; SourceId; Note }.
    #>
    $r = @{ Controlled = [bool]$Details.IsControlledByInstance; Override = [bool]$Details.HasDocumentContentOverride
            Chain = @(); Source = [string]$Details.FullName; SourceId = $Details.Id.ToInt32(); Note = $null }
    $cur = $Details
    $guard = 0
    while ($cur.IsControlledByInstance -and -not $cur.HasDocumentContentOverride) {
        if (++$guard -gt 50) { $r.Note = 'Template chain longer than 50 levels; stopped'; break }
        $key = $cur.Id.ToInt32()
        if (-not $script:TemplateObjectCache.ContainsKey($key)) {
            try { $script:TemplateObjectCache[$key] = $script:Conn.Server.GetTemplateObject($cur.Id) }
            catch { $script:TemplateObjectCache[$key] = $null; $script:TemplateObjectCache["err$key"] = (Get-InnerMessage $_) }
        }
        $next = $script:TemplateObjectCache[$key]
        if ($null -eq $next) { $r.Note = "Template object for '$($cur.FullName)' could not be resolved: $($script:TemplateObjectCache["err$key"])"; break }
        $r.Chain += [string]$next.FullName
        $r.Source = [string]$next.FullName; $r.SourceId = $next.Id.ToInt32()
        $cur = $next
    }
    return $r
}

function Add-TemplateEntries([System.Collections.ArrayList]$E, $Obj, $Src) {
    Add-Entry $E 'InTemplate' $Obj.InTemplate
    Add-Entry $E 'ControlledByInstance' $Src.Controlled
    Add-Entry $E 'DocumentContentOverride' $Src.Override
    foreach ($t in $Src.Chain) { Add-Entry $E 'TemplateSource' $t }
    Add-Entry $E 'EditableSource' $Src.Source
    Add-Entry $E 'EditableSourceId' $Src.SourceId
    Add-Entry $E 'EditInPlace' ($Src.Chain.Count -eq 0 -and -not $Src.Note)
    if ($Src.Note) { Add-Entry $E 'EditableSourceNote' $Src.Note }
}

function Get-MimicScriptInfo([System.Xml.XmlDocument]$Xml) {
    <#
      From the mimic XML returns @{ HasScriptElement; LangGuid; Script; Libraries; PickActions }.
      The server writes CRLF before each <Library> element and CRLF after the last one, then the
      script text; so the script is the text after the last <Library> minus that one CRLF.
    #>
    $root = $Xml.DocumentElement
    $info = @{ HasScriptElement = $false; LangGuid = ''; Script = ''; Libraries = @(); PickActions = @() }

    $sc = $root.SelectSingleNode('Script')
    if ($sc) {
        $info.HasScriptElement = $true
        $info.LangGuid = $sc.GetAttribute('lang')
        $libs = @(); $sb = New-Object System.Text.StringBuilder
        foreach ($c in $sc.ChildNodes) {
            if ($c.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                if ($c.LocalName -eq 'Library') { $libs += $c.GetAttribute('name') }
                [void]$sb.Clear()               # only text after the last element is the mimic's own script
            } elseif ($c.NodeType -in @([System.Xml.XmlNodeType]::Text, [System.Xml.XmlNodeType]::CDATA,
                                         [System.Xml.XmlNodeType]::Whitespace, [System.Xml.XmlNodeType]::SignificantWhitespace)) {
                [void]$sb.Append($c.Value)
            }
        }
        $own = $sb.ToString()
        if ($libs.Count -gt 0) {
            if ($own.StartsWith("`r`n")) { $own = $own.Substring(2) } elseif ($own.StartsWith("`n")) { $own = $own.Substring(1) }
        }
        $info.Script = $own
        $info.Libraries = $libs
    }

    # Pick actions that run script:
    #   scriptCall     - calls a named function in this mimic's script  <Action type="scriptCall">Name<StdArguments>...</StdArguments></Action>
    #   scriptCode     - inline code, run in this mimic                  <Action type="scriptCode">code</Action>
    #   scriptTopLevel - inline code, run in the topmost mimic           <Action type="scriptTopLevel">code</Action>
    $acts = @()
    foreach ($a in $root.SelectNodes('//Action')) {
        $type = $a.GetAttribute('type')
        if ($type -notmatch 'script') { continue }
        $text = ''; $afterElement = $false
        foreach ($c in $a.ChildNodes) {
            if ($c.NodeType -eq [System.Xml.XmlNodeType]::Element) { $afterElement = $true; $text = '' }
            elseif ($c.NodeType -in @([System.Xml.XmlNodeType]::Text, [System.Xml.XmlNodeType]::CDATA,
                                       [System.Xml.XmlNodeType]::Whitespace, [System.Xml.XmlNodeType]::SignificantWhitespace)) { $text += $c.Value }
        }
        if ($afterElement) { if ($text.StartsWith("`r`n")) { $text = $text.Substring(2) } elseif ($text.StartsWith("`n")) { $text = $text.Substring(1) } }
        $context = switch ($type) { 'scriptCall' { 'Current' } 'scriptCode' { 'Current' } 'scriptTopLevel' { 'TopLevel' } default { 'Unknown' } }
        if ($type -eq 'scriptCall') { $fn = $text.Trim(); $code = $null }
        else { $fn = '(code)'; $code = $text.TrimEnd() }
        $argList = @()
        $std = $a.SelectSingleNode('StdArguments')
        if ($std) {
            $anims = @{}
            foreach ($an in $std.SelectNodes('Animation')) { $anims[$an.GetAttribute('property')] = $an.InnerText.Trim() }
            foreach ($ar in $std.SelectNodes('Arg')) {
                $v = $ar.InnerText.Trim()
                $argList += $(if ($anims.ContainsKey($v)) { "$v <= $($anims[$v])" } else { $v })
            }
        }
        $acts += [pscustomobject]@{ Where = (Get-ElementPath $a.ParentNode $root); Action = $a.GetAttribute('name'); Type = $type
                                    Context = $context; Function = $fn; Code = $code; Args = ($argList -join ', ') }
    }
    $info.PickActions = $acts

    # Embedded mimics: <EmbeddedMimic name="M_1" ref="SCX:////CMimic/.Test"> (the ref may be relative).
    $info.Embeds = @(foreach ($em in $root.SelectNodes('//EmbeddedMimic')) {
        [pscustomobject]@{ Where = (Get-ElementPath $em $root); Ref = $em.GetAttribute('ref') }
    })
    return $info
}

function Get-ElementPath($Node, $Root) {
    <# Names of the enclosing layer / groups / object (and sub-menus) from the mimic down to Node, e.g. "Layer[Layer]/Btn_3[Button]". #>
    $path = @(); $n = $Node
    while ($n -and $n -ne $Root -and $n.NodeType -eq [System.Xml.XmlNodeType]::Element) {
        if ($n.LocalName -ne 'PickMenu' -and $n.HasAttribute('name')) { $path = @("$($n.GetAttribute('name'))[$($n.LocalName)]") + $path }
        $n = $n.ParentNode
    }
    if ($path.Count) { return ($path -join '/') } return '(mimic)'
}

function Resolve-MimicRef([string]$Ref, [string]$MimicFullName) {
    <# "SCX:////CMimic/.Test" in "Claude.Default" -> "Claude.Test". Returns '' for an empty (animated) reference. #>
    if ([string]::IsNullOrWhiteSpace($Ref)) { return '' }
    $name = $Ref
    if ($name -match '^[A-Za-z]+:/+[^/]*/(.*)$') { $name = $Matches[1] }
    elseif ($name.Contains('/')) { $name = $name.Substring($name.LastIndexOf('/') + 1) }
    return (Resolve-RelativeName -Ref $name -ParentFullName $MimicFullName)
}

function Format-Field([string]$Value) { <# A ' | '-separated header field must not itself contain the separator. #> return ([string]$Value).Replace(' | ', ' / ') }

# --- Main -----------------------------------------------------------------
if (-not (Confirm-Disclaimer -Accepted:$AcceptDisclaimer)) { return }

if ($PSVersionTable.PSVersion.Major -ge 6) {
    Write-Warning 'ClearScada.Client.dll targets .NET Framework 4.8; this script is tested on Windows PowerShell 5.1.'
}

if (-not $Credential -and $Username) {
    $sec = if ($Password) { ConvertTo-SecureString $Password -AsPlainText -Force } else { New-Object System.Security.SecureString }
    $Credential = New-Object System.Management.Automation.PSCredential($Username, $sec)
}
if (-not $Credential) {
    $canPrompt = [Environment]::UserInteractive
    try { if ([Console]::IsInputRedirected) { $canPrompt = $false } } catch { }
    if (-not $canPrompt) { throw 'No credentials. Supply -Credential or -Username/-Password.' }
    $Credential = Get-Credential -Message "Geo SCADA logon for $Server"
    if (-not $Credential) { throw 'No credentials supplied.' }
}

$dll = Resolve-ClientDll
Write-Verbose "Using $dll"
Add-Type -Path $dll

if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }

$exportedAt = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
$manifest   = New-Object System.Collections.Generic.List[object]
$written    = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$script:Errors = 0
$stats = [ordered]@{ MimicsScanned = 0; MimicFiles = 0; MimicsWithOwnScript = 0; LibraryFiles = 0; LibrariesEmpty = 0; MissingLibraryLinks = 0; CrossCheckMismatches = 0; PickActionCode = 0 }
$libraryUse = @{}     # library full name -> number of in-scope mimics linking to it

function Add-ManifestRow {
    param($Obj, [string]$Status, [string]$Message, [hashtable]$Extra)
    $row = [ordered]@{
        FullName = $Obj.FullName; Class = $Obj.Class; Id = $Obj.Id; Status = $Status; File = ''
        Language = ''; LanguageGuid = ''; ScriptLines = 0; ScriptChars = 0; HasOwnScript = $false
        LibraryCount = 0; Libraries = ''; MissingLibraries = ''; PickActionCount = 0; ReferencedByMimics = ''
        ConfigVersion = $Obj.ConfigVersion; ConfigTime = $Obj.ConfigTime; ConfigUser = $Obj.ConfigUser
        TemplateId = $(if ($Obj.TemplateId -gt 0) { $Obj.TemplateId } else { '' })
        InstanceId = $(if ($Obj.InstanceId -gt 0) { $Obj.InstanceId } else { '' })
        InTemplate = $Obj.InTemplate; ControlledByInstance = ''; DocumentContentOverride = ''; EditableSource = ''; EditableSourceId = ''
        Message = $Message
    }
    if ($Extra) { foreach ($k in $Extra.Keys) { $row[$k] = $Extra[$k] } }
    $manifest.Add([pscustomobject]$row)
}

function Get-InnerMessage($Err) {
    <# The API's own message, without PowerShell's "Exception calling ... with N argument(s)" wrapper. #>
    $ex = $Err.Exception
    if ($ex -is [System.Management.Automation.MethodInvocationException] -and $ex.InnerException) { $ex = $ex.InnerException }
    return $ex.Message
}

function Register-Failure($Obj, [string]$What, $Err) {
    $script:Errors++
    $inner = Get-InnerMessage $Err
    $msg = "$What failed for '$($Obj.FullName)': $inner"
    Write-Warning $msg
    Add-ManifestRow -Obj $Obj -Status 'Error' -Message $inner
    if ($FailOnError) { throw $msg }
}

function Add-CommonEntries([System.Collections.ArrayList]$E, $Obj, $Lang, [string]$Source) {
    Add-Entry $E 'FormatVersion' $FormatVersion
    Add-Entry $E 'FullName'      $Obj.FullName
    Add-Entry $E 'Class'         $Obj.Class
    Add-Entry $E 'ObjectId'      $Obj.Id
    Add-Entry $E 'Language'      $Lang.Name
    Add-Entry $E 'LanguageGuid'  $Lang.Guid
    if ($Lang.Defaulted) { Add-Entry $E 'LanguageNote' 'No language stored; Geo SCADA default (VBScript) assumed' }
    Add-Entry $E 'ConfigVersion' $Obj.ConfigVersion
    Add-Entry $E 'ConfigTime'    $Obj.ConfigTime
    Add-Entry $E 'ConfigUser'    $Obj.ConfigUser
    if ($Obj.TemplateId -gt 0) { Add-Entry $E 'TemplateId' $Obj.TemplateId }
    if ($Obj.InstanceId -gt 0) { Add-Entry $E 'InstanceId' $Obj.InstanceId }
    Add-Entry $E 'Source'        $Source
    Add-Entry $E 'Server'        $Server
    Add-Entry $E 'ExportedAt'    $exportedAt
    Add-Entry $E 'Exporter'      "GeoSCADA-Script-Exporter $ExporterVersion"
}

$exitCode = 0
$script:Conn = $null
try {
    Write-Host "Connecting to Geo SCADA at ${Server}:$Port ..." -ForegroundColor Cyan
    $script:Conn = New-Object ClearScada.Client.Simple.Connection('GeoSCADA-Script-Exporter')
    try { $script:Conn.Connect((New-Object ClearScada.Client.ServerNode($Server, $Port))) }
    catch { throw "Could not connect to Geo SCADA at ${Server}:${Port}: $(Get-InnerMessage $_)" }
    try { $script:Conn.LogOn($Credential.UserName, $Credential.Password) }
    catch { throw "Logon as '$($Credential.UserName)' failed: $(Get-InnerMessage $_)" }
    Write-Host "Logged on as $($Credential.UserName)." -ForegroundColor Green

    # Library names are needed even when only mimics are exported (to tell which links resolve).
    $libraries = @(Get-ObjectList 'CScriptLibrary')
    $mimics    = if ($SkipMimics) { @() } else { @(Get-ObjectList 'CMimic') }
    Write-Host ("In scope: {0} mimic(s), {1} script librar{2}." -f $mimics.Count, $libraries.Count, $(if ($libraries.Count -eq 1) { 'y' } else { 'ies' })) -ForegroundColor Cyan
    Write-Host "Writing to $OutputFolder" -ForegroundColor Cyan

    # ---- Mimics ----
    $i = 0
    foreach ($m in $mimics) {
        $i++
        if ($i % 25 -eq 1 -or $i -eq $mimics.Count) {
            Show-Progress 'Exporting mimic scripts' "$i of $($mimics.Count): $($m.FullName)" ([int](100 * $i / [Math]::Max(1, $mimics.Count)))
        }
        $stats.MimicsScanned++
        try {
            $webDoc = Get-Document -Id $m.Id -Web
            $info = Get-MimicScriptInfo (ConvertTo-MimicXml $webDoc.Bytes)
        } catch { Register-Failure $m 'Reading mimic' $_; continue }

        $missing = @(); $checkNote = $null
        if ($CheckMissingLibraries -and $info.HasScriptElement) {
            try {
                $doc = [GeoScadaScriptDoc]::InflateMimic((Get-Document -Id $m.Id).Bytes)
                $lg  = (Get-LanguageInfo $info.LangGuid).Guid
                $binScript = $null; $binLibs = $null
                if ([GeoScadaScriptDoc]::ReadMimicScript($doc, $lg, [ref]$binScript, [ref]$binLibs)) {
                    $resolved = @($binLibs | ForEach-Object { Resolve-RelativeName -Ref $_ -ParentFullName $m.FullName })
                    $missing  = @($resolved | Where-Object { $info.Libraries -notcontains $_ })
                    if ((ConvertTo-CrLf $binScript) -ne (ConvertTo-CrLf $info.Script)) {
                        $checkNote = 'Stored script differs from the XML script (XML exported)'; $stats.CrossCheckMismatches++
                    }
                } else { $checkNote = 'Stored document could not be decoded; missing-library check not done' }
            } catch { $checkNote = "Stored document check failed: $($_.Exception.Message)" }
            if ($checkNote) { Write-Warning "$($m.FullName): $checkNote" }
        }

        $hasOwn = -not [string]::IsNullOrWhiteSpace($info.Script)
        if (-not $hasOwn -and $info.Libraries.Count -eq 0 -and $missing.Count -eq 0 -and $info.PickActions.Count -eq 0) { continue }

        foreach ($l in $info.Libraries) { $libraryUse[$l] = 1 + [int]$libraryUse[$l] }
        $stats.MissingLibraryLinks += $missing.Count
        if ($hasOwn) { $stats.MimicsWithOwnScript++ }

        $lang = Get-LanguageInfo $info.LangGuid
        if (-not $info.HasScriptElement) { $lang = @{ Name = 'None'; Comment = "'"; Guid = ''; Defaulted = $false } }
        $code = ConvertTo-CrLf $info.Script
        $src  = Resolve-EditableSource $webDoc.Details
        $e = New-Object System.Collections.ArrayList
        Add-CommonEntries $e $m $lang 'GetWebDocumentContent'
        Add-TemplateEntries $e $m $src
        Add-Entry $e 'HasOwnScript' $hasOwn
        Add-Entry $e 'ScriptLines'  (Get-LineCount $code)
        Add-Entry $e 'ScriptChars'  $code.Length
        Add-Entry $e 'LibraryCount' $info.Libraries.Count
        foreach ($l in $info.Libraries) { Add-Entry $e 'Library' $l }
        Add-Entry $e 'LibraryCheck' $(if (-not $CheckMissingLibraries) { 'NotPerformed (XML lists only libraries that exist)' } elseif ($checkNote) { "Incomplete: $checkNote" } else { 'Performed' })
        foreach ($l in $missing) { Add-Entry $e 'LibraryMissing' $l }
        Add-Entry $e 'PickActionCount' $info.PickActions.Count
        # Seven ' | '-separated fields: N | Where | Action | Type | Context | Function | Args. Args is last so it
        # may itself contain ' | '. Inline code (Function "(code)") is written to section N after the script.
        $sections = New-Object System.Collections.Generic.List[object]
        $k = 0
        foreach ($p in $info.PickActions) {
            $k++
            Add-Entry $e 'PickAction' ("{0} | {1} | {2} | {3} | {4} | {5} | {6}" -f $k, (Format-Field $p.Where), (Format-Field $p.Action), $p.Type, $p.Context,
                                       (Format-Field $p.Function), $(if ($p.Args) { $p.Args } else { '(none)' }))
            if ($null -ne $p.Code) { $sections.Add(@{ Id = $k; Code = $p.Code }); $stats.PickActionCode++ }
        }
        # Embedded mimics: Where | full name (relative references resolved; "(animated)" when set by animation).
        Add-Entry $e 'EmbeddedMimicCount' $info.Embeds.Count
        foreach ($em in $info.Embeds) {
            $target = Resolve-MimicRef $em.Ref $m.FullName
            Add-Entry $e 'EmbeddedMimic' ("{0} | {1}" -f (Format-Field $em.Where), $(if ($target) { $target } else { '(animated)' }))
        }

        try {
            $file = Get-ExportFileName -FullName $m.FullName -Id $m.Id
            Write-TextFile -Path (Join-Path $OutputFolder $file) -Text (New-ExportText -Comment $lang.Comment -Entries $e -Script $code -Sections $sections)
            [void]$written.Add($file)
            $stats.MimicFiles++
            Add-ManifestRow -Obj $m -Status 'Exported' -Message $checkNote -Extra @{
                File = $file; Language = $lang.Name; LanguageGuid = $lang.Guid; ScriptLines = (Get-LineCount $code); ScriptChars = $code.Length
                HasOwnScript = $hasOwn; LibraryCount = $info.Libraries.Count; Libraries = ($info.Libraries -join '; ')
                MissingLibraries = ($missing -join '; '); PickActionCount = $info.PickActions.Count
                ControlledByInstance = $src.Controlled; DocumentContentOverride = $src.Override
                EditableSource = $src.Source; EditableSourceId = $src.SourceId }
        } catch { Register-Failure $m 'Writing file' $_ }
    }
    Write-Progress -Activity 'Exporting mimic scripts' -Completed

    # ---- Script libraries ----
    if (-not $SkipLibraries) {
        $i = 0
        foreach ($l in $libraries) {
            $i++
            Write-Progress -Activity 'Exporting script libraries' -Status "$i of $($libraries.Count): $($l.FullName)" -PercentComplete ([int](100 * $i / [Math]::Max(1, $libraries.Count)))
            try {
                $libSrc = [string]$script:Conn.GetObject((New-Object ClearScada.Client.ObjectId($l.Id))).InvokeMethod('GetScriptSource', [object[]]@())
                $libDoc = Get-Document -Id $l.Id
                $guid   = [GeoScadaScriptDoc]::LibraryLanguageGuid($libDoc.Bytes)
            } catch { Register-Failure $l 'Reading script library' $_; continue }
            $src = Resolve-EditableSource $libDoc.Details

            $lang = if ($guid) { Get-LanguageInfo $guid } else { @{ Name = 'Unknown'; Comment = '//'; Guid = ''; Defaulted = $false } }
            $refs = if ($SkipMimics) { '' } else { [int]$libraryUse[$l.FullName] }
            $tmplCols = @{ ControlledByInstance = $src.Controlled; DocumentContentOverride = $src.Override; EditableSource = $src.Source; EditableSourceId = $src.SourceId }
            if ([string]::IsNullOrWhiteSpace($libSrc)) {
                $stats.LibrariesEmpty++
                Add-ManifestRow -Obj $l -Status 'Empty' -Message 'Library has no code; no file written' -Extra (@{ Language = $lang.Name; LanguageGuid = $lang.Guid; ReferencedByMimics = $refs } + $tmplCols)
                continue
            }
            $code = ConvertTo-CrLf $libSrc
            $e = New-Object System.Collections.ArrayList
            Add-CommonEntries $e $l $lang 'GetScriptSource'
            if (-not $guid) { Add-Entry $e 'LanguageNote' 'Language GUID could not be read from the stored record' }
            Add-TemplateEntries $e $l $src
            Add-Entry $e 'ScriptLines' (Get-LineCount $code)
            Add-Entry $e 'ScriptChars' $code.Length
            Add-Entry $e 'ReferencedByMimics' $(if ($SkipMimics) { 'NotChecked (-SkipMimics)' } elseif ($Group -or $Name) { "$refs (in scope only)" } else { $refs })
            try {
                $file = Get-ExportFileName -FullName $l.FullName -Id $l.Id
                Write-TextFile -Path (Join-Path $OutputFolder $file) -Text (New-ExportText -Comment $lang.Comment -Entries $e -Script $code)
                [void]$written.Add($file)
                $stats.LibraryFiles++
                Add-ManifestRow -Obj $l -Status 'Exported' -Message $null -Extra (@{
                    File = $file; Language = $lang.Name; LanguageGuid = $lang.Guid; ScriptLines = (Get-LineCount $code); ScriptChars = $code.Length
                    HasOwnScript = $true; ReferencedByMimics = $refs } + $tmplCols)
            } catch { Register-Failure $l 'Writing file' $_ }
        }
        Write-Progress -Activity 'Exporting script libraries' -Completed
    }

    # ---- Manifest ----
    $manifestPath = Join-Path $OutputFolder '_ScriptExport_Manifest.csv'
    $manifest | ConvertTo-SafeCsvRow | Export-Csv -LiteralPath $manifestPath -NoTypeInformation -Encoding UTF8

    $stale = @(Get-ChildItem -LiteralPath $OutputFolder -Filter '*.txt' -File | Where-Object { -not $written.Contains($_.Name) })

    Write-Host ''
    Write-Host 'Done.' -ForegroundColor Green
    Write-Host ("  Mimics scanned          : {0}" -f $stats.MimicsScanned)
    Write-Host ("  Mimic files written     : {0} ({1} with their own script, {2} inline pick-action scripts)" -f $stats.MimicFiles, $stats.MimicsWithOwnScript, $stats.PickActionCode)
    if (-not $SkipLibraries) {
    Write-Host ("  Library files written   : {0} ({1} empty libraries skipped)" -f $stats.LibraryFiles, $stats.LibrariesEmpty) }
    if ($CheckMissingLibraries) {
    Write-Host ("  Missing library links   : {0}" -f $stats.MissingLibraryLinks)
    Write-Host ("  Cross-check mismatches  : {0}" -f $stats.CrossCheckMismatches) }
    Write-Host ("  Manifest                : {0}" -f $manifestPath)
    if ($script:LongPaths) {
        Write-Warning "$($script:LongPaths) file path(s) are 248+ characters long. Some tools cannot open them; use a shorter -OutputFolder (e.g. C:\Export) if that is a problem."
    }
    if ($stale.Count) {
        Write-Warning "$($stale.Count) .txt file(s) in the output folder were not written by this run (left over from an earlier export or another scope). They were not changed."
    }
    Write-Host ''
    Write-Host 'NOTICE: The exported files contain SCADA script code and object names, which may include' -ForegroundColor Yellow
    Write-Host '        sensitive information (system topology, logic, and occasionally embedded credentials).' -ForegroundColor Yellow
    Write-Host '        Store, share and dispose of them in line with your data protection and security policies.' -ForegroundColor Yellow
    if ($script:Errors) {
        Write-Warning "EXPORT INCOMPLETE: $($script:Errors) object(s) could not be exported (see warnings above and the manifest). Exit code 2."
        $exitCode = 2
    }
}
finally {
    if ($script:Conn) {
        try { if ($script:Conn.IsLoggedOn) { $script:Conn.LogOff() } } catch { Write-Verbose "LogOff failed: $($_.Exception.Message)" }
        try { $script:Conn.Disconnect() } catch { }
        try { $script:Conn.Dispose() } catch { }
    }
}
if ($exitCode) { exit $exitCode }
