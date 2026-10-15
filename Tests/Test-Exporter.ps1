<#
.SYNOPSIS
    Unit tests for pure helper functions in GeoSCADA-Script-Exporter.ps1. No Geo SCADA server needed.

.DESCRIPTION
    The exporter is a script, not a module, so its functions are taken from its syntax tree and
    defined here without running the script. Covered: file naming (invalid characters, Windows
    device names, collisions), the file layout written by New-ExportText (ScriptStartLine /
    ScriptEndLine / PickActionCode point at the right lines, including empty scripts, scripts ending
    in a line break and empty sections), relative library names, and the manifest formula guard.
    Not covered: anything that needs the server or the client DLL (XML extraction from live
    documents, template resolution, binary document parsing, failure / retry handling).
    Returns the number of failed checks.
#>
$exporter = Join-Path (Split-Path -Parent $PSScriptRoot) 'GeoSCADA-Script-Exporter.ps1'
$ast = [System.Management.Automation.Language.Parser]::ParseFile($exporter, [ref]$null, [ref]$null)
$wanted = 'ConvertTo-CrLf', 'Get-LineCount', 'ConvertTo-HeaderValue', 'Resolve-RelativeName', 'Get-ExportFileName', 'New-ExportText', 'Add-Entry', 'ConvertTo-SafeCsvRow'
foreach ($fn in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -in $wanted }, $true)) {
    . ([scriptblock]::Create($fn.Extent.Text))
}
# Script-level variables the functions use
$ReservedNames = ($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$ReservedNames' }, $true) | Select-Object -First 1).Right.Extent.Text.Trim("'")
$script:UsedFileNames = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)

$failed = 0
function Check([string]$Name, [bool]$Ok, [string]$Got) {
    if ($Ok) { Write-Host "PASS exporter $Name" } else { $script:failed++; Write-Host "FAIL exporter $Name (got: $Got)" -ForegroundColor Red }
}

# --- File names ----------------------------------------------------------------
$cases = @(     # full name, expected file name (pairs, not a hashtable: keys differ only in case)
    @('Water.Site 1.Display',   'Water.Site 1.Display.txt'),
    @('A:B*C?"D<E>F|G',         'A_B_C__D_E_F_G.txt'),
    @('Name ends with dots...', 'Name ends with dots.txt'),
    @('CON',                    '_CON.txt'),
    @('aux',                    '_aux.txt'),
    @('CON.Child',              '_CON.Child.txt'),
    @('AUX.Display',            '_AUX.Display.txt'),
    @('LPT1.Script',            '_LPT1.Script.txt'),
    @('NUL .x',                 '_NUL .x.txt'),
    @('COM¹.x',                 '_COM¹.x.txt'),
    @('CONSOLE.Display',        'CONSOLE.Display.txt'),
    @('Site.CON',               'Site.CON.txt')
)
foreach ($c in $cases) {
    $got = Get-ExportFileName -FullName $c[0] -Id 1
    Check "file name '$($c[0])'" ($got -ceq $c[1]) $got
}
$first = Get-ExportFileName -FullName 'Same.Name' -Id 10
$second = Get-ExportFileName -FullName 'same.name' -Id 11
Check 'file name collision (case-insensitive) gets the object id' ($first -eq 'Same.Name.txt' -and $second -eq 'same.name [11].txt') "$first / $second"

# --- File layout ---------------------------------------------------------------
function Test-Layout([string]$Label, [string]$Script, $Sections) {
    $e = New-Object System.Collections.ArrayList
    Add-Entry $e 'FullName' 'T.Layout'
    $text = New-ExportText -Comment "'" -Entries $e -Script $Script -Sections $Sections
    $lines = $text -split "`r`n"
    $get = { param($key) ($lines | Where-Object { $_ -like "' ${key}: *" } | ForEach-Object { $_.Substring($key.Length + 4) }) }
    $start = [int](& $get 'ScriptStartLine'); $end = [int](& $get 'ScriptEndLine')
    $endMarker = [array]::IndexOf($lines, "' ===== GEOSCADA-SCRIPT-EXPORT END =====") + 1
    $body = if ($end -ge $start) { ($lines[($start - 1)..($end - 1)]) -join "`r`n" } else { '' }
    Check "$Label`: script lines" ($start -eq $endMarker + 1 -and $body -eq (ConvertTo-CrLf $Script)) "start $start end $end body '$body'"
    foreach ($pc in @(& $get 'PickActionCode')) {
        $p = $pc -split ' \| '; $id = [int]$p[0]; $a = [int]$p[1]; $b = [int]$p[2]
        $want = ConvertTo-CrLf (($Sections | Where-Object { $_.Id -eq $id }).Code)
        $code = if ($b -ge $a) { ($lines[($a - 1)..($b - 1)]) -join "`r`n" } else { '' }
        Check "$Label`: section $id lines" ($code -eq $want -and $lines[$a - 2] -like "*PICKACTION $id =====") "'$code'"
    }
}
Test-Layout 'empty script' '' @()
Test-Layout 'one-line script' 'MsgBox 1' @()
Test-Layout 'script ending in a line break' "Sub A()`r`nEnd Sub`r`n" @()
Test-Layout 'LF-only script' "Sub A()`nEnd Sub" @()
Test-Layout 'script and sections' "Sub A()`r`nEnd Sub" @(@{ Id = 2; Code = "MsgBox 2`r`nMsgBox 3" }, @{ Id = 5; Code = '' }, @{ Id = 6; Code = 'MsgBox 6' })
Test-Layout 'sections without a script' '' @(@{ Id = 1; Code = 'MsgBox 1' })

# --- Relative names ------------------------------------------------------------
$rel = [ordered]@{ '.Scripts.X|A.B.Mimic' = 'A.B.Scripts.X'; '..Scripts.X|A.B.Mimic' = 'A.Scripts.X'; 'Abs.Lib|A.B.Mimic' = 'Abs.Lib'; '....Lib|A.B' = 'Lib' }
foreach ($k in $rel.Keys) {
    $r, $p = $k -split '\|'
    $got = Resolve-RelativeName -Ref $r -ParentFullName $p
    Check "relative name '$r' from '$p'" ($got -eq $rel[$k]) $got
}

# --- Manifest formula guard ------------------------------------------------------
$rows = @([pscustomobject]@{ FullName = '=HYPERLINK("x")'; ConfigUser = '-user'; Id = -5; Note = 'plain' }, [pscustomobject]@{ FullName = 'Safe'; ConfigUser = '@x'; Id = 1; Note = '+1' })
$safe = @($rows | ConvertTo-SafeCsvRow)
Check 'formula guard prefixes = - @ + text cells' ($safe[0].FullName -eq "'=HYPERLINK(""x"")" -and $safe[0].ConfigUser -eq "'-user" -and $safe[1].ConfigUser -eq "'@x" -and $safe[1].Note -eq "'+1") ($safe | ConvertTo-Json -Compress)
Check 'formula guard leaves numbers, plain text and the original rows alone' ($safe[0].Id -eq -5 -and $safe[0].Note -eq 'plain' -and $rows[0].FullName -eq '=HYPERLINK("x")') ($rows | ConvertTo-Json -Compress)

return $failed
