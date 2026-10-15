<#
.SYNOPSIS
    Regression tests for GeoSCADA-Script-Analyser.ps1. No Geo SCADA server needed.

.DESCRIPTION
    1. Rule tags: builds synthetic export files (New-Fixtures.ps1 = format 1, New-FixturesV2.ps1 =
       format 2), runs the analyser on them and checks every tag in the script bodies, by file line:
         EXPECT:<rule>  the analyser must report <rule> on this line
         NOT:<rule>     the analyser must not report <rule> on this line
    2. Template batching and export check (New-FixturesBatching.ps1): template / instance /
       overridden-instance roles, InstancesFixed, the Actions / Instances CSVs, names starting with
       '-' (CSV formula guard), and one case per export-check issue, with and without the manifest;
       then -ManifestOnly.
    3. Exporter helpers (Test-Exporter.ps1): file names, file layout and line numbers, relative names,
       the manifest formula guard.

    Work files go to -WorkFolder (default %TEMP%\GeoSCADA-Script-Analyser-Tests), never into the repo.
    Exit code 0 when everything passes, 1 otherwise.
#>
param([string]$WorkFolder = (Join-Path $env:TEMP 'GeoSCADA-Script-Analyser-Tests'))

$here = $PSScriptRoot
$analyser = Join-Path (Split-Path -Parent $here) 'GeoSCADA-Script-Analyser.ps1'
New-Item -ItemType Directory -Force $WorkFolder | Out-Null
$failed = 0

# --- 1. Rule tags ------------------------------------------------------------
$in = Join-Path $WorkFolder 'fixtures'; $out = Join-Path $WorkFolder 'fixtures_out'
& "$here\New-Fixtures.ps1" -Out $in
& "$here\New-FixturesV2.ps1" -Out $in
& $analyser -InputFolder $in -OutputFolder $out 6>$null 3>$null | Out-Null
$f = Import-Csv "$out\_ScriptAnalysis.csv"
$pass = 0; $fail = 0
foreach ($file in Get-ChildItem $in -Filter *.txt) {
    $lines = [IO.File]::ReadAllLines($file.FullName)
    $fn = ($lines | Select-String -Pattern 'FullName: (.*)$').Matches[0].Groups[1].Value
    $endH = ($lines | Select-String -Pattern 'GEOSCADA-SCRIPT-EXPORT END').LineNumber
    for ($k = $endH; $k -lt $lines.Count; $k++) {
        $fl = $k + 1
        foreach ($m in [regex]::Matches($lines[$k], '(EXPECT|NOT):(\w+)')) {
            $hit = @($f | Where-Object { $_.FullName -eq $fn -and $_.Rule -eq $m.Groups[2].Value -and [int]$_.FileLine -eq $fl }).Count -gt 0
            if (($m.Groups[1].Value -eq 'EXPECT') -eq $hit) { $pass++ } else { $fail++; Write-Host "FAIL $($m.Value) $fn file line $fl" -ForegroundColor Red }
        }
    }
}
Write-Host "Rule tags: pass=$pass fail=$fail; rules seen: $((($f | Select-Object -ExpandProperty Rule -Unique | Sort-Object) -join ','))"
$failed += $fail

$bad = @($f | Where-Object { $_.FullName -eq 'T.TopLevel No Script' -and $_.Rule -in 'GS02', 'GS03' })
if ($bad.Count) { $failed++; Write-Host "FAIL format-1 inline pick-action code treated as function calls: $($bad.Rule -join ',')" -ForegroundColor Red }
else { Write-Host 'PASS format-1 inline pick-action code (scriptTopLevel/scriptCode) not treated as function calls' }

# --- 2. Template batching ----------------------------------------------------
foreach ($withManifest in $false, $true) {
    $in = Join-Path $WorkFolder 'fixtures_t3'; $out = Join-Path $WorkFolder 'fixtures_t3_out'
    foreach ($d in $in, $out) { if (Test-Path $d) { Remove-Item $d -Recurse -Force }; New-Item -ItemType Directory $d | Out-Null }
    & "$here\New-FixturesBatching.ps1" -Out $in -WithManifest:$withManifest
    & $analyser -InputFolder $in -OutputFolder $out 6>$null 3>$null | Out-Null
    $exit = $LASTEXITCODE
    $f = @(Import-Csv "$out\_ScriptAnalysis.csv" | Where-Object Rule -eq 'RN21')
    $all = @(Import-Csv "$out\_ScriptAnalysis.csv")
    $a = @(Import-Csv "$out\_ScriptAnalysis_Actions.csv")
    $i = @(Import-Csv "$out\_ScriptAnalysis_Instances.csv")
    $c = @(Import-Csv "$out\_ScriptAnalysis_ExportCheck.csv")
    $issue = { param($Issue, $Name) [bool]($c | Where-Object { $_.Issue -eq $Issue -and ($_.FullName -eq $Name -or $_.File -eq $Name) }) }
    $rawActions = [IO.File]::ReadAllText("$out\_ScriptAnalysis_Actions.csv")
    $checks = [ordered]@{
        'two RN21 findings (template + overridden instance only)' = ($f.Count -eq 2)
        'template finding: Role Template, InstancesFixed 2'       = [bool]($f | Where-Object { $_.FullName -eq 'T3.Tmpl.Symbol' -and $_.Role -eq 'Template' -and $_.InstancesFixed -eq '2' })
        'overridden instance reported on its own'                 = [bool]($f | Where-Object { $_.FullName -eq 'T3.Site3.Symbol' -and $_.Role -eq 'OverriddenInstance' -and $_.InstancesFixed -eq '0' })
        'no findings against instance copies'                     = -not ($f | Where-Object { $_.FullName -in 'T3.Site1.Symbol', 'T3.Site2.Symbol' })
        'actions: template action mentions 2 copies'              = [bool]($a | Where-Object { $_.FullName -eq 'T3.Tmpl.Symbol' -and $_.Action -match '2 instance copies' })
        'actions: overridden instance says edit directly'         = [bool]($a | Where-Object { $_.FullName -eq 'T3.Site3.Symbol' -and $_.Action -match 'directly' })
        'instances CSV lists Site1 and Site2 under the template'  = ((@($i | Where-Object EditableSource -eq 'T3.Tmpl.Symbol' | ForEach-Object Instance) -join ',') -eq 'T3.Site1.Symbol,T3.Site2.Symbol')
        "'-' names: template finding counts its instance"          = [bool]($all | Where-Object { $_.FullName -eq "'-T3.Neg.Template" -and $_.Rule -eq 'RN20' -and $_.InstancesFixed -eq '1' })
        "'-' names: written to the CSVs with a leading '"          = ($rawActions.Contains('"''-T3.Neg.Template"') -and -not $rawActions.Contains('"-T3.Neg.Template"'))
        "'-' names: no export-check problem for them"              = -not ($c | Where-Object { $_.FullName -like '*-T3.Neg*' -or $_.File -like '*-T3.Neg*' })
        'export check: truncated file is Malformed'               = (& $issue 'Malformed' 'T3.Trunc.Symbol')
        'export check: ScriptStartLine past the end is Malformed' = (& $issue 'Malformed' 'T3.Short.Symbol')
        'export check: malformed files are not analysed'          = -not ($all | Where-Object { $_.FullName -in 'T3.Trunc.Symbol', 'T3.Short.Symbol' })
        'export check: script ending in a line break is not Malformed' = -not (& $issue 'Malformed' 'T3.TrailingBreak.Symbol')
        'export check: duplicate FullName reported'               = (& $issue 'DuplicateName' 'T3.Old.Symbol')
        'export check: exit code 3 (incomplete export)'           = ($exit -eq 3)
    }
    if ($withManifest) {
        $checks['export check: Exported row with no file is MissingFile'] = (& $issue 'MissingFile' 'T3.Gone.Symbol')
        $checks['export check: manifest Error row is ExportError']       = (& $issue 'ExportError' 'T3.Broken.Symbol')
        $checks['export check: unlisted file is NotInManifest']          = (& $issue 'NotInManifest' 'T3.Old.Symbol.txt')
        $checks['export check: header/manifest name mismatch']           = (& $issue 'ManifestMismatch' 'T3.Mismatch.Symbol')
    } else {
        $checks['export check: no manifest is reported']                 = [bool]($c | Where-Object Issue -eq 'NoManifest')
    }
    foreach ($k in $checks.Keys) {
        if ($checks[$k]) { Write-Host "PASS batching (manifest: $withManifest) $k" }
        else { $failed++; Write-Host "FAIL batching (manifest: $withManifest) $k" -ForegroundColor Red }
    }
}

# --- 3. Exporter helpers (file names, file layout, relative names, formula guard) ---
$r = $null
try { $r = & "$here\Test-Exporter.ps1" } catch { Write-Host "FAIL exporter tests could not run: $($_.Exception.Message)" -ForegroundColor Red }
if ($r -is [int]) { $failed += $r } else { $failed++; Write-Host 'FAIL exporter tests did not report a result (did the script load?)' -ForegroundColor Red }

# -ManifestOnly leaves files that are not in the manifest out of the analysis (the last batching run had a manifest)
$out = Join-Path $WorkFolder 'fixtures_t3_strict'
& $analyser -InputFolder (Join-Path $WorkFolder 'fixtures_t3') -OutputFolder $out -ManifestOnly 6>$null 3>$null | Out-Null
$strict = [ordered]@{
    'unlisted file not analysed'  = -not (Import-Csv "$out\_ScriptAnalysis.csv" | Where-Object FullName -eq 'T3.Old.Symbol')
    'unlisted file still listed'  = [bool](Import-Csv "$out\_ScriptAnalysis_ExportCheck.csv" | Where-Object { $_.Issue -eq 'NotInManifest' -and $_.File -eq 'T3.Old.Symbol.txt' })
}
foreach ($k in $strict.Keys) {
    if ($strict[$k]) { Write-Host "PASS -ManifestOnly $k" } else { $failed++; Write-Host "FAIL -ManifestOnly $k" -ForegroundColor Red }
}

if ($failed) { Write-Host "$failed failure(s)." -ForegroundColor Red; exit 1 }
Write-Host 'All tests passed.' -ForegroundColor Green
exit 0
