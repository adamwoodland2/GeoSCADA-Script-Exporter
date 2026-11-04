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

# Versions: every finding records the exporter / format / analyser version; the mixed fixture folder is flagged
$av = ([regex]::Match([IO.File]::ReadAllText($analyser), "\`$AnalyserVersion = '([^']+)'")).Groups[1].Value
$vok = [bool]($f | Where-Object { $_.FullName -eq 'T.VBS Mimic' -and $_.FormatVersion -eq '1' -and $_.AnalyserVersion -eq $av -and $_.ExporterVersion -eq '(unknown)' }) -and
       -not ($f | Where-Object { $_.AnalyserVersion -ne $av })
if ($vok) { Write-Host "PASS versions recorded on every finding (analyser $av)" } else { $failed++; Write-Host 'FAIL versions not recorded on the findings' -ForegroundColor Red }
if (Import-Csv "$out\_ScriptAnalysis_ExportCheck.csv" | Where-Object Issue -eq 'MixedExports') { Write-Host 'PASS mixed format 1 / 2 folder reported as MixedExports' }
else { $failed++; Write-Host 'FAIL mixed format 1 / 2 folder not reported as MixedExports' -ForegroundColor Red }

# RC12 is a header finding (on the PickAction line), so it is checked per button rather than by tag
$rc12 = @($f | Where-Object { $_.FullName -eq 'T.RC12 Mimic' -and $_.Rule -eq 'RC12' })
foreach ($b in @(@('B1', $true, 'second parameter has no value'), @('B2', $false, 'all parameters set'), @('B3', $true, 'no arguments at all'),
                 @('B4', $false, 'function has no parameters'), @('B5', $false, 'inline call with its own arguments'),
                 @('B6', $true, 'library function, ByVal parameter unset'), @('B7', $true, 'parameter named without a value'),
                 @('B8', $false, 'names differ only in case (VBScript)'))) {
    $hit = [bool]($rc12 | Where-Object { $_.Detail -like "*$($b[0])``[Button``] calls*" })
    if ($hit -eq $b[1]) { Write-Host "PASS RC12 $($b[0]): $($b[2])" } else { $failed++; Write-Host "FAIL RC12 $($b[0]): $($b[2]) (reported: $hit)" -ForegroundColor Red }
}

# RC13 severity: Warning when code follows the log on / log off in the same procedure, Info when it is the last statement
$rc13 = @($f | Where-Object { $_.Rule -eq 'RC13' })
foreach ($c in @(@('T.RC13 Mimic', 'Sub LogOnThenWork', 'Warning'), @('T.RC13 Mimic', 'Sub LogOffLast', 'Info'),
                 @('T.RC13 Mimic', 'Sub LogOnUserInIf', 'Info'), @('T.RC13 Mimic', 'Sub LogOffWithElse', 'Warning'),
                 @('T.RC13 Mimic', 'Sub OpenViewX', 'Warning'), @('T.RC13 Mimic', 'Sub Chained', 'Info'),
                 @('T.RC13 JS', 'function LogOnFirst', 'Warning'), @('T.RC13 JS', 'function LogOffLast', 'Info'),
                 @('T.RC13 JS', 'function Other', 'Info'), @('T.V8 Mimic', 'function Test', 'Warning'))) {
    $hit = @($rc13 | Where-Object { $_.FullName -eq $c[0] -and $_.Detail -like "* in $($c[1]) *" })
    $sev = ($hit | ForEach-Object Severity) -join ','
    if ($hit.Count -eq 1 -and $sev -eq $c[2]) { Write-Host "PASS RC13 $($c[0]) $($c[1]): $($c[2])" }
    else { $failed++; Write-Host "FAIL RC13 $($c[0]) $($c[1]): expected one $($c[2]), got '$sev'" -ForegroundColor Red }
}

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
    $f = @(Import-Csv "$out\_ScriptAnalysis.csv" | Where-Object Rule -eq 'RX10')
    $all = @(Import-Csv "$out\_ScriptAnalysis.csv")
    $a = @(Import-Csv "$out\_ScriptAnalysis_Actions.csv")
    $i = @(Import-Csv "$out\_ScriptAnalysis_Instances.csv")
    $c = @(Import-Csv "$out\_ScriptAnalysis_ExportCheck.csv")
    $issue = { param($Issue, $Name) [bool]($c | Where-Object { $_.Issue -eq $Issue -and ($_.FullName -eq $Name -or $_.File -eq $Name) }) }
    $rawActions = [IO.File]::ReadAllText("$out\_ScriptAnalysis_Actions.csv")
    $checks = [ordered]@{
        'two RX10 findings (template + overridden instance only)' = ($f.Count -eq 2)
        'template finding: Role Template, InstancesFixed 2'       = [bool]($f | Where-Object { $_.FullName -eq 'T3.Tmpl.Symbol' -and $_.Role -eq 'Template' -and $_.InstancesFixed -eq '2' })
        'overridden instance reported on its own'                 = [bool]($f | Where-Object { $_.FullName -eq 'T3.Site3.Symbol' -and $_.Role -eq 'OverriddenInstance' -and $_.InstancesFixed -eq '0' })
        'no findings against instance copies'                     = -not ($f | Where-Object { $_.FullName -in 'T3.Site1.Symbol', 'T3.Site2.Symbol' })
        'actions: template action mentions 2 copies'              = [bool]($a | Where-Object { $_.FullName -eq 'T3.Tmpl.Symbol' -and $_.Action -match '2 instance copies' })
        'actions: overridden instance says edit directly'         = [bool]($a | Where-Object { $_.FullName -eq 'T3.Site3.Symbol' -and $_.Action -match 'directly' })
        'instances CSV lists Site1 and Site2 under the template'  = ((@($i | Where-Object EditableSource -eq 'T3.Tmpl.Symbol' | ForEach-Object Instance) -join ',') -eq 'T3.Site1.Symbol,T3.Site2.Symbol')
        "'-' names: template finding counts its instance"          = [bool]($all | Where-Object { $_.FullName -eq "'-T3.Neg.Template" -and $_.Rule -eq 'RX04' -and $_.InstancesFixed -eq '1' })
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

# --- 4. 1.x rule IDs (RN01-RN28) ----------------------------------------------
# The map in the analyser ($OldRuleIds, read from its syntax tree) must cover RN01-RN28 once each, point at
# real rules, never reuse an ID, and agree with the README rules table and the CHANGELOG table.
$ast = [Management.Automation.Language.Parser]::ParseFile($analyser, [ref]$null, [ref]$null)
$hashOf = { param($Name)
    $a = $ast.Find({ param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq $Name }, $true)
    $a.Right.Find({ param($n) $n -is [Management.Automation.Language.HashtableAst] }, $true) }
$ruleKeys = @((& $hashOf '$Rules').KeyValuePairs | ForEach-Object { $_.Item1.Value })
$oldMap = [ordered]@{}
foreach ($p in (& $hashOf '$OldRuleIds').KeyValuePairs) {
    $oldMap[$p.Item1.Value] = $p.Item2.Find({ param($n) $n -is [Management.Automation.Language.StringConstantExpressionAst] }, $true).Value }
$docMap = { param($Text) $m = @{}; foreach ($x in [regex]::Matches($Text, '\| (\w+) \| (RN\d\d) \|')) { $m[$x.Groups[2].Value] = $x.Groups[1].Value }; $m }
$readme = & $docMap ([IO.File]::ReadAllText((Join-Path (Split-Path -Parent $here) 'README.md')))
$changelog = @{}; $cl = [IO.File]::ReadAllText((Join-Path (Split-Path -Parent $here) 'CHANGELOG.md'))
foreach ($x in [regex]::Matches($cl.Substring(0, $cl.IndexOf('### 1.4.0')), '\| (RN\d\d) \| (\w+) \|')) { $changelog[$x.Groups[1].Value] = $x.Groups[2].Value }
$want = @(1..28 | ForEach-Object { 'RN{0:D2}' -f $_ })
$old = [ordered]@{
    'map covers RN01-RN28'                   = ((@($oldMap.Keys) -join ',') -eq ($want -join ','))
    'map points at real rules, once each'    = (@($oldMap.Values | Where-Object { $ruleKeys -ccontains $_ } | Sort-Object -Unique).Count -eq 28)
    'no rule ID starts with RN'              = -not ($ruleKeys | Where-Object { $_ -like 'RN*' })
    'only RC14a / RC14b and GS have no 1.x ID' = ((@($ruleKeys | Where-Object { $oldMap.Values -notcontains $_ -and $_ -notlike 'GS*' }) -join ',') -eq 'RC14a,RC14b')
    'README 1.x rule column matches'         = ($readme.Count -eq 28 -and -not ($want | Where-Object { $readme[$_] -cne $oldMap[$_] }))
    'CHANGELOG 2.0.0 table matches'          = ($changelog.Count -eq 28 -and -not ($want | Where-Object { $changelog[$_] -cne $oldMap[$_] }))
}
# -IncludeRule / -ExcludeRule translate a 1.x ID (RN21 = RX10); an unknown ID still stops the run
$t3 = Join-Path $WorkFolder 'fixtures_t3'; $out = Join-Path $WorkFolder 'fixtures_t3_old'
& $analyser -InputFolder $t3 -OutputFolder $out -IncludeRule RN21 6>$null 3>$null | Out-Null
$incl = @(Import-Csv "$out\_ScriptAnalysis.csv")
$html = [IO.File]::ReadAllText("$out\_ScriptAnalysis.html")
& $analyser -InputFolder $t3 -OutputFolder $out -ExcludeRule RN21 6>$null 3>$null | Out-Null
$excl = @(Import-Csv "$out\_ScriptAnalysis.csv")
$unknown = $false
try { & $analyser -InputFolder $t3 -OutputFolder $out -IncludeRule RN29 6>$null 3>$null | Out-Null } catch { $unknown = $_.Exception.Message -like "Unknown rule 'RN29'*" }
$old['-IncludeRule RN21 runs RX10 only']        = ($incl.Count -gt 0 -and -not ($incl | Where-Object Rule -ne 'RX10'))
$old['-ExcludeRule RN21 leaves RX10 out']       = ($excl.Count -gt 0 -and -not ($excl | Where-Object Rule -eq 'RX10'))
$old['HTML report shows the 1.x ID (RX10 was RN21)'] = $html.Contains('"o":"RN21"')
$old['unknown rule ID RN29 is an error']        = $unknown
foreach ($k in $old.Keys) {
    if ($old[$k]) { Write-Host "PASS 1.x rule IDs: $k" } else { $failed++; Write-Host "FAIL 1.x rule IDs: $k" -ForegroundColor Red }
}

if ($failed) { Write-Host "$failed failure(s)." -ForegroundColor Red; exit 1 }
Write-Host 'All tests passed.' -ForegroundColor Green
exit 0
