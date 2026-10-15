param([string]$Out, [switch]$WithManifest)
# Template batching fixtures: a template mimic, two instance copies of it, and an overridden instance.
$utf8 = New-Object Text.UTF8Encoding($false)
function New-T3([string]$FullName, [string]$File, [hashtable]$Tmpl, [string]$Script) {
    $h = @("FormatVersion: 2", "FullName: $FullName", "Class: CMimic", "ObjectId: 3", "Language: VBScript",
           "InTemplate: $($Tmpl.InTemplate)", "ControlledByInstance: $($Tmpl.Controlled)", "DocumentContentOverride: $($Tmpl.Override)",
           "EditableSource: $($Tmpl.Source)", "EditInPlace: $($Tmpl.InPlace)")
    $body = $Script -replace "`r?`n", "`r`n"; $bl = ($body -split "`r`n").Count
    $start = 1 + $h.Count + 3 + 1
    $lines = @("' ===== GEOSCADA-SCRIPT-EXPORT BEGIN =====") + ($h | ForEach-Object { "' $_" }) +
             @("' ScriptStartLine: $start", "' ScriptEndLine: $($start + $bl - 1)", "' ===== GEOSCADA-SCRIPT-EXPORT END =====", $body)
    [IO.File]::WriteAllText((Join-Path $Out $File), ($lines -join "`r`n"), $utf8)
}
$script = "Sub S()`r`n    v = CLng(GetRegistry(""Key""))       ' EXPECT:RN21`r`nEnd Sub"
New-T3 'T3.Tmpl.Symbol' 'T3.Tmpl.Symbol.txt' @{ InTemplate = 'True'; Controlled = 'False'; Override = 'False'; Source = 'T3.Tmpl.Symbol'; InPlace = 'True' } $script
New-T3 'T3.Site1.Symbol' 'T3.Site1.Symbol.txt' @{ InTemplate = 'False'; Controlled = 'True'; Override = 'False'; Source = 'T3.Tmpl.Symbol'; InPlace = 'False' } $script
New-T3 'T3.Site2.Symbol' 'T3.Site2.Symbol.txt' @{ InTemplate = 'False'; Controlled = 'True'; Override = 'False'; Source = 'T3.Tmpl.Symbol'; InPlace = 'False' } $script
New-T3 'T3.Site3.Symbol' 'T3.Site3.Symbol.txt' @{ InTemplate = 'False'; Controlled = 'True'; Override = 'True'; Source = 'T3.Site3.Symbol'; InPlace = 'True' } $script

# Names starting with '-': the manifest holds them with a leading ' (spreadsheet formula guard), which the
# analyser must remove, or the instance is not matched to its file (MissingFile + NotInManifest).
$stop = "Sub S()`r`n    Stop`r`nEnd Sub"
New-T3 '-T3.Neg.Template' '-T3.Neg.Template.txt' @{ InTemplate = 'True'; Controlled = 'False'; Override = 'False'; Source = '-T3.Neg.Template'; InPlace = 'True' } $stop
New-T3 '-T3.Neg.Instance' '-T3.Neg.Instance.txt' @{ InTemplate = 'False'; Controlled = 'True'; Override = 'False'; Source = '-T3.Neg.Template'; InPlace = 'False' } $stop

# Export check cases (see the analyser help): a left-over file and a duplicate of it, a truncated file, a
# file whose ScriptStartLine is past its end, and a file the manifest lists under another name.
$plain = "Sub X()`r`nEnd Sub"
$standalone = @{ InTemplate = 'False'; Controlled = 'False'; Override = 'False'; Source = 'T3.Old.Symbol'; InPlace = 'True' }
New-T3 'T3.Old.Symbol' 'T3.Old.Symbol.txt' $standalone $plain
New-T3 'T3.Old.Symbol' 'T3.Old.Symbol [99].txt' $standalone $plain
# A script ending with a line break: the exporter counts the empty last line in ScriptEndLine. Not malformed.
New-T3 'T3.TrailingBreak.Symbol' 'T3.TrailingBreak.Symbol.txt' @{ InTemplate = 'False'; Controlled = 'False'; Override = 'False'; Source = 'T3.TrailingBreak.Symbol'; InPlace = 'True' } "Sub X()`r`nEnd Sub`r`n"
New-T3 'T3.Mismatch.Symbol' 'T3.Mismatch.Symbol.txt' @{ InTemplate = 'False'; Controlled = 'False'; Override = 'False'; Source = 'T3.Mismatch.Symbol'; InPlace = 'True' } $plain
[IO.File]::WriteAllText((Join-Path $Out 'T3.Trunc.Symbol.txt'), "' ===== GEOSCADA-SCRIPT-EXPORT BEGIN =====`r`n' FormatVersion: 2`r`n' FullName: T3.Trunc.Symbol`r`n' Class: CMimic", $utf8)
[IO.File]::WriteAllText((Join-Path $Out 'T3.Short.Symbol.txt'), "' ===== GEOSCADA-SCRIPT-EXPORT BEGIN =====`r`n' FormatVersion: 1`r`n' FullName: T3.Short.Symbol`r`n' Class: CMimic`r`n' ScriptStartLine: 40`r`n' ===== GEOSCADA-SCRIPT-EXPORT END =====`r`nSub X()`r`nEnd Sub", $utf8)

if ($WithManifest) {
    @(
        [pscustomobject]@{ FullName = 'T3.Tmpl.Symbol';  Class = 'CMimic'; Status = 'Exported'; File = 'T3.Tmpl.Symbol.txt';  EditableSource = 'T3.Tmpl.Symbol'; Message = '' }
        [pscustomobject]@{ FullName = 'T3.Site1.Symbol'; Class = 'CMimic'; Status = 'Exported'; File = 'T3.Site1.Symbol.txt'; EditableSource = 'T3.Tmpl.Symbol'; Message = '' }
        [pscustomobject]@{ FullName = 'T3.Site2.Symbol'; Class = 'CMimic'; Status = 'Exported'; File = 'T3.Site2.Symbol.txt'; EditableSource = 'T3.Tmpl.Symbol'; Message = '' }
        [pscustomobject]@{ FullName = 'T3.Site3.Symbol'; Class = 'CMimic'; Status = 'Exported'; File = 'T3.Site3.Symbol.txt'; EditableSource = 'T3.Site3.Symbol'; Message = '' }
        [pscustomobject]@{ FullName = "'-T3.Neg.Template"; Class = 'CMimic'; Status = 'Exported'; File = "'-T3.Neg.Template.txt"; EditableSource = "'-T3.Neg.Template"; Message = '' }
        [pscustomobject]@{ FullName = "'-T3.Neg.Instance"; Class = 'CMimic'; Status = 'Exported'; File = "'-T3.Neg.Instance.txt"; EditableSource = "'-T3.Neg.Template"; Message = '' }
        [pscustomobject]@{ FullName = 'T3.TrailingBreak.Symbol'; Class = 'CMimic'; Status = 'Exported'; File = 'T3.TrailingBreak.Symbol.txt'; EditableSource = 'T3.TrailingBreak.Symbol'; Message = '' }
        [pscustomobject]@{ FullName = 'T3.Renamed.Symbol'; Class = 'CMimic'; Status = 'Exported'; File = 'T3.Mismatch.Symbol.txt'; EditableSource = 'T3.Renamed.Symbol'; Message = '' }
        [pscustomobject]@{ FullName = 'T3.Trunc.Symbol'; Class = 'CMimic'; Status = 'Exported'; File = 'T3.Trunc.Symbol.txt'; EditableSource = 'T3.Trunc.Symbol'; Message = '' }
        [pscustomobject]@{ FullName = 'T3.Short.Symbol'; Class = 'CMimic'; Status = 'Exported'; File = 'T3.Short.Symbol.txt'; EditableSource = 'T3.Short.Symbol'; Message = '' }
        [pscustomobject]@{ FullName = 'T3.Gone.Symbol';  Class = 'CMimic'; Status = 'Exported'; File = 'T3.Gone.Symbol.txt';  EditableSource = 'T3.Gone.Symbol'; Message = '' }
        [pscustomobject]@{ FullName = 'T3.Broken.Symbol'; Class = 'CMimic'; Status = 'Error'; File = ''; EditableSource = ''; Message = 'Simulated read failure' }
    ) | Export-Csv -LiteralPath (Join-Path $Out '_ScriptExport_Manifest.csv') -NoTypeInformation -Encoding UTF8
}
