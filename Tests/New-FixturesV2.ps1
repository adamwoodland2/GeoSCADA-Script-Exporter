param([string]$Out)
# Format 2 fixtures: inline pick-action code sections, contexts, embedded mimics. Lines tagged
# EXPECT:<rule> / NOT:<rule> are checked by file line.
$utf8 = New-Object Text.UTF8Encoding($false)
function New-FixtureV2([string]$FullName, [string]$Language, [string[]]$Extra, [string]$Script, [object[]]$Sections) {
    $c = if ($Language -in 'VBScript', 'None') { "'" } else { '//' }
    $h = @("FormatVersion: 2", "FullName: $FullName", "Class: CMimic", "ObjectId: 2", "Language: $Language", "EditInPlace: True", "EditableSource: $FullName") + $Extra
    $body = if ($Script) { $Script -replace "`r?`n", "`r`n" } else { '' }
    $bodyLines = if ($body) { ($body -split "`r`n").Count } else { 0 }
    $hc = 1 + $h.Count + $Sections.Count + 3
    $start = $hc + 1; $end = $start + $bodyLines - 1; $next = $end + 1
    $codeEntries = @(); $secText = @()
    foreach ($s in $Sections) {
        $code = $s.Code -replace "`r?`n", "`r`n"; $cl = ($code -split "`r`n").Count
        $codeEntries += "PickActionCode: $($s.Id) | $($next + 1) | $($next + $cl)"
        $secText += "$c ===== GEOSCADA-SCRIPT-EXPORT PICKACTION $($s.Id) ====="; $secText += $code; $next += 1 + $cl
    }
    $lines = @("$c ===== GEOSCADA-SCRIPT-EXPORT BEGIN =====") + ($h | ForEach-Object { "$c $_" }) + ($codeEntries | ForEach-Object { "$c $_" }) +
             @("$c ScriptStartLine: $start", "$c ScriptEndLine: $end", "$c ===== GEOSCADA-SCRIPT-EXPORT END =====")
    if ($bodyLines) { $lines += $body }
    $lines += $secText
    [IO.File]::WriteAllText((Join-Path $Out "$FullName.txt"), ($lines -join "`r`n"), $utf8)
}

New-FixtureV2 'T2.Symbol' 'VBScript' @() "Sub OpenDeviceBook(site)`r`nEnd Sub" @()

New-FixtureV2 'T2.Parent' 'VBScript' @(
    'PickAction: 1 | L[Layer]/M_1[EmbeddedMimic] | A | scriptCall | Current | OpenDeviceBook | site <= ".FullName"',
    'PickAction: 2 | L[Layer]/M_2[EmbeddedMimic] | A | scriptCall | Current | NotAnywhere | (none)',
    'PickAction: 3 | L[Layer]/B[Button] | A | scriptTopLevel | TopLevel | (code) | (none)',
    'PickAction: 4 | L[Layer]/B2[Button] | A | scriptCode | Current | (code) | (none)',
    'EmbeddedMimic: L[Layer]/M_1[EmbeddedMimic] | T2.Symbol',
    'EmbeddedMimic: L[Layer]/M_2[EmbeddedMimic] | T2.Symbol') "Sub Local()`r`nEnd Sub" @(
    @{ Id = 3; Code = "strParent = Server.ThisObject.Parent.FullName`r`nSet x = Server.FindObject(strParent)`r`nx.Interface.Value(1)                    ' EXPECT:RN01" },
    @{ Id = 4; Code = "Set ie = CreateObject(""InternetExplorer.Application"")   ' EXPECT:RN09`r`nMsgBox ""ok""" })

New-FixtureV2 'T2.Code Only No Script' 'None' @('PickAction: 1 | L[Layer]/B[Button] | A | scriptCode | Current | (code) | (none)') '' @(
    @{ Id = 1; Code = "Stop                                                   ' EXPECT:RN20" })

New-FixtureV2 'T2.V8 Code' 'JavaScriptV8' @('PickAction: 1 | L[Layer]/B[Button] | A | scriptCode | Current | (code) | (none)') "function Local() { }" @(
    @{ Id = 1; Code = "MsgBox(""hi"");                                        // EXPECT:RN16" })

New-FixtureV2 'T2.Registry' 'VBScript' @() ("Sub Score()`r`n" +
"    OldScore = CLng(GetRegistry( ""SameGameScore"") )            ' EXPECT:RN21`r`n" +
"    OldScore = CLng(GetRegistry( ""SameGameScore"", 0) )         ' NOT:RN21`r`n" +
"    BestScore = CLng(App.GetRegistry(""SameGameBestScore""))     ' EXPECT:RN21`r`n" +
"    s = GetRegistry(""Name"")                                     ' EXPECT:RN21`r`n" +
"    t = GetRegistry(""Name"", """")                               ' NOT:RN21`r`n" +
"    ' x = CLng(GetRegistry(""InComment""))                         NOT:RN21`r`n" +
"    m = ""CLng(GetRegistry(""""InString""""))""                     ' NOT:RN21`r`n" +
"End Sub") @()

New-FixtureV2 'T2.Registry JS' 'JavaScriptV8' @() ("function f() {`r`n" +
"    let n = parseInt(GetRegistry(""Count""));                    // EXPECT:RN21`r`n" +
"    let m = parseInt(GetRegistry(""Count"", ""0""));              // NOT:RN21`r`n" +
"}") @()
