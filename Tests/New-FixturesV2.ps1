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
    @{ Id = 3; Code = "strParent = Server.ThisObject.Parent.FullName`r`nSet x = Server.FindObject(strParent)`r`nx.Interface.Value(1)                    ' EXPECT:RC01" },
    @{ Id = 4; Code = "Set ie = CreateObject(""InternetExplorer.Application"")   ' EXPECT:RC08`r`nMsgBox ""ok""" })

New-FixtureV2 'T2.Code Only No Script' 'None' @('PickAction: 1 | L[Layer]/B[Button] | A | scriptCode | Current | (code) | (none)') '' @(
    @{ Id = 1; Code = "Stop                                                   ' EXPECT:RX04" })

New-FixtureV2 'T2.V8 Code' 'JavaScriptV8' @('PickAction: 1 | L[Layer]/B[Button] | A | scriptCode | Current | (code) | (none)') "function Local() { }" @(
    @{ Id = 1; Code = "MsgBox(""hi"");                                        // EXPECT:RX06" })

New-FixtureV2 'T2.Registry' 'VBScript' @() ("Sub Score()`r`n" +
"    OldScore = CLng(GetRegistry( ""SameGameScore"") )            ' EXPECT:RC16`r`n" +
"    OldScore = CLng(GetRegistry( ""SameGameScore"", 0) )         ' NOT:RC16`r`n" +
"    BestScore = CLng(App.GetRegistry(""SameGameBestScore""))     ' EXPECT:RC16`r`n" +
"    s = GetRegistry(""Name"")                                     ' EXPECT:RC16`r`n" +
"    t = GetRegistry(""Name"", """")                               ' NOT:RC16`r`n" +
"    ' x = CLng(GetRegistry(""InComment""))                         NOT:RC16`r`n" +
"    m = ""CLng(GetRegistry(""""InString""""))""                     ' NOT:RC16`r`n" +
"End Sub`r`n" +
"Sub ArticleFix()`r`n" +
"    a = GetRegistry(""newkey"")                                   ' NOT:RC16`r`n" +
"    If IsNull(a) Then`r`n" +
"        MsgBox """"`r`n" +
"    Else`r`n" +
"        MsgBox a`r`n" +
"    End If`r`n" +
"End Sub`r`n" +
"Sub ArticleBefore()`r`n" +
"    a = GetRegistry(""newkey"")                                   ' EXPECT:RC16`r`n" +
"    MsgBox a`r`n" +
"End Sub`r`n" +
"Sub NullTestElsewhere()`r`n" +
"    b = GetRegistry(""k"")                                        ' EXPECT:RC16`r`n" +
"    If IsNull(c) Then MsgBox ""other variable""`r`n" +
"End Sub") @()

# RC15: testing for an aggregate (article #15). Error with an Is Nothing test, Warning for any other Set.
New-FixtureV2 'T2.Aggregate' 'VBScript' @() ("Sub Before()`r`n" +
"    sPtFullName = ""Test.Point""`r`n" +
"    set pControl = Server.FindObject(sPtFullName)`r`n" +
"    Set y = pControl.Interface.Control                          ' EXPECT:RC15`r`n" +
"    If y Is Nothing Then`r`n" +
"        MsgBox ""Has No Control""`r`n" +
"    End If`r`n" +
"End Sub`r`n" +
"Sub After()`r`n" +
"    set pControl = Server.FindObject(""Test.Point"")`r`n" +
"    y = pControl.Interface.Control                              ' NOT:RC15`r`n" +
"    If not IsObject(y) Then MsgBox ""Has No Control""          ' NOT:RC09`r`n" +
"End Sub`r`n" +
"Sub Others()`r`n" +
"    Set h = Server.FindObject(""A"").Interface.Historic          ' EXPECT:RC15`r`n" +
"    Set g = Server.ThisObject.Interface.GISLocation            ' NOT:RC15`r`n" +
"    Set s = obj.Interface.Shared                               ' NOT:RC15`r`n" +
"    Set c = obj.Interface.Control.ControlMaximum               ' NOT:RC15`r`n" +
"    Set i = obj.Interface                                      ' NOT:RC15`r`n" +
"    x = 1 : Set k = obj.Interface.ScanParams                   ' EXPECT:RC15`r`n" +
"    ' Set z = obj.Interface.Control                              NOT:RC15`r`n" +
"End Sub") @()

New-FixtureV2 'T2.Registry JS' 'JavaScriptV8' @() ("function f() {`r`n" +
"    let n = parseInt(GetRegistry(""Count""));                    // EXPECT:RC16`r`n" +
"    let m = parseInt(GetRegistry(""Count"", ""0""));              // NOT:RC16`r`n" +
"}") @()
