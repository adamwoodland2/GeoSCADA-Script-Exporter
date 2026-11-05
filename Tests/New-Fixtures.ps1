param([string]$Out)
# Builds synthetic exporter files to exercise every analyser rule. Lines tagged EXPECT:<rule> must be
# reported on that script line; lines tagged NOT:<rule> must not be.
if (Test-Path $Out) { Remove-Item $Out -Recurse -Force }
New-Item -ItemType Directory $Out | Out-Null
$utf8 = New-Object Text.UTF8Encoding($false)
function New-Fixture([string]$FullName, [string]$Class, [string]$Language, [string[]]$Extra, [string]$Script) {
    $c = if ($Language -eq 'VBScript' -or $Language -eq 'None') { "'" } else { '//' }
    $guid = @{ VBScript = '{B54F3741-5B07-11CF-A4B0-00AA004A55E8}'; JScript = '{F414C260-6AC0-11CF-B6D1-00AA00BBBB58}'; JavaScriptV8 = '{16BA25BB-FB03-48BA-AECA-C53BF6ABD2A1}'; None = '' }[$Language]
    $h = @("FormatVersion: 1", "FullName: $FullName", "Class: $Class", "ObjectId: 1", "Language: $Language", "LanguageGuid: $guid", "EditInPlace: True", "EditableSource: $FullName") + $Extra
    $lines = @("$c ===== GEOSCADA-SCRIPT-EXPORT BEGIN =====") + ($h | ForEach-Object { "$c $_" })
    $lines += "$c ScriptStartLine: $($lines.Count + 3)"
    $lines += "$c ===== GEOSCADA-SCRIPT-EXPORT END ====="
    $text = ($lines -join "`r`n") + "`r`n" + ($Script -replace "`r?`n", "`r`n")
    [IO.File]::WriteAllText((Join-Path $Out "$FullName.txt"), $text, $utf8)
}

New-Fixture 'T.VBS Mimic' 'CMimic' 'VBScript' @('Library: T.VBS Lib', 'Library: T.JS Lib', 'LibraryMissing: T.Gone Lib',
    'PickAction: L[Layer]/B[Button] | A | scriptCall | DoThing | (none)',
    'PickAction: L[Layer]/B2[Button] | A | scriptCall | LibThing | (none)',
    'PickAction: L[Layer]/B3[Button] | A | scriptCall | NoSuchSub | (none)',
    'PickAction: L[Layer]/B4[Button] | A | scriptCall | DoThing 1, "x" | (none)',
    'PickAction: L[Layer]/B5[Button] | A | scriptCall | strX = Server.ThisObject.FullName MsgBox strX | (none)',
    'PickAction: L[Layer]/B6[Button] | A | scriptCall | MissingInline 1, "x" | (none)') @'
Sub DoThing()
    Set x = Server.FindObject("Example Projects.New Float")
    x.Interface.Value(3.14)                                   ' EXPECT:RC01
    x.Interface.Value = 3.14                                  ' NOT:RC01
    Set q = Form.AddQuery(2, 5, 50, 20, "SELECT Name FROM CMimic")  ' EXPECT:RC07 NOT:RC06b
    v = q.Value(1, 2)                                         ' NOT:RC01
    r = Menu.AddItem("Open")                                  ' EXPECT:RC02
    Menu.AddItem "Close"                                      ' NOT:RC02
    Set subm = Menu.AddMenu("More")                           ' NOT:RC02
    Form.Init("Test")
    Set field = Form.AddEditBox(2, 2, 10, 1)
    field.Value = "test"                                      ' NOT:RC03
    MsgBox field                                              ' EXPECT:RC03
    MsgBox field.Value                                        ' NOT:RC03
    s = "Name: " & field                                      ' EXPECT:RC03
    If a Then
    ELSEIF field Mod Bound = 0 THEN                  ' EXPECT:RC03
    End If
    n = field * 2                                             ' EXPECT:RC03
    field = "new"                                             ' EXPECT:RC03
    If field Is Nothing Then b = 1                            ' NOT:RC03
    Set other = field                                         ' NOT:RC03
    LogControl field                                          ' NOT:RC03
    r = Validate(1, field)                                    ' NOT:RC03
    If IsObject(field) Then b = 1                             ' NOT:RC03
    r = CStr(field)                                           ' EXPECT:RC03
    Set lb = Form.AddListBox(2, 5, 20, 1)                     ' EXPECT:RC06a EXPECT:RC06b
    msg = "Form.AddListBox(1, 1, 5, 1) in a string"           ' NOT:RC06b
    Set ie = CreateObject("InternetExplorer.Application")     ' EXPECT:RC08
    Set fso = CreateObject("Scripting.FileSystemObject")      ' NOT:RX02
    Set bad = CreateObject("No.Such.ProgId.Anywhere")         ' EXPECT:RX02
    Set db = CreateObject("DAO.DBEngine.36")                 ' EXPECT:RX01
    Set dyn = CreateObject(progName)                          ' EXPECT:RX03
    Set pt = objServer.CreateObject(pointClass, "Grp.Points") ' NOT:RX03
    Set pt2 = objServer.CreateObject("CPointAlgManual", "G")  ' NOT:RX02
    Set pt3 = Server.RootObject . CreateObject(cls, "G")      ' NOT:RX03
    ' Set y = CreateObject("No.Such.In.Comment")              NOT:RX02
    msg = "x.Interface.Value(1) inside a string"              ' NOT:RC01
    y = (x.Interface.Value(0))                                ' NOT:RC01
    MsgBox x.Interface.Value(0)                               ' NOT:RC01
    If x.Interface.Value(0) > 1 Then b = 2                    ' NOT:RC01
    If a Then x.Interface.Value(2)                            ' EXPECT:RC01
    Call x.Interface.Value(2)                                 ' EXPECT:RC01
    Server.FindObject("A").Interface.Value(2) : b = 1         ' EXPECT:RC01
    With x.Interface
        .Value(2)                                             ' EXPECT:RC01
    End With
    toString = "Site 25"                                      ' EXPECT:RC10
    s = "x to " + toString                                    ' EXPECT:RC10
    s = d.ToString                                            ' NOT:RC10
    t = "toString in a string"                                ' NOT:RC10
    b = isEqual(1, 2)                                         ' NOT:RC10
    If a Then b = 1
    End If
    Menu.Init                                                 ' NOT:RC02
    msg = "Before: CreateObject(""No.Such.ProgId"")"             ' NOT:RX02
    Stop                                                      ' EXPECT:RX04
End Sub
'@

New-Fixture 'T.VBS Lib' 'CScriptLibrary' 'VBScript' @('ReferencedByMimics: 1') @'
Public Function LibThing()
End Function
'@
New-Fixture 'T.Unused Lib' 'CScriptLibrary' 'VBScript' @('ReferencedByMimics: 0') @'
Sub Nothing()
End Sub
'@
New-Fixture 'T.JS Lib' 'CScriptLibrary' 'JScript' @('ReferencedByMimics: 1') @'
function jsLibThing() { }
'@

New-Fixture 'T.JS Mimic' 'CMimic' 'JScript' @('Library: T.JS Lib',
    'PickAction: L[Layer]/B[Button] | A | scriptCall | ConfigureQueries | site <= ".FullName"',
    'PickAction: L[Layer]/B2[Button] | A | scriptCall | jslibthing | (none)') @'
function ConfigureQueries(site) {
    var o = Server.LookUpObject(12);                         // EXPECT:JS02
    form.Init("Filter");                                     // EXPECT:JS02
    Form.init("Filter");                                     // EXPECT:JS02
    var box = Form.AddEditBox(5, 3, 60, 1);
    box.value = "x";                                         // EXPECT:JS02
    box.Value = "x";                                         // NOT:JS02
    var chk = Form.AddCheckBox(1, 1, "Opt");
    if (chk.checked) { }                                     // EXPECT:JS02
    var ok = Form.AddPushButton(5, 5, "OK");
    ok.Default = 1;                                          // EXPECT:JS03
    ok.Default = true;                                       // NOT:JS03
    var sql = Mimic.Layers("Content").Item("List").sql;      // EXPECT:JS02
    var vbe = new ActiveXObject("ScriptControl");            // EXPECT:JS01
    var pt = srv.CreateObject(cls, "Grp.Points");            // NOT:RX03
    var lst = Form.AddListBox(1, 1, 5, 1);                   // EXPECT:RC06b
    var dyn = new ActiveXObject(progName);                   // EXPECT:RX03
    var data = { value: 1 }; data.value = 2;                 // NOT:JS02
    _.prototype.value = function () { return 1; };           // NOT:JS02
    var arr = [3, 1]; arr.sort();                            // NOT:JS02
    var toString = arr.toString();                           // NOT:RC10
    var t = getRegistry("x");                                // EXPECT:JS02
    // box.value in a comment                                   NOT:JS02
    var s = "box.value in a string";                         // NOT:JS02
    debugger;                                                // EXPECT:RX04
}
function RegexCases(s, total) {
    if (s) return /Form.init/.test(s);                       // NOT:JS02
    if (typeof /box.value/ == "object") { }                  // NOT:JS02
    return /a\/\//.test(s) && Form.init("x");                // EXPECT:JS02
}
function DivisionCase(total) {
    var h = total / 2; Form.init("y");                       // EXPECT:JS02
    var w = (total) / 2 / 3; form.Init("z");                 // EXPECT:JS02
}
'@

New-Fixture 'T.V8 Mimic' 'CMimic' 'JavaScriptV8' @() @'
function Test() {
    MsgBox("Hello");                                         // EXPECT:RX06
    Window.Alert("Hello");                                   // NOT:RX06
    let d = LocalTimeToUTC(new Date());                      // EXPECT:RX07
    let x = new ActiveXObject("Scripting.FileSystemObject"); // EXPECT:RX05
    Form.Init("Pixels");                                     // EXPECT:RX08
    App.LogOnUser("SYS", "user", "pw");                      // EXPECT:RC13
    let re = /it's "quoted"/;                                // NOT:JS02
    let p = Server.FindObject("A").Interface.Value(5);       // NOT:RC01
    Server.FindObject("A").Interface.Value(5);               // EXPECT:RC01
    if (p) Server.FindObject("A").Interface.Value(5);        // EXPECT:RC01
}
'@

New-Fixture 'T.VBS Declared' 'CScriptLibrary' 'VBScript' @('ReferencedByMimics: 1') @'
Dim GetType                                                   ' EXPECT:RC10
Sub A(ByVal equals)                                           ' EXPECT:RC10
    Dim x, arr(2, 3), toString                                ' EXPECT:RC10
    toString = "Hello"                                        ' NOT:RC10
    MsgBox toString                                           ' NOT:RC10
    x = equals                                                ' NOT:RC10
End Sub
Sub B()
    toString = "x"                                            ' EXPECT:RC10
    y = GetType                                               ' NOT:RC10
    GetHashCode = 5                                           ' EXPECT:RC10
End Sub
'@

# RC12: every parameter of a pick action's function needs a value (checked in Run-Tests.ps1 by button)
New-Fixture 'T.RC12 Mimic' 'CMimic' 'VBScript' @('Library: T.RC12 Lib',
    "PickAction: L[Layer]/B1[Button] | A | scriptCall | TwoArgs | One <= 'hello'",
    "PickAction: L[Layer]/B2[Button] | A | scriptCall | TwoArgs | One <= 'hello', Two <= 'goodbye'",
    'PickAction: L[Layer]/B3[Button] | A | scriptCall | TwoArgs | (none)',
    'PickAction: L[Layer]/B4[Button] | A | scriptCall | NoArgs | (none)',
    'PickAction: L[Layer]/B5[Button] | A | scriptCall | TwoArgs 1, 2 | (none)',
    'PickAction: L[Layer]/B6[Button] | A | scriptCall | LibTwo | x <= 1',
    'PickAction: L[Layer]/B7[Button] | A | scriptCall | TwoArgs | One <= 1, Two',
    "PickAction: L[Layer]/B8[Button] | A | scriptCall | twoargs | two <= 'b', ONE <= 'a'",
    "PickAction: L[Layer]/B9[Button] | A | scriptCall | ArrayArgs | vals <= 1, n <= 2",
    "PickAction: L[Layer]/B10[Button] | A | scriptCall | ArrayArgs | vals <= 1") @'
Option Explicit
Public Sub TwoArgs(One, Two)
    MsgBox "One"
End Sub
Sub NoArgs()
End Sub
Sub ArrayArgs(vals(), n)
End Sub
'@
New-Fixture 'T.RC12 Lib' 'CScriptLibrary' 'VBScript' @('ReferencedByMimics: 1') @'
Function LibTwo(x, ByVal y)
End Function
'@

# RC13: log on / log off stops the script. Warning / Info per procedure is checked in Run-Tests.ps1.
New-Fixture 'T.RC13 Mimic' 'CMimic' 'VBScript' @() @'
Sub LogOnThenWork()
    App.LogOn                                                 ' EXPECT:RC13
    MsgBox "Logged on"
End Sub
Sub LogOffLast()
    MsgBox "Bye"
    App.LogOff                                                ' EXPECT:RC13
End Sub
Sub LogOnUserInIf(ok)
    If ok Then
        App.LogOnUser "SYS", "user", "pw"                     ' EXPECT:RC13
    End If
End Sub
Sub LogOffWithElse(ok)
    If ok Then app.logoff Else MsgBox "Stay"                  ' EXPECT:RC13
End Sub
Sub Mentions()
    s = "App.LogOn in a string"                               ' NOT:RC13
    ' App.LogOff in a comment                                   NOT:RC13
    x = MyApp.LogOn                                           ' NOT:RC13
    Set u = Server.FindObject("A.B")
    u.LogOn                                                   ' NOT:RC13
End Sub
Public Sub OpenViewX
    Set vx = CreateObject("Serck.ViewXApplication")
    vx.Logon "MyServer", "MyUser", "MyPassword"               ' EXPECT:RC13
    MsgBox "Logon OK"
End Sub
Sub Chained()
    CreateObject("Serck.ViewXApplication").Logon "S", "U", "P"   ' EXPECT:RC13
End Sub
'@
New-Fixture 'T.RC13 JS' 'CMimic' 'JScript' @() @'
function LogOnFirst() {
    App.LogOn();                                             // EXPECT:RC13
    Form.Init("After");
}
function LogOffLast(ok) {
    if (ok) {
        App.LogOff();                                        // EXPECT:RC13
    }
}
function Other() {
    var s = "App.LogOn()";                                   // NOT:RC13
    var app = GetSession(); app.logon();                     // NOT:RC13
    var vx = new ActiveXObject("Serck.ViewXApplication");
    vx.Logon("S", "U", "P");                                 // EXPECT:RC13
}
function LogonTypeLib() {
    App.Logon();                                             // EXPECT:RC13 NOT:JS02
}
'@

# RC14: GetOPCValue cannot read aggregates. RC14a = the tag names a known aggregate; RC14b = every other call.
New-Fixture 'T.RC14 Mimic' 'CMimic' 'VBScript' @() @'
Sub Reads(myPoint, tag)
    y = Server.GetOPCValue(myPoint & ".Control")                      ' EXPECT:RC14a NOT:RC14b
    z = Server.GetOPCValue(myPoint & ".Control.ControlMaximum")       ' EXPECT:RC14a
    h = server.getopcvalue("Plant.Pump 1.Historic.TrendInterval")     ' EXPECT:RC14a
    v = Server.GetOPCValue(myPoint & ".CurrentValue")                 ' EXPECT:RC14b NOT:RC14a
    t = Server.GetOPCValue(tag)                                       ' EXPECT:RC14b NOT:RC14a
    s = Server.GetOPCValue("Plant.Site.Shared")                       ' EXPECT:RC14b NOT:RC14a
    s2 = Server.GetOPCValue("Plant.Site.Shared.Something")            ' EXPECT:RC14a
    w = Server.GetOPCValue(myPoint & ".ControlValue")                 ' EXPECT:RC14b NOT:RC14a
    msg = "Server.GetOPCValue(p & .Control)"                          ' NOT:RC14a NOT:RC14b
    ' Server.GetOPCValue(myPoint & ".Control")                          NOT:RC14a NOT:RC14b
    Server.SetOPCValue myPoint & ".Control.ControlMaximum", 5         ' NOT:RC14a NOT:RC14b
    q = objServer.GetOPCValue(myPoint & ".Control")                   ' NOT:RC14a NOT:RC14b
End Sub
'@
New-Fixture 'T.RC14 JS' 'CMimic' 'JavaScriptV8' @() @'
function Reads(myPoint) {
    let z = Server.GetOPCValue(myPoint + ".Control.ControlMaximum"); // EXPECT:RC14a
    let g = Server.GetOPCValue(`${myPoint}.GISLocation.Latitude`);   // EXPECT:RC14a
    let v = Server.GetOPCValue(myPoint + '.CurrentValue');           // EXPECT:RC14b NOT:RC14a
    let c = Server.getOPCValue(myPoint + ".Control");                // NOT:RC14a NOT:RC14b
}
'@

# RC04-RC11 (Knowledge Base "Adjustments for VBScript in Geo SCADA 2025 R2")
New-Fixture 'T.KB Mimic' 'CMimic' 'VBScript' @() @'
Sub Objects(rowId)
    Set obj = Server.LookupObject(rowId)
    MsgBox CStr(obj)                                          ' EXPECT:RC04
    s = "Id: " & obj                                          ' EXPECT:RC04
    MsgBox CStr(obj.Id)                                       ' NOT:RC04
    If obj Is Nothing Then MsgBox "Not found"                 ' NOT:RC04
    If Not obj Is Nothing Then b = 1                          ' NOT:RC04
    ShowObject obj                                            ' NOT:RC04
    Set other = obj                                           ' NOT:RC04
    If IsNull(obj) Then MsgBox "Not found"                    ' EXPECT:RC09
    If IsNull(obj) Then b = 2                                 ' NOT:RC04
    If obj = Null Then b = 3                                  ' EXPECT:RC09
    If IsObject(obj) Then b = 4                               ' EXPECT:RC09
    If IsEmpty(obj) Then b = 5                                ' EXPECT:RC09
    Set me1 = Server.ThisObject
    MsgBox me1                                                ' EXPECT:RC04
    If IsNull(me1) Then b = 6                                 ' NOT:RC09
    Set itf = Server.FindObject("A.B").Interface
    MsgBox itf                                                ' NOT:RC04
    Set par = Server.FindObject("A.B").Parent
    MsgBox par                                                ' EXPECT:RC04
End Sub
Sub OtherProc()
    MsgBox obj                                                ' NOT:RC04
    If IsNull(obj) Then b = 1                                 ' NOT:RC09
End Sub
Sub Dates()
    Form.Init "Dates"
    Set StartTime = Form.AddDateTime(2, 2)
    Form.Show
    SetRegistry "start.date", StartTime.Value                 ' EXPECT:RC11
    SetRegistry "start.date", CLng(Int(CDbl(StartTime.Value)))   ' NOT:RC11
    SetRegistry "now", Now                                    ' EXPECT:RC11
    SetRegistry("tomorrow", DateAdd("d", 1, Now))             ' EXPECT:RC11
    App.SetRegistry "yesterday", Now - 1                      ' EXPECT:RC11
    SetRegistry "Start Time", "TIMESTAMP '" & sql & "'"       ' NOT:RC11
    SetRegistry "EndTime", Timer*10                           ' NOT:RC11
    SetRegistry "text", "Date: " & Now                        ' NOT:RC11
    t0 = Now
    SetRegistry "t0", t0                                      ' EXPECT:RC11
    n = 5
    SetRegistry "n", n                                        ' NOT:RC11
    StartTime.Value = GetRegistry("start.date", 0)            ' EXPECT:RC11
    StartTime.Value = CDate(GetRegistry("start.date", 0))     ' NOT:RC11
    MsgBox FormatDateTime(GetRegistry("start.date", 0))       ' EXPECT:RC11
    MsgBox FormatDateTime(CDate(GetRegistry("start.date", 0)))   ' NOT:RC11
    d = DateDiff("d", Now, GetRegistry("start.date", 0))      ' EXPECT:RC11
End Sub
Sub NoDate()
    SetRegistry "t0", t0                                      ' NOT:RC11
End Sub
'@
# A library that defines its own isNull (as some real systems do): its callers are not RC09.
New-Fixture 'T.KB Lib' 'CScriptLibrary' 'VBScript' @('ReferencedByMimics: 1') @'
Function isNull(o)
    isNull = o Is Nothing
End Function
'@
New-Fixture 'T.KB Own IsNull' 'CMimic' 'VBScript' @('Library: T.KB Lib') @'
Sub Check()
    Set o = Server.FindObject("A.B")
    If isNull(o) Then MsgBox "Not found"                      ' NOT:RC09
End Sub
'@

New-Fixture 'T.No Script Mimic' 'CMimic' 'None' @('PickAction: L[Layer]/Lst[List] | A | scriptCall | DisplayCustomMenu | (none)') ''

New-Fixture 'T.JS Labelled VBS' 'CMimic' 'JScript' @() @'
Dim a
Set q = Form.AddQuery(5, 11, 100, 20, sSQL)
If Form.Selection = "Apply" Then
    a = 1
End If
'@

New-Fixture 'T.TopLevel No Script' 'CMimic' 'None' @('PickAction: L[Layer]/B[Button] | A | scriptTopLevel | MsgBox "x" | (none)', 'PickAction: L[Layer]/B2[Button] | A | scriptCode | DoIt 1 | (none)') ''
