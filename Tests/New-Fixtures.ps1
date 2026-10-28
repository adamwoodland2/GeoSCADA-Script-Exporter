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
    x.Interface.Value(3.14)                                   ' EXPECT:RN01
    x.Interface.Value = 3.14                                  ' NOT:RN01
    Set q = Form.AddQuery(2, 5, 50, 20, "SELECT Name FROM CMimic")  ' EXPECT:RN05 NOT:RN27
    v = q.Value(1, 2)                                         ' NOT:RN01
    r = Menu.AddItem("Open")                                  ' EXPECT:RN02
    Menu.AddItem "Close"                                      ' NOT:RN02
    Set subm = Menu.AddMenu("More")                           ' NOT:RN02
    Form.Init("Test")
    Set field = Form.AddEditBox(2, 2, 10, 1)
    field.Value = "test"                                      ' NOT:RN03
    MsgBox field                                              ' EXPECT:RN03
    MsgBox field.Value                                        ' NOT:RN03
    s = "Name: " & field                                      ' EXPECT:RN03
    If a Then
    ELSEIF field Mod Bound = 0 THEN                  ' EXPECT:RN03
    End If
    n = field * 2                                             ' EXPECT:RN03
    field = "new"                                             ' EXPECT:RN03
    If field Is Nothing Then b = 1                            ' NOT:RN03
    Set other = field                                         ' NOT:RN03
    LogControl field                                          ' NOT:RN03
    r = Validate(1, field)                                    ' NOT:RN03
    If IsObject(field) Then b = 1                             ' NOT:RN03
    r = CStr(field)                                           ' EXPECT:RN03
    Set lb = Form.AddListBox(2, 5, 20, 1)                     ' EXPECT:RN04 EXPECT:RN27
    msg = "Form.AddListBox(1, 1, 5, 1) in a string"           ' NOT:RN27
    Set ie = CreateObject("InternetExplorer.Application")     ' EXPECT:RN09
    Set fso = CreateObject("Scripting.FileSystemObject")      ' NOT:RN07
    Set bad = CreateObject("No.Such.ProgId.Anywhere")         ' EXPECT:RN07
    Set db = CreateObject("DAO.DBEngine.36")                 ' EXPECT:RN06
    Set dyn = CreateObject(progName)                          ' EXPECT:RN08
    Set pt = objServer.CreateObject(pointClass, "Grp.Points") ' NOT:RN08
    Set pt2 = objServer.CreateObject("CPointAlgManual", "G")  ' NOT:RN07
    Set pt3 = Server.RootObject . CreateObject(cls, "G")      ' NOT:RN08
    ' Set y = CreateObject("No.Such.In.Comment")              NOT:RN07
    msg = "x.Interface.Value(1) inside a string"              ' NOT:RN01
    y = (x.Interface.Value(0))                                ' NOT:RN01
    MsgBox x.Interface.Value(0)                               ' NOT:RN01
    If x.Interface.Value(0) > 1 Then b = 2                    ' NOT:RN01
    If a Then x.Interface.Value(2)                            ' EXPECT:RN01
    Call x.Interface.Value(2)                                 ' EXPECT:RN01
    Server.FindObject("A").Interface.Value(2) : b = 1         ' EXPECT:RN01
    With x.Interface
        .Value(2)                                             ' EXPECT:RN01
    End With
    toString = "Site 25"                                      ' EXPECT:RN22
    s = "x to " + toString                                    ' EXPECT:RN22
    s = d.ToString                                            ' NOT:RN22
    t = "toString in a string"                                ' NOT:RN22
    b = isEqual(1, 2)                                         ' NOT:RN22
    If a Then b = 1
    End If
    Menu.Init                                                 ' NOT:RN02
    msg = "Before: CreateObject(""No.Such.ProgId"")"             ' NOT:RN07
    Stop                                                      ' EXPECT:RN20
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
    var o = Server.LookUpObject(12);                         // EXPECT:RN11
    form.Init("Filter");                                     // EXPECT:RN11
    Form.init("Filter");                                     // EXPECT:RN11
    var box = Form.AddEditBox(5, 3, 60, 1);
    box.value = "x";                                         // EXPECT:RN11
    box.Value = "x";                                         // NOT:RN11
    var chk = Form.AddCheckBox(1, 1, "Opt");
    if (chk.checked) { }                                     // EXPECT:RN11
    var ok = Form.AddPushButton(5, 5, "OK");
    ok.Default = 1;                                          // EXPECT:RN12
    ok.Default = true;                                       // NOT:RN12
    var sql = Mimic.Layers("Content").Item("List").sql;      // EXPECT:RN11
    var vbe = new ActiveXObject("ScriptControl");            // EXPECT:RN10
    var pt = srv.CreateObject(cls, "Grp.Points");            // NOT:RN08
    var lst = Form.AddListBox(1, 1, 5, 1);                   // EXPECT:RN27
    var dyn = new ActiveXObject(progName);                   // EXPECT:RN08
    var data = { value: 1 }; data.value = 2;                 // NOT:RN11
    _.prototype.value = function () { return 1; };           // NOT:RN11
    var arr = [3, 1]; arr.sort();                            // NOT:RN11
    var toString = arr.toString();                           // NOT:RN22
    var t = getRegistry("x");                                // EXPECT:RN11
    // box.value in a comment                                   NOT:RN11
    var s = "box.value in a string";                         // NOT:RN11
    debugger;                                                // EXPECT:RN20
}
function RegexCases(s, total) {
    if (s) return /Form.init/.test(s);                       // NOT:RN11
    if (typeof /box.value/ == "object") { }                  // NOT:RN11
    return /a\/\//.test(s) && Form.init("x");                // EXPECT:RN11
}
function DivisionCase(total) {
    var h = total / 2; Form.init("y");                       // EXPECT:RN11
    var w = (total) / 2 / 3; form.Init("z");                 // EXPECT:RN11
}
'@

New-Fixture 'T.V8 Mimic' 'CMimic' 'JavaScriptV8' @() @'
function Test() {
    MsgBox("Hello");                                         // EXPECT:RN16
    Window.Alert("Hello");                                   // NOT:RN16
    let d = LocalTimeToUTC(new Date());                      // EXPECT:RN17
    let x = new ActiveXObject("Scripting.FileSystemObject"); // EXPECT:RN15
    Form.Init("Pixels");                                     // EXPECT:RN18
    let re = /it's "quoted"/;                                // NOT:RN11
    let p = Server.FindObject("A").Interface.Value(5);       // NOT:RN01
    Server.FindObject("A").Interface.Value(5);               // EXPECT:RN01
    if (p) Server.FindObject("A").Interface.Value(5);        // EXPECT:RN01
}
'@

New-Fixture 'T.VBS Declared' 'CScriptLibrary' 'VBScript' @('ReferencedByMimics: 1') @'
Dim GetType                                                   ' EXPECT:RN22
Sub A(ByVal equals)                                           ' EXPECT:RN22
    Dim x, arr(2, 3), toString                                ' EXPECT:RN22
    toString = "Hello"                                        ' NOT:RN22
    MsgBox toString                                           ' NOT:RN22
    x = equals                                                ' NOT:RN22
End Sub
Sub B()
    toString = "x"                                            ' EXPECT:RN22
    y = GetType                                               ' NOT:RN22
    GetHashCode = 5                                           ' EXPECT:RN22
End Sub
'@

# RN26: every parameter of a pick action's function needs a value (checked in Run-Tests.ps1 by button)
New-Fixture 'T.RN26 Mimic' 'CMimic' 'VBScript' @('Library: T.RN26 Lib',
    "PickAction: L[Layer]/B1[Button] | A | scriptCall | TwoArgs | One <= 'hello'",
    "PickAction: L[Layer]/B2[Button] | A | scriptCall | TwoArgs | One <= 'hello', Two <= 'goodbye'",
    'PickAction: L[Layer]/B3[Button] | A | scriptCall | TwoArgs | (none)',
    'PickAction: L[Layer]/B4[Button] | A | scriptCall | NoArgs | (none)',
    'PickAction: L[Layer]/B5[Button] | A | scriptCall | TwoArgs 1, 2 | (none)',
    'PickAction: L[Layer]/B6[Button] | A | scriptCall | LibTwo | x <= 1',
    'PickAction: L[Layer]/B7[Button] | A | scriptCall | TwoArgs | One <= 1, Two',
    "PickAction: L[Layer]/B8[Button] | A | scriptCall | twoargs | two <= 'b', ONE <= 'a'") @'
Option Explicit
Public Sub TwoArgs(One, Two)
    MsgBox "One"
End Sub
Sub NoArgs()
End Sub
'@
New-Fixture 'T.RN26 Lib' 'CScriptLibrary' 'VBScript' @('ReferencedByMimics: 1') @'
Function LibTwo(x, ByVal y)
End Function
'@

# RN23-RN25 (Knowledge Base "Adjustments for VBScript in Geo SCADA 2025 R2")
New-Fixture 'T.KB Mimic' 'CMimic' 'VBScript' @() @'
Sub Objects(rowId)
    Set obj = Server.LookupObject(rowId)
    MsgBox CStr(obj)                                          ' EXPECT:RN23
    s = "Id: " & obj                                          ' EXPECT:RN23
    MsgBox CStr(obj.Id)                                       ' NOT:RN23
    If obj Is Nothing Then MsgBox "Not found"                 ' NOT:RN23
    If Not obj Is Nothing Then b = 1                          ' NOT:RN23
    ShowObject obj                                            ' NOT:RN23
    Set other = obj                                           ' NOT:RN23
    If IsNull(obj) Then MsgBox "Not found"                    ' EXPECT:RN24
    If IsNull(obj) Then b = 2                                 ' NOT:RN23
    If obj = Null Then b = 3                                  ' EXPECT:RN24
    If IsObject(obj) Then b = 4                               ' EXPECT:RN24
    If IsEmpty(obj) Then b = 5                                ' EXPECT:RN24
    Set me1 = Server.ThisObject
    MsgBox me1                                                ' EXPECT:RN23
    If IsNull(me1) Then b = 6                                 ' NOT:RN24
    Set itf = Server.FindObject("A.B").Interface
    MsgBox itf                                                ' NOT:RN23
    Set par = Server.FindObject("A.B").Parent
    MsgBox par                                                ' EXPECT:RN23
End Sub
Sub OtherProc()
    MsgBox obj                                                ' NOT:RN23
    If IsNull(obj) Then b = 1                                 ' NOT:RN24
End Sub
Sub Dates()
    Form.Init "Dates"
    Set StartTime = Form.AddDateTime(2, 2)
    Form.Show
    SetRegistry "start.date", StartTime.Value                 ' EXPECT:RN25
    SetRegistry "start.date", CLng(Int(CDbl(StartTime.Value)))   ' NOT:RN25
    SetRegistry "now", Now                                    ' EXPECT:RN25
    SetRegistry("tomorrow", DateAdd("d", 1, Now))             ' EXPECT:RN25
    App.SetRegistry "yesterday", Now - 1                      ' EXPECT:RN25
    SetRegistry "Start Time", "TIMESTAMP '" & sql & "'"       ' NOT:RN25
    SetRegistry "EndTime", Timer*10                           ' NOT:RN25
    SetRegistry "text", "Date: " & Now                        ' NOT:RN25
    t0 = Now
    SetRegistry "t0", t0                                      ' EXPECT:RN25
    n = 5
    SetRegistry "n", n                                        ' NOT:RN25
    StartTime.Value = GetRegistry("start.date", 0)            ' EXPECT:RN25
    StartTime.Value = CDate(GetRegistry("start.date", 0))     ' NOT:RN25
    MsgBox FormatDateTime(GetRegistry("start.date", 0))       ' EXPECT:RN25
    MsgBox FormatDateTime(CDate(GetRegistry("start.date", 0)))   ' NOT:RN25
    d = DateDiff("d", Now, GetRegistry("start.date", 0))      ' EXPECT:RN25
End Sub
Sub NoDate()
    SetRegistry "t0", t0                                      ' NOT:RN25
End Sub
'@
# A library that defines its own isNull (as some real systems do): its callers are not RN24.
New-Fixture 'T.KB Lib' 'CScriptLibrary' 'VBScript' @('ReferencedByMimics: 1') @'
Function isNull(o)
    isNull = o Is Nothing
End Function
'@
New-Fixture 'T.KB Own IsNull' 'CMimic' 'VBScript' @('Library: T.KB Lib') @'
Sub Check()
    Set o = Server.FindObject("A.B")
    If isNull(o) Then MsgBox "Not found"                      ' NOT:RN24
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
