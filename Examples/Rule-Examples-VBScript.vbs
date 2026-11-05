' Geo SCADA 2025 R2 script runtime - before / after examples (VBScript)
'
' Click the grey rectangle (pick action: Call Specific Function -> Run_Examples). Each issue is run twice:
'   Before - the code as commonly written for Geo SCADA 2025 and earlier
'   After  - the code that works in Geo SCADA 2025 R2 (and in earlier versions)
' On 2025 and earlier both steps should report OK. On 2025 R2 the Before step is expected to fail
' (or, for RC06a / RC06b / RC07, to look wrong) and the After step to work. Every step reports its
' result in a message box; an error in one step does not stop the run. A summary is shown at the end.
'
' Issue IDs match GeoSCADA-Script-Analyser.ps1 2.1.0 and later: RC = the numbered items of the Resource
' Centre (Knowledge Base) article "Adjustments for VBScript in Geo SCADA 2025 R2", JS = its JScript
' section, RX = other 2025 R2 changes. (Analyser 1.x used RN01-RN28, and 2.x called RC16 RX10; see the
' README rules table.)
' The only database objects used are ones that always exist - this mimic (Server.ThisObject) and the
' root group (Id 0) - plus the RC01 / RC15 test point, which is looked for first.
' RC01 and RC15 need an Internal Analog Point called "Test Analog" in the same group as this mimic, with
' no control configured. RC01 writes its value; RC15 only tests whether it has a Control aggregate. If
' the point is not there, you are asked whether to carry on without RC01 and RC15 or stop.
' RC06b is a visual check, like RC06a / RC07: look at the list box's scroll bar.
' RC09 is the exception: its Before step fails in every version (the IsNull test never worked).
' RC11 writes and reads registry key "Claude Examples - Date" (ViewX's own registry area).
' RC12 is not part of Run_Examples: it is about how a pick action passes arguments, so it needs two
' buttons of its own calling RC12_PickAction - see the RC12 section below.
' RC13 is not part of Run_Examples either: in 2025 R2 a log on stops the script, which would end the
' run there. It needs two buttons of its own, calling RC13_Before and RC13_After.
' RC16 reads registry key "Claude Examples - Never Set", which must never be set.
' Expected behaviours marked "(to confirm)" come from the release notes or a field report and have
' not yet been observed side by side on both versions.
Option Explicit

Dim Results, TestPoint      ' TestPoint: the "Test Analog" point, or Nothing (set by RC01_Run)

Public Sub Run_Examples()
    Results = ""
    Set TestPoint = Nothing
    If Not RC01_Run() Then Exit Sub
    RC02_Before : RC02_After
    RC03_Before : RC03_After
    RC04_Before : RC04_After
    RC06a_Before : RC06a_After
    RC06b_Before : RC06b_After
    RC07_Before : RC07_After
    RC09_Before : RC09_After
    RC11_Before : RC11_After
    Results = Results & "RC12: run from its own two buttons (see RC12_PickAction in the script)" & vbCrLf
    Results = Results & "RC13: run from its own two buttons (RC13_Before and RC13_After)" & vbCrLf
    RC14_Before : RC14_After
    If TestPoint Is Nothing Then
        Results = Results & "RC15: skipped - no test point" & vbCrLf
    Else
        RC15_Before TestPoint : RC15_After TestPoint
    End If
    RC16_Before : RC16_After
    JS01_Before : JS01_After
    RX01_Before : RX01_After
    MsgBox Results, vbInformation, "Geo SCADA 2025 R2 script examples - summary"
End Sub

' Records and shows the result of one step. ErrNumber / ErrText are captured by the caller.
Sub Report(Id, Stage, ErrNumber, ErrText, Expectation)
    Dim msg
    If ErrNumber <> 0 Then
        msg = Id & " " & Stage & ": FAILED - error " & ErrNumber & ": " & ErrText
    Else
        msg = Id & " " & Stage & ": OK"
    End If
    Results = Results & msg & vbCrLf
    MsgBox msg & vbCrLf & vbCrLf & Expectation, vbOKOnly, "Geo SCADA 2025 R2 script examples"
End Sub

'----------------------------------------------------------------------------------------------
' RC01 - .Value(x) used to set a value (Knowledge Base #1; release notes p20)
' Value is both a property and a method of a database item. The Value property now takes
' precedence, so x.Interface.Value(3.14) as a statement is an error. Before 2025 R2 the call was
' accepted but did not set the value. Reads such as y = (x.Interface.Value(0)) are not affected.
'
' The test point is "Test Analog" in this mimic's group. It is never created by the script: if it is
' not there, RC01_Run asks whether to carry on without RC01 (OK) or stop (Cancel) so the point can
' be created and the examples run again. RC01 writes the point's value, so never use a real tag.
'----------------------------------------------------------------------------------------------
Function RC01_Run()
    Dim pointName, p, msg
    RC01_Run = True
    pointName = Left(Server.ThisObject.FullName, InStrRev(Server.ThisObject.FullName, ".")) & "Test Analog"
    Set p = Server.FindObject(pointName)
    If p Is Nothing Then
        msg = "RC01 needs an Internal Analog Point called:" & vbCrLf & pointName & vbCrLf & vbCrLf
        msg = msg & "It was not found. RC01 writes the point's value, so never use a real tag." & vbCrLf & vbCrLf
        msg = msg & "OK - carry on without RC01 and RC15" & vbCrLf
        msg = msg & "Cancel - stop here, so you can create the point and run the examples again"
        If MsgBox(msg, vbOKCancel + vbQuestion, "Geo SCADA 2025 R2 script examples") = vbCancel Then
            RC01_Run = False
        Else
            Results = Results & "RC01: skipped - no test point " & pointName & vbCrLf
        End If
    Else
        Set TestPoint = p
        RC01_Before p : RC01_After p : RC01_Read p
    End If
End Function

Sub RC01_Before(p)
    On Error Resume Next
    p.Interface.Value(3.14)
    Report "RC01", "Before", Err.Number, Err.Description, _
        "x.Interface.Value(3.14)" & vbCrLf & "Expected: OK before 2025 R2 (the value is NOT changed); script error in 2025 R2."
End Sub

Sub RC01_After(p)
    On Error Resume Next
    p.Interface.Value = 3.14
    Report "RC01", "After", Err.Number, Err.Description, _
        "x.Interface.Value = 3.14" & vbCrLf & "Expected: OK in all versions; the point value is now 3.14."
End Sub

' Reading with an argument is not affected (the analyser does not report it).
Sub RC01_Read(p)
    Dim y
    On Error Resume Next
    y = (p.Interface.Value(0))
    Report "RC01", "Read", Err.Number, Err.Description, _
        "y = (x.Interface.Value(0))   - a read, not a set" & vbCrLf & "Expected: OK in all versions (to confirm; y = " & y & ")."
End Sub

'----------------------------------------------------------------------------------------------
' RC02 - Return value of a Menu method used (Knowledge Base #2; release notes p20-21)
' Menu.Init, Menu.AddItem and Menu.Separator returned an empty string; they now return Null.
'----------------------------------------------------------------------------------------------
Sub RC02_Before()
    Dim r
    On Error Resume Next
    r = CStr(Menu.Init())
    Report "RC02", "Before", Err.Number, Err.Description, _
        "r = CStr(Menu.Init())" & vbCrLf & "Expected: OK before 2025 R2 (empty string); 'Invalid use of Null' in 2025 R2 (to confirm)."
End Sub

Sub RC02_After()
    On Error Resume Next
    Menu.Init
    Report "RC02", "After", Err.Number, Err.Description, _
        "Menu.Init   (return value not used)" & vbCrLf & "Expected: OK in all versions."
End Sub

'----------------------------------------------------------------------------------------------
' RC03 - Form control used without .Value (Knowledge Base #3; release notes p21)
' Controls returned by Form.Add... no longer have a default property.
'----------------------------------------------------------------------------------------------
Sub RC03_Before()
    Dim f, s
    On Error Resume Next
    Form.Init "RC03"
    Set f = Form.AddEditBox(2, 2, 10, 1)
    f.Value = "test"
    s = "Value: " & f
    Report "RC03", "Before", Err.Number, Err.Description, _
        "s = ""Value: "" & field" & vbCrLf & "Expected: OK before 2025 R2; script error in 2025 R2."
End Sub

Sub RC03_After()
    Dim f, s
    On Error Resume Next
    Form.Init "RC03"
    Set f = Form.AddEditBox(2, 2, 10, 1)
    f.Value = "test"
    s = "Value: " & f.Value
    Report "RC03", "After", Err.Number, Err.Description, _
        "s = ""Value: "" & field.Value" & vbCrLf & "Expected: OK in all versions."
End Sub

'----------------------------------------------------------------------------------------------
' RC04 - Database object used as a value (Knowledge Base #4: Object Default Properties)
' A database object variable no longer has a default property, so CStr(obj) no longer gives its Id.
' Uses this mimic's own database object (Server.ThisObject), so it needs no test point.
'----------------------------------------------------------------------------------------------
Sub RC04_Before()
    Dim o, s
    On Error Resume Next
    Set o = Server.ThisObject
    s = CStr(o)
    If Err.Number = 0 Then
        If s <> CStr(o.Id) Then Err.Raise vbObjectError + 2, , "CStr(obj) gave '" & s & "', not the object Id " & o.Id
    End If
    Report "RC04", "Before", Err.Number, Err.Description, _
        "s = CStr(obj)   (obj = Server.ThisObject)" & vbCrLf & "Expected: OK before 2025 R2 (s = the object Id); error or not the Id in 2025 R2 (to confirm)."
End Sub

Sub RC04_After()
    Dim o, s
    On Error Resume Next
    Set o = Server.ThisObject
    s = CStr(o.Id)
    Report "RC04", "After", Err.Number, Err.Description, _
        "s = CStr(obj.Id)" & vbCrLf & "Expected: OK in all versions (s = " & s & ")."
End Sub

'----------------------------------------------------------------------------------------------
' RC06a - Form.AddListBox row count is now exact (Knowledge Base #6; release notes p21-22)
' Visual check. Before 2025 R2 a list box showed fewer rows than requested, so a button placed below
' it at a smaller row number did not overlap. In 2025 R2 all requested rows are shown.
'----------------------------------------------------------------------------------------------
Sub RC06a_Before()
    Dim lb, i
    On Error Resume Next
    Form.Init "RC06a Before - list asks for 20 rows, OK button at row 14"
    Set lb = Form.AddListBox(2, 2, 20, 1)
    For i = 1 To 30 : lb.Add "Item " & i : Next
    Form.AddPushButton 2, 14, "OK"
    Form.Show
    Report "RC06a", "Before", Err.Number, Err.Description, _
        "Form.AddListBox(2, 2, 20, 1) with a button at row 14" & vbCrLf & "Expected: button clear of the list before 2025 R2; button over the list in 2025 R2."
End Sub

Sub RC06a_After()
    Dim lb, i
    On Error Resume Next
    Form.Init "RC06a After - list asks for 10 rows, OK button at row 14"
    Set lb = Form.AddListBox(2, 2, 10, 1)
    For i = 1 To 30 : lb.Add "Item " & i : Next
    Form.AddPushButton 2, 14, "OK"
    Form.Show
    Report "RC06a", "After", Err.Number, Err.Description, _
        "Form.AddListBox(2, 2, 10, 1) - rows reduced" & vbCrLf & "Expected: button clear of the list in all versions."
End Sub

'----------------------------------------------------------------------------------------------
' RC06b - List box with more items than rows scrolls horizontally (Knowledge Base #6; first found in the field)
' Visual check. Before 2025 R2 a list box holding more items than its rows scrolled vertically; in
' 2025 R2 it shows a horizontal scroll bar instead. The After step makes the list tall enough for
' its items, so nothing needs to scroll.
'----------------------------------------------------------------------------------------------
Function RC06b_Items()
    Dim items(10), i
    For i = 0 To 10
        items(i) = CStr(i) & " Some long text so the scroll bar is obvious"
    Next
    RC06b_Items = items
End Function

Sub RC06b_Before()
    Dim items, lb
    On Error Resume Next
    items = RC06b_Items()
    Form.Init "RC06b Before - 11 items in a 3-row list"
    Set lb = Form.AddListBox(1, 1, 3, 1, items)
    Form.Show
    Report "RC06b", "Before", Err.Number, Err.Description, _
        "Form.AddListBox(1, 1, 3, 1, items) - 11 items, 3 rows" & vbCrLf & "Expected: vertical scroll bar before 2025 R2; horizontal scroll bar in 2025 R2."
End Sub

Sub RC06b_After()
    Dim items, lb
    On Error Resume Next
    items = RC06b_Items()
    Form.Init "RC06b After - 11 items in an 11-row list"
    Set lb = Form.AddListBox(1, 1, 11, 1, items)
    Form.Show
    Report "RC06b", "After", Err.Number, Err.Description, _
        "Form.AddListBox(1, 1, 11, 1, items) - the list is tall enough for its items" & vbCrLf & _
        "Expected: no scroll bar in 2025 R2 (earlier versions showed fewer rows than requested - see RC06a - so may still scroll vertically)."
End Sub

'----------------------------------------------------------------------------------------------
' RC07 - Form.AddQuery row count is now exact (Knowledge Base #7; release notes p22-23)
' Visual check, as RC06a, for a query list.
'----------------------------------------------------------------------------------------------
Sub RC07_Before()
    On Error Resume Next
    Form.Init "RC07 Before - query asks for 20 rows, OK button at row 14"
    Form.AddQuery 2, 2, 60, 20, "SELECT TOP(30) FullName FROM CGroup ORDER BY FullName"
    Form.AddPushButton 2, 14, "OK"
    Form.Show
    Report "RC07", "Before", Err.Number, Err.Description, _
        "Form.AddQuery(2, 2, 60, 20, sql) with a button at row 14" & vbCrLf & "Expected: button clear of the list before 2025 R2; button over the list in 2025 R2."
End Sub

Sub RC07_After()
    On Error Resume Next
    Form.Init "RC07 After - query asks for 10 rows, OK button at row 14"
    Form.AddQuery 2, 2, 60, 10, "SELECT TOP(30) FullName FROM CGroup ORDER BY FullName"
    Form.AddPushButton 2, 14, "OK"
    Form.Show
    Report "RC07", "After", Err.Number, Err.Description, _
        "Form.AddQuery(2, 2, 60, 10, sql) - rows reduced" & vbCrLf & "Expected: button clear of the list in all versions."
End Sub

'----------------------------------------------------------------------------------------------
' RC09 - Object not found tested with IsNull (Knowledge Base #9: Server.LookupObject returns Nothing
' if the object is not found, not Null). NOT a 2025 R2 change: FindObject and LookupObject have
' always returned Nothing, so the IsNull test never sees a missing object and the Before step fails
' in EVERY version. The object looked for is a child of this mimic, which can never exist because a
' mimic cannot contain other objects.
'----------------------------------------------------------------------------------------------
Function RC09_MissingName()
    RC09_MissingName = Server.ThisObject.FullName & ".RC09 No Such Object"
End Function

Sub RC09_Before()
    Dim o, s
    On Error Resume Next
    Set o = Server.FindObject(RC09_MissingName())
    If IsNull(o) Then
        s = "not found"
    Else
        s = CStr(o.Id)              ' reached for a missing object too: "Object required"
    End If
    Report "RC09", "Before", Err.Number, Err.Description, _
        "If IsNull(obj) Then ... Else ... obj.Id" & vbCrLf & "Expected: error (Object required) in ALL versions - not an R2 change, the test never worked."
End Sub

Sub RC09_After()
    Dim o, s
    On Error Resume Next
    Set o = Server.FindObject(RC09_MissingName())
    If o Is Nothing Then
        s = "not found"
    Else
        s = CStr(o.Id)
    End If
    Report "RC09", "After", Err.Number, Err.Description, _
        "If obj Is Nothing Then ..." & vbCrLf & "Expected: OK in all versions (s = " & s & ")."
End Sub

'----------------------------------------------------------------------------------------------
' RC11 - Date written to the registry without conversion (Knowledge Base #11: No implicit date/time to
' number conversions or vice versa). Store the date as a Long and read it back with CDate.
'----------------------------------------------------------------------------------------------
Sub RC11_Before()
    Dim t
    On Error Resume Next
    SetRegistry "Claude Examples - Date", Now
    t = FormatDateTime(GetRegistry("Claude Examples - Date", 0), vbShortDate)
    Report "RC11", "Before", Err.Number, Err.Description, _
        "SetRegistry ""Claude Examples - Date"", Now" & vbCrLf & "Expected: OK before 2025 R2 (t = " & t & "); error in 2025 R2 (to confirm)."
End Sub

Sub RC11_After()
    Dim t
    On Error Resume Next
    SetRegistry "Claude Examples - Date", CLng(Int(CDbl(Now)))
    t = FormatDateTime(CDate(GetRegistry("Claude Examples - Date", 0)), vbShortDate)
    Report "RC11", "After", Err.Number, Err.Description, _
        "SetRegistry ""Claude Examples - Date"", CLng(Int(CDbl(Now)))   and   CDate(GetRegistry(...))" & vbCrLf & "Expected: OK in all versions (t = " & t & ")."
End Sub

'----------------------------------------------------------------------------------------------
' RC12 - Pick action leaves a function parameter without a value (Knowledge Base #12: You need to provide
' all pick action parameters or the script will not run; first found in the field)
' In the 2025 R2 release version a Call Specific Function pick action whose function has a parameter
' with no value in the pick action configuration does not run at all, and no error is shown. Before
' 2025 R2 the missing parameter was passed as Empty. Schneider expect a future update to fix this:
' check the current release notes for the version installed.
'
' Not part of Run_Examples. Add two buttons to this mimic:
'   "RC12 Before" - Call Specific Function -> RC12_PickAction, One = 'hello', Two left blank
'   "RC12 After"  - Call Specific Function -> RC12_PickAction, One = 'hello', Two = 'goodbye'
' Expected: before 2025 R2 both buttons show the message (Before shows Two as Empty).
' In the 2025 R2 release the Before button does nothing at all (until the expected fix); the After
' button works in all versions.
'----------------------------------------------------------------------------------------------
Public Sub RC12_PickAction(One, Two)
    Dim twoText
    If IsEmpty(Two) Then
        twoText = "(Empty - not set in the pick action)"
    Else
        twoText = "'" & Two & "'"
    End If
    MsgBox "RC12: the function ran." & vbCrLf & vbCrLf & "One = '" & One & "'" & vbCrLf & "Two = " & twoText, _
        vbInformation, "Geo SCADA 2025 R2 script examples"
End Sub

'----------------------------------------------------------------------------------------------
' RC13 - Log on / log off stops the script (Knowledge Base #13: App.LogOn/Off and App.LogOnUser/OffUser
' and COM .LogOn will stop execution)
' In 2025 R2 App.LogOn, App.LogOff, App.LogOnUser, App.LogOffUser and a Logon through ViewX COM
' automation stop the script, so nothing after the call runs. Do the other work first.
'
' Not part of Run_Examples (it would end the run there). Add two buttons to this mimic:
'   "RC13 Before" - Call Specific Function -> RC13_Before
'   "RC13 After"  - Call Specific Function -> RC13_After
' Each opens the ViewX Log On window: log on again as the same user, or cancel.
' Expected: before 2025 R2 the Before button shows its message once the Log On window closes; in
' 2025 R2 it never does (to confirm, also whether cancelling stops the script). The After button
' shows its message first, then opens the Log On window, in all versions.
'----------------------------------------------------------------------------------------------
Public Sub RC13_Before()
    App.LogOn
    MsgBox "RC13 Before: the script carried on after App.LogOn." & vbCrLf & vbCrLf & _
        "Expected: shown before 2025 R2; never shown in 2025 R2.", vbInformation, "Geo SCADA 2025 R2 script examples"
End Sub

Public Sub RC13_After()
    MsgBox "RC13 After: everything else is done first. App.LogOn is the last statement, so nothing is lost when the script stops." & vbCrLf & vbCrLf & _
        "Expected: shown in all versions, then the Log On window opens.", vbInformation, "Geo SCADA 2025 R2 script examples"
    App.LogOn
End Sub

'----------------------------------------------------------------------------------------------
' RC14 - GetOPCValue cannot read aggregates (Knowledge Base #14: The Server.GetOPCValue method cannot
' read aggregates)
' In the 2025 R2 release version Server.GetOPCValue raises an error for an aggregate, or a property
' inside one. Schneider describe this as a temporary limitation that a future update will fix (with
' errors still raised, under different codes): check the current release notes for the version
' installed. Reading through the object's Interface works in all versions.
' Uses this mimic's GISLocation aggregate, which every database object has and which cannot be
' disabled, so it needs no test point.
'----------------------------------------------------------------------------------------------
Sub RC14_Before()
    Dim y
    On Error Resume Next
    y = Server.GetOPCValue(Server.ThisObject.FullName & ".GISLocation.Latitude")
    Report "RC14", "Before", Err.Number, Err.Description, _
        "Server.GetOPCValue(mimic & "".GISLocation.Latitude"")" & vbCrLf & "Expected: OK before 2025 R2 (y = " & y & "); error in the 2025 R2 release until the expected fix (to confirm)."
End Sub

Sub RC14_After()
    Dim y
    On Error Resume Next
    y = Server.ThisObject.Interface.GISLocation.Latitude
    Report "RC14", "After", Err.Number, Err.Description, _
        "Server.ThisObject.Interface.GISLocation.Latitude" & vbCrLf & "Expected: OK in all versions (to confirm; y = " & y & ")."
End Sub

'----------------------------------------------------------------------------------------------
' RC15 - Testing for an aggregate (Knowledge Base #15: Changes to checking an Aggregate, e.g. to test
' for Controls)
' Before 2025 R2 a script could Set a variable to the aggregate and test it with Is Nothing. In 2025 R2
' the article says to omit the Set and test with IsObject.
' Uses the "Test Analog" point (see RC01), which should have no control configured; it is only read.
'----------------------------------------------------------------------------------------------
Sub RC15_Before(p)
    Dim y, s
    On Error Resume Next
    Set y = p.Interface.Control
    If y Is Nothing Then
        s = "has no control"
    Else
        s = "has control"
    End If
    Report "RC15", "Before", Err.Number, Err.Description, _
        "Set y = x.Interface.Control : If y Is Nothing Then" & vbCrLf & "Expected: OK before 2025 R2 (" & s & "); error or wrong answer in 2025 R2 (to confirm)."
End Sub

Sub RC15_After(p)
    Dim y, s
    On Error Resume Next
    y = p.Interface.Control
    If Not IsObject(y) Then
        s = "has no control"
    Else
        s = "has control"
    End If
    Report "RC15", "After", Err.Number, Err.Description, _
        "y = x.Interface.Control : If Not IsObject(y) Then" & vbCrLf & "Expected: OK in 2025 R2 (" & s & "); before 2025 R2 to confirm."
End Sub

'----------------------------------------------------------------------------------------------
' RC16 - GetRegistry without a default value (Knowledge Base #16: GetRegistry returns Null if there is
' no default; first found in the field; RX10 in analyser 2.x)
' A key that has not been set comes back as Null, which in 2025 R2 cannot be used as a string, so
' converting it fails. Pass a default, or test the result with IsNull.
'----------------------------------------------------------------------------------------------
Sub RC16_Before()
    Dim n
    On Error Resume Next
    n = CLng(GetRegistry("Claude Examples - Never Set"))
    Report "RC16", "Before", Err.Number, Err.Description, _
        "n = CLng(GetRegistry(""Claude Examples - Never Set""))" & vbCrLf & "Expected: OK before 2025 R2 (to confirm); conversion error in 2025 R2."
End Sub

Sub RC16_After()
    Dim n
    On Error Resume Next
    n = CLng(GetRegistry("Claude Examples - Never Set", 0))
    Report "RC16", "After", Err.Number, Err.Description, _
        "n = CLng(GetRegistry(""Claude Examples - Never Set"", 0))" & vbCrLf & "Expected: OK in all versions (n = " & n & ")."
End Sub

'----------------------------------------------------------------------------------------------
' JS01 - ScriptControl COM object (Knowledge Base: Specific JScript Issues; release notes p23)
' ScriptControl (msscript.ocx) is 32-bit and not recommended. Do the work in the script itself.
'----------------------------------------------------------------------------------------------
Sub JS01_Before()
    Dim sc, r
    On Error Resume Next
    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "VBScript"
    r = sc.Eval("1 + 1")
    Report "JS01", "Before", Err.Number, Err.Description, _
        "CreateObject(""MSScriptControl.ScriptControl"").Eval(""1 + 1"")" & vbCrLf & "Expected: OK before 2025 R2; error in 2025 R2."
End Sub

Sub JS01_After()
    Dim r
    On Error Resume Next
    r = Eval("1 + 1")
    Report "JS01", "After", Err.Number, Err.Description, _
        "r = Eval(""1 + 1"")  - no COM object" & vbCrLf & "Expected: OK in all versions (r = " & r & ")."
End Sub

'----------------------------------------------------------------------------------------------
' RX01 - 32-bit only COM class (release notes p16)
' ViewX is now 64-bit and can only create 64-bit (or out-of-process) COM objects. DAO 3.6 (Jet) is a
' 32-bit only component; ADO is available as 64-bit.
'----------------------------------------------------------------------------------------------
Sub RX01_Before()
    Dim o
    On Error Resume Next
    Set o = CreateObject("DAO.DBEngine.36")
    Report "RX01", "Before", Err.Number, Err.Description, _
        "CreateObject(""DAO.DBEngine.36"")  - 32-bit only" & vbCrLf & "Expected: OK before 2025 R2 (32-bit ViewX, if DAO is installed); error in 2025 R2."
End Sub

Sub RX01_After()
    Dim o
    On Error Resume Next
    Set o = CreateObject("ADODB.Connection")
    Report "RX01", "After", Err.Number, Err.Description, _
        "CreateObject(""ADODB.Connection"")  - 32 and 64-bit" & vbCrLf & "Expected: OK in all versions."
End Sub
