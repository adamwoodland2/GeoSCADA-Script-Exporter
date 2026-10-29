' Geo SCADA 2025 R2 script runtime - before / after examples (VBScript)
'
' Click the grey rectangle (pick action: Call Specific Function -> Run_Examples). Each issue is run twice:
'   Before - the code as commonly written for Geo SCADA 2025 and earlier
'   After  - the code that works in Geo SCADA 2025 R2 (and in earlier versions)
' On 2025 and earlier both steps should report OK. On 2025 R2 the Before step is expected to fail
' (or, for RN04 / RN05 / RN27, to look wrong) and the After step to work. Every step reports its result in
' a message box; an error in one step does not stop the run. A summary is shown at the end.
'
' Issue numbers match GeoSCADA-Script-Analyser.ps1 (RN = 2025 R2 script runtime changes).
' The only database objects used are ones that always exist - this mimic (Server.ThisObject) and the
' root group (Id 0) - plus RN01's test point, which is looked for first.
' RN01 needs an Internal Analog Point called "Test Analog" in the same group as this mimic, and writes
' its value. If the point is not there, you are asked whether to carry on without RN01 or stop.
' RN21 reads registry key "Claude Examples - Never Set", which must never be set.
' RN24 is the exception: its Before step fails in every version (the IsNull test never worked).
' RN25 writes and reads registry key "Claude Examples - Date" (ViewX's own registry area).
' RN26 is not part of Run_Examples: it is about how a pick action passes arguments, so it needs two
' buttons of its own calling RN26_PickAction - see the RN26 section below.
' RN27 is a visual check, like RN04 / RN05: look at the list box's scroll bar.
' Expected behaviours marked "(to confirm)" come from the release notes or a field report and have
' not yet been observed side by side on both versions.
Option Explicit

Dim Results

Public Sub Run_Examples()
    Results = ""
    If Not RN01_Run() Then Exit Sub
    RN02_Before : RN02_After
    RN03_Before : RN03_After
    RN04_Before : RN04_After
    RN05_Before : RN05_After
    RN06_Before : RN06_After
    RN10_Before : RN10_After
    RN21_Before : RN21_After
    RN23_Before : RN23_After
    RN24_Before : RN24_After
    RN25_Before : RN25_After
    RN27_Before : RN27_After
    Results = Results & "RN26: run from its own two buttons (see RN26_PickAction in the script)" & vbCrLf
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
' RN01 - .Value(x) used to set a value (release notes p20)
' Value is both a property and a method of a database item. The Value property now takes
' precedence, so x.Interface.Value(3.14) as a statement is an error. Before 2025 R2 the call was
' accepted but did not set the value. Reads such as y = (x.Interface.Value(0)) are not affected.
'
' The test point is "Test Analog" in this mimic's group. It is never created by the script: if it is
' not there, RN01_Run asks whether to carry on without RN01 (OK) or stop (Cancel) so the point can
' be created and the examples run again. RN01 writes the point's value, so never use a real tag.
'----------------------------------------------------------------------------------------------
Function RN01_Run()
    Dim pointName, p, msg
    RN01_Run = True
    pointName = Left(Server.ThisObject.FullName, InStrRev(Server.ThisObject.FullName, ".")) & "Test Analog"
    Set p = Server.FindObject(pointName)
    If p Is Nothing Then
        msg = "RN01 needs an Internal Analog Point called:" & vbCrLf & pointName & vbCrLf & vbCrLf
        msg = msg & "It was not found. RN01 writes the point's value, so never use a real tag." & vbCrLf & vbCrLf
        msg = msg & "OK - carry on without RN01" & vbCrLf
        msg = msg & "Cancel - stop here, so you can create the point and run the examples again"
        If MsgBox(msg, vbOKCancel + vbQuestion, "Geo SCADA 2025 R2 script examples") = vbCancel Then
            RN01_Run = False
        Else
            Results = Results & "RN01: skipped - no test point " & pointName & vbCrLf
        End If
    Else
        RN01_Before p : RN01_After p : RN01_Read p
    End If
End Function

Sub RN01_Before(p)
    On Error Resume Next
    p.Interface.Value(3.14)
    Report "RN01", "Before", Err.Number, Err.Description, _
        "x.Interface.Value(3.14)" & vbCrLf & "Expected: OK before 2025 R2 (the value is NOT changed); script error in 2025 R2."
End Sub

Sub RN01_After(p)
    On Error Resume Next
    p.Interface.Value = 3.14
    Report "RN01", "After", Err.Number, Err.Description, _
        "x.Interface.Value = 3.14" & vbCrLf & "Expected: OK in all versions; the point value is now 3.14."
End Sub

' Reading with an argument is not affected (the analyser does not report it).
Sub RN01_Read(p)
    Dim y
    On Error Resume Next
    y = (p.Interface.Value(0))
    Report "RN01", "Read", Err.Number, Err.Description, _
        "y = (x.Interface.Value(0))   - a read, not a set" & vbCrLf & "Expected: OK in all versions (to confirm; y = " & y & ")."
End Sub

'----------------------------------------------------------------------------------------------
' RN02 - Return value of a Menu method used (release notes p20-21)
' Menu.Init, Menu.AddItem and Menu.Separator returned an empty string; they now return Null.
'----------------------------------------------------------------------------------------------
Sub RN02_Before()
    Dim r
    On Error Resume Next
    r = CStr(Menu.Init())
    Report "RN02", "Before", Err.Number, Err.Description, _
        "r = CStr(Menu.Init())" & vbCrLf & "Expected: OK before 2025 R2 (empty string); 'Invalid use of Null' in 2025 R2 (to confirm)."
End Sub

Sub RN02_After()
    On Error Resume Next
    Menu.Init
    Report "RN02", "After", Err.Number, Err.Description, _
        "Menu.Init   (return value not used)" & vbCrLf & "Expected: OK in all versions."
End Sub

'----------------------------------------------------------------------------------------------
' RN03 - Form control used without .Value (release notes p21)
' Controls returned by Form.Add... no longer have a default property.
'----------------------------------------------------------------------------------------------
Sub RN03_Before()
    Dim f, s
    On Error Resume Next
    Form.Init "RN03"
    Set f = Form.AddEditBox(2, 2, 10, 1)
    f.Value = "test"
    s = "Value: " & f
    Report "RN03", "Before", Err.Number, Err.Description, _
        "s = ""Value: "" & field" & vbCrLf & "Expected: OK before 2025 R2; script error in 2025 R2."
End Sub

Sub RN03_After()
    Dim f, s
    On Error Resume Next
    Form.Init "RN03"
    Set f = Form.AddEditBox(2, 2, 10, 1)
    f.Value = "test"
    s = "Value: " & f.Value
    Report "RN03", "After", Err.Number, Err.Description, _
        "s = ""Value: "" & field.Value" & vbCrLf & "Expected: OK in all versions."
End Sub

'----------------------------------------------------------------------------------------------
' RN04 - Form.AddListBox row count is now exact (release notes p21-22)
' Visual check. Before 2025 R2 a list box showed fewer rows than requested, so a button placed below
' it at a smaller row number did not overlap. In 2025 R2 all requested rows are shown.
'----------------------------------------------------------------------------------------------
Sub RN04_Before()
    Dim lb, i
    On Error Resume Next
    Form.Init "RN04 Before - list asks for 20 rows, OK button at row 14"
    Set lb = Form.AddListBox(2, 2, 20, 1)
    For i = 1 To 30 : lb.Add "Item " & i : Next
    Form.AddPushButton 2, 14, "OK"
    Form.Show
    Report "RN04", "Before", Err.Number, Err.Description, _
        "Form.AddListBox(2, 2, 20, 1) with a button at row 14" & vbCrLf & "Expected: button clear of the list before 2025 R2; button over the list in 2025 R2."
End Sub

Sub RN04_After()
    Dim lb, i
    On Error Resume Next
    Form.Init "RN04 After - list asks for 10 rows, OK button at row 14"
    Set lb = Form.AddListBox(2, 2, 10, 1)
    For i = 1 To 30 : lb.Add "Item " & i : Next
    Form.AddPushButton 2, 14, "OK"
    Form.Show
    Report "RN04", "After", Err.Number, Err.Description, _
        "Form.AddListBox(2, 2, 10, 1) - rows reduced" & vbCrLf & "Expected: button clear of the list in all versions."
End Sub

'----------------------------------------------------------------------------------------------
' RN05 - Form.AddQuery row count is now exact (release notes p22-23)
' Visual check, as RN04, for a query list.
'----------------------------------------------------------------------------------------------
Sub RN05_Before()
    On Error Resume Next
    Form.Init "RN05 Before - query asks for 20 rows, OK button at row 14"
    Form.AddQuery 2, 2, 60, 20, "SELECT TOP(30) FullName FROM CGroup ORDER BY FullName"
    Form.AddPushButton 2, 14, "OK"
    Form.Show
    Report "RN05", "Before", Err.Number, Err.Description, _
        "Form.AddQuery(2, 2, 60, 20, sql) with a button at row 14" & vbCrLf & "Expected: button clear of the list before 2025 R2; button over the list in 2025 R2."
End Sub

Sub RN05_After()
    On Error Resume Next
    Form.Init "RN05 After - query asks for 10 rows, OK button at row 14"
    Form.AddQuery 2, 2, 60, 10, "SELECT TOP(30) FullName FROM CGroup ORDER BY FullName"
    Form.AddPushButton 2, 14, "OK"
    Form.Show
    Report "RN05", "After", Err.Number, Err.Description, _
        "Form.AddQuery(2, 2, 60, 10, sql) - rows reduced" & vbCrLf & "Expected: button clear of the list in all versions."
End Sub

'----------------------------------------------------------------------------------------------
' RN06 - 32-bit only COM class (release notes p16)
' ViewX is now 64-bit and can only create 64-bit (or out-of-process) COM objects. DAO 3.6 (Jet) is a
' 32-bit only component; ADO is available as 64-bit.
'----------------------------------------------------------------------------------------------
Sub RN06_Before()
    Dim o
    On Error Resume Next
    Set o = CreateObject("DAO.DBEngine.36")
    Report "RN06", "Before", Err.Number, Err.Description, _
        "CreateObject(""DAO.DBEngine.36"")  - 32-bit only" & vbCrLf & "Expected: OK before 2025 R2 (32-bit ViewX, if DAO is installed); error in 2025 R2."
End Sub

Sub RN06_After()
    Dim o
    On Error Resume Next
    Set o = CreateObject("ADODB.Connection")
    Report "RN06", "After", Err.Number, Err.Description, _
        "CreateObject(""ADODB.Connection"")  - 32 and 64-bit" & vbCrLf & "Expected: OK in all versions."
End Sub

'----------------------------------------------------------------------------------------------
' RN10 - ScriptControl COM object (release notes p23)
' ScriptControl (msscript.ocx) is 32-bit and not recommended. Do the work in the script itself.
'----------------------------------------------------------------------------------------------
Sub RN10_Before()
    Dim sc, r
    On Error Resume Next
    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "VBScript"
    r = sc.Eval("1 + 1")
    Report "RN10", "Before", Err.Number, Err.Description, _
        "CreateObject(""MSScriptControl.ScriptControl"").Eval(""1 + 1"")" & vbCrLf & "Expected: OK before 2025 R2; error in 2025 R2."
End Sub

Sub RN10_After()
    Dim r
    On Error Resume Next
    r = Eval("1 + 1")
    Report "RN10", "After", Err.Number, Err.Description, _
        "r = Eval(""1 + 1"")  - no COM object" & vbCrLf & "Expected: OK in all versions (r = " & r & ")."
End Sub

'----------------------------------------------------------------------------------------------
' RN21 - GetRegistry without a default value (2025 R2 runtime, found in the field)
' A key that has not been set comes back as Null / an empty string, which fails when converted.
'----------------------------------------------------------------------------------------------
Sub RN21_Before()
    Dim n
    On Error Resume Next
    n = CLng(GetRegistry("Claude Examples - Never Set"))
    Report "RN21", "Before", Err.Number, Err.Description, _
        "n = CLng(GetRegistry(""Claude Examples - Never Set""))" & vbCrLf & "Expected: OK before 2025 R2 (to confirm); conversion error in 2025 R2."
End Sub

Sub RN21_After()
    Dim n
    On Error Resume Next
    n = CLng(GetRegistry("Claude Examples - Never Set", 0))
    Report "RN21", "After", Err.Number, Err.Description, _
        "n = CLng(GetRegistry(""Claude Examples - Never Set"", 0))" & vbCrLf & "Expected: OK in all versions (n = " & n & ")."
End Sub

'----------------------------------------------------------------------------------------------
' RN23 - Database object used as a value (Knowledge Base: Object Default Properties)
' A database object variable no longer has a default property, so CStr(obj) no longer gives its Id.
' Uses this mimic's own database object (Server.ThisObject), so it needs no test point.
'----------------------------------------------------------------------------------------------
Sub RN23_Before()
    Dim o, s
    On Error Resume Next
    Set o = Server.ThisObject
    s = CStr(o)
    If Err.Number = 0 Then
        If s <> CStr(o.Id) Then Err.Raise vbObjectError + 2, , "CStr(obj) gave '" & s & "', not the object Id " & o.Id
    End If
    Report "RN23", "Before", Err.Number, Err.Description, _
        "s = CStr(obj)   (obj = Server.ThisObject)" & vbCrLf & "Expected: OK before 2025 R2 (s = the object Id); error or not the Id in 2025 R2 (to confirm)."
End Sub

Sub RN23_After()
    Dim o, s
    On Error Resume Next
    Set o = Server.ThisObject
    s = CStr(o.Id)
    Report "RN23", "After", Err.Number, Err.Description, _
        "s = CStr(obj.Id)" & vbCrLf & "Expected: OK in all versions (s = " & s & ")."
End Sub

'----------------------------------------------------------------------------------------------
' RN24 - Object not found tested with IsNull (Knowledge Base: Server.LookupObject returns Nothing
' if the object is not found, not Null). NOT a 2025 R2 change: FindObject and LookupObject have
' always returned Nothing, so the IsNull test never sees a missing object and the Before step fails
' in EVERY version. The object looked for is a child of this mimic, which can never exist because a
' mimic cannot contain other objects.
'----------------------------------------------------------------------------------------------
Function RN24_MissingName()
    RN24_MissingName = Server.ThisObject.FullName & ".RN24 No Such Object"
End Function

Sub RN24_Before()
    Dim o, s
    On Error Resume Next
    Set o = Server.FindObject(RN24_MissingName())
    If IsNull(o) Then
        s = "not found"
    Else
        s = CStr(o.Id)              ' reached for a missing object too: "Object required"
    End If
    Report "RN24", "Before", Err.Number, Err.Description, _
        "If IsNull(obj) Then ... Else ... obj.Id" & vbCrLf & "Expected: error (Object required) in ALL versions - not an R2 change, the test never worked."
End Sub

Sub RN24_After()
    Dim o, s
    On Error Resume Next
    Set o = Server.FindObject(RN24_MissingName())
    If o Is Nothing Then
        s = "not found"
    Else
        s = CStr(o.Id)
    End If
    Report "RN24", "After", Err.Number, Err.Description, _
        "If obj Is Nothing Then ..." & vbCrLf & "Expected: OK in all versions (s = " & s & ")."
End Sub

'----------------------------------------------------------------------------------------------
' RN25 - Date written to the registry without conversion (Knowledge Base: No implicit date/time to
' number conversions or vice versa). Store the date as a Long and read it back with CDate.
'----------------------------------------------------------------------------------------------
Sub RN25_Before()
    Dim t
    On Error Resume Next
    SetRegistry "Claude Examples - Date", Now
    t = FormatDateTime(GetRegistry("Claude Examples - Date", 0), vbShortDate)
    Report "RN25", "Before", Err.Number, Err.Description, _
        "SetRegistry ""Claude Examples - Date"", Now" & vbCrLf & "Expected: OK before 2025 R2 (t = " & t & "); error in 2025 R2 (to confirm)."
End Sub

Sub RN25_After()
    Dim t
    On Error Resume Next
    SetRegistry "Claude Examples - Date", CLng(Int(CDbl(Now)))
    t = FormatDateTime(CDate(GetRegistry("Claude Examples - Date", 0)), vbShortDate)
    Report "RN25", "After", Err.Number, Err.Description, _
        "SetRegistry ""Claude Examples - Date"", CLng(Int(CDbl(Now)))   and   CDate(GetRegistry(...))" & vbCrLf & "Expected: OK in all versions (t = " & t & ")."
End Sub

'----------------------------------------------------------------------------------------------
' RN26 - Pick action leaves a function parameter without a value (found in the field)
' In 2025 R2 a Call Specific Function pick action whose function has a parameter with no value in
' the pick action configuration does not run at all, and no error is shown. Before 2025 R2 the
' missing parameter was passed as Empty.
'
' Not part of Run_Examples. Add two buttons to this mimic:
'   "RN26 Before" - Call Specific Function -> RN26_PickAction, One = 'hello', Two left blank
'   "RN26 After"  - Call Specific Function -> RN26_PickAction, One = 'hello', Two = 'goodbye'
' Expected: before 2025 R2 both buttons show the message (Before shows Two as Empty).
' In 2025 R2 the Before button does nothing at all; the After button works.
'----------------------------------------------------------------------------------------------
Public Sub RN26_PickAction(One, Two)
    Dim twoText
    If IsEmpty(Two) Then
        twoText = "(Empty - not set in the pick action)"
    Else
        twoText = "'" & Two & "'"
    End If
    MsgBox "RN26: the function ran." & vbCrLf & vbCrLf & "One = '" & One & "'" & vbCrLf & "Two = " & twoText, _
        vbInformation, "Geo SCADA 2025 R2 script examples"
End Sub

'----------------------------------------------------------------------------------------------
' RN27 - List box with more items than rows scrolls horizontally (found in the field)
' Visual check. Before 2025 R2 a list box holding more items than its rows scrolled vertically; in
' 2025 R2 it shows a horizontal scroll bar instead. The After step makes the list tall enough for
' its items, so nothing needs to scroll.
'----------------------------------------------------------------------------------------------
Function RN27_Items()
    Dim items(10), i
    For i = 0 To 10
        items(i) = CStr(i) & " Some long text so the scroll bar is obvious"
    Next
    RN27_Items = items
End Function

Sub RN27_Before()
    Dim items, lb
    On Error Resume Next
    items = RN27_Items()
    Form.Init "RN27 Before - 11 items in a 3-row list"
    Set lb = Form.AddListBox(1, 1, 3, 1, items)
    Form.Show
    Report "RN27", "Before", Err.Number, Err.Description, _
        "Form.AddListBox(1, 1, 3, 1, items) - 11 items, 3 rows" & vbCrLf & "Expected: vertical scroll bar before 2025 R2; horizontal scroll bar in 2025 R2."
End Sub

Sub RN27_After()
    Dim items, lb
    On Error Resume Next
    items = RN27_Items()
    Form.Init "RN27 After - 11 items in an 11-row list"
    Set lb = Form.AddListBox(1, 1, 11, 1, items)
    Form.Show
    Report "RN27", "After", Err.Number, Err.Description, _
        "Form.AddListBox(1, 1, 11, 1, items) - the list is tall enough for its items" & vbCrLf & _
        "Expected: no scroll bar in 2025 R2 (earlier versions showed fewer rows than requested - see RN04 - so may still scroll vertically)."
End Sub
