' Geo SCADA 2025 R2 script runtime - before / after examples (VBScript)
'
' Click the grey rectangle (pick action: Call Specific Function -> Run_Examples). Each issue is run twice:
'   Before - the code as commonly written for Geo SCADA 2025 and earlier
'   After  - the code that works in Geo SCADA 2025 R2 (and in earlier versions)
' On 2025 and earlier both steps should report OK. On 2025 R2 the Before step is expected to fail
' (or, for RN04 / RN05, to look wrong) and the After step to work. Every step reports its result in
' a message box; an error in one step does not stop the run. A summary is shown at the end.
'
' Issue numbers match GeoSCADA-Script-Analyser.ps1 (RN = 2025 R2 script runtime changes).
' RN01 needs an Internal Analog Point called "Test Analog" in the same group as this mimic.
' RN21 reads registry key "Claude Examples - Never Set", which must never be set.
' RN24 is the exception: its Before step fails in every version (the IsNull test never worked).
' RN25 writes and reads registry key "Claude Examples - Date" (ViewX's own registry area).
' Expected behaviours marked "(to confirm)" come from the release notes or a field report and have
' not yet been observed side by side on both versions.
Option Explicit

Dim Results

Public Sub Run_Examples()
    Results = ""
    RN01_Before : RN01_After : RN01_Read
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

' Full name of the test point next to this mimic: "<this mimic's group>.Test Analog".
Function TestPointName()
    TestPointName = Left(Mimic.FullName, InStrRev(Mimic.FullName, ".")) & "Test Analog"
End Function

'----------------------------------------------------------------------------------------------
' RN01 - .Value(x) used to set a value (release notes p20)
' Value is both a property and a method of a database item. The Value property now takes
' precedence, so x.Interface.Value(3.14) as a statement is an error. Before 2025 R2 the call was
' accepted but did not set the value. Reads such as y = (x.Interface.Value(0)) are not affected.
'----------------------------------------------------------------------------------------------
Sub RN01_Before()
    Dim p
    On Error Resume Next
    Set p = Server.FindObject(TestPointName())
    If p Is Nothing Then
        Err.Raise vbObjectError + 1, , "Test point " & TestPointName() & " not found - create it to run RN01"
    Else
        p.Interface.Value(3.14)
    End If
    Report "RN01", "Before", Err.Number, Err.Description, _
        "x.Interface.Value(3.14)" & vbCrLf & "Expected: OK before 2025 R2 (the value is NOT changed); script error in 2025 R2."
End Sub

Sub RN01_After()
    Dim p
    On Error Resume Next
    Set p = Server.FindObject(TestPointName())
    If p Is Nothing Then
        Err.Raise vbObjectError + 1, , "Test point " & TestPointName() & " not found - create it to run RN01"
    Else
        p.Interface.Value = 3.14
    End If
    Report "RN01", "After", Err.Number, Err.Description, _
        "x.Interface.Value = 3.14" & vbCrLf & "Expected: OK in all versions; the point value is now 3.14."
End Sub

' Reading with an argument is not affected (the analyser does not report it).
Sub RN01_Read()
    Dim p, y
    On Error Resume Next
    Set p = Server.FindObject(TestPointName())
    If p Is Nothing Then
        Err.Raise vbObjectError + 1, , "Test point " & TestPointName() & " not found - create it to run RN01"
    Else
        y = (p.Interface.Value(0))
    End If
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
' Uses this mimic's own database object, so it needs no test point.
'----------------------------------------------------------------------------------------------
Sub RN23_Before()
    Dim o, s
    On Error Resume Next
    Set o = Server.FindObject(Mimic.FullName)
    s = CStr(o)
    If Err.Number = 0 Then
        If s <> CStr(o.Id) Then Err.Raise vbObjectError + 2, , "CStr(obj) gave '" & s & "', not the object Id " & o.Id
    End If
    Report "RN23", "Before", Err.Number, Err.Description, _
        "s = CStr(obj)   (obj = Server.FindObject(...))" & vbCrLf & "Expected: OK before 2025 R2 (s = the object Id); error or not the Id in 2025 R2 (to confirm)."
End Sub

Sub RN23_After()
    Dim o, s
    On Error Resume Next
    Set o = Server.FindObject(Mimic.FullName)
    s = CStr(o.Id)
    Report "RN23", "After", Err.Number, Err.Description, _
        "s = CStr(obj.Id)" & vbCrLf & "Expected: OK in all versions (s = " & s & ")."
End Sub

'----------------------------------------------------------------------------------------------
' RN24 - Object not found tested with IsNull (Knowledge Base: Server.LookupObject returns Nothing
' if the object is not found, not Null). NOT a 2025 R2 change: LookupObject has always returned
' Nothing, so the IsNull test never sees a missing object and the Before step fails in EVERY
' version. Object Id 99999999 must not exist (as in the article).
'----------------------------------------------------------------------------------------------
Sub RN24_Before()
    Dim o, s
    On Error Resume Next
    Set o = Server.LookupObject(99999999)
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
    Set o = Server.LookupObject(99999999)
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
