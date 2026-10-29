// Geo SCADA 2025 R2 script runtime - before / after examples (JScript)
//
// Click the grey rectangle (pick action: Call Specific Function -> Run_Examples). Each issue is run twice:
//   Before - the code as commonly written for Geo SCADA 2025 and earlier
//   After  - the code that works in Geo SCADA 2025 R2 (and in earlier versions)
// On 2025 and earlier both steps should report OK. On 2025 R2 the Before step is expected to fail
// and the After step to work. Results are shown in a small form after each step (JScript has no
// MsgBox); an error in one step does not stop the run. A summary is shown at the end.
//
// Issue numbers match GeoSCADA-Script-Analyser.ps1 (RN = 2025 R2 script runtime changes).
// The only database object used is the root group (Id 0), which always exists. Nothing is written.

var Results = "";

function Run_Examples() {
    Results = "";
    RN11_Before(); RN11_After();
    RN12_Before(); RN12_After();
    ShowText("Geo SCADA 2025 R2 script examples - summary", Results);
}

// Shows text in a form, one static text line per line of text.
function ShowText(title, text) {
    var lines = text.split("\n");
    Form.Init(title);
    for (var i = 0; i < lines.length; i++) { Form.AddStaticText(1, i, lines[i]); }
    var ok = Form.AddPushButton(1, lines.length + 1, "OK");
    ok.Default = true;
    Form.Show();
}

function Report(id, stage, e, expectation) {
    var msg = id + " " + stage + ": " + (e ? "FAILED - " + (e.message || e.description || e) : "OK");
    Results += msg + "\n";
    ShowText("Geo SCADA 2025 R2 script examples", msg + "\n\n" + expectation);
}

//----------------------------------------------------------------------------------------------
// RN11 - Wrong case for a Geo SCADA name (release notes p24)
// The JScript runtime in 2025 R2 is case-sensitive for Geo SCADA objects, methods and properties.
// Before 2025 R2, names were matched regardless of case.
//----------------------------------------------------------------------------------------------
function RN11_Before() {
    try {
        Form.init("RN11");                                     // should be Form.Init
        var box = Form.AddEditBox(1, 1, 20, 1);
        box.value = "test";                                    // should be .Value
        var o = Server.LookUpObject(0);                        // should be LookupObject (0 = root group)
        Report("RN11", "Before", null,
            "Form.init / box.value / Server.LookUpObject\nExpected: OK before 2025 R2; error in 2025 R2.");
    } catch (e) {
        Report("RN11", "Before", e,
            "Form.init / box.value / Server.LookUpObject\nExpected: OK before 2025 R2; error in 2025 R2.");
    }
}

function RN11_After() {
    try {
        Form.Init("RN11");
        var box = Form.AddEditBox(1, 1, 20, 1);
        box.Value = "test";
        var o = Server.LookupObject(0);                        // 0 = root group
        Report("RN11", "After", null, "Form.Init / box.Value / Server.LookupObject\nExpected: OK in all versions.");
    } catch (e) {
        Report("RN11", "After", e, "Form.Init / box.Value / Server.LookupObject\nExpected: OK in all versions.");
    }
}

//----------------------------------------------------------------------------------------------
// RN12 - Number assigned to a Boolean property (release notes p24)
// The JScript runtime in 2025 R2 needs a real Boolean (true / false) for Boolean properties.
//----------------------------------------------------------------------------------------------
function RN12_Before() {
    try {
        Form.Init("RN12");
        var ok = Form.AddPushButton(5, 5, "OK");
        ok.Default = 1;
        Report("RN12", "Before", null, "ok.Default = 1\nExpected: OK before 2025 R2; error in 2025 R2.");
    } catch (e) {
        Report("RN12", "Before", e, "ok.Default = 1\nExpected: OK before 2025 R2; error in 2025 R2.");
    }
}

function RN12_After() {
    try {
        Form.Init("RN12");
        var ok = Form.AddPushButton(5, 5, "OK");
        ok.Default = true;
        Report("RN12", "After", null, "ok.Default = true\nExpected: OK in all versions.");
    } catch (e) {
        Report("RN12", "After", e, "ok.Default = true\nExpected: OK in all versions.");
    }
}
