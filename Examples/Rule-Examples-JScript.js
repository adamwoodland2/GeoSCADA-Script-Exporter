// Geo SCADA 2025 R2 script runtime - before / after examples (JScript)
//
// Click the grey rectangle (pick action: Call Specific Function -> Run_Examples). Each issue is run twice:
//   Before - the code as commonly written for Geo SCADA 2025 and earlier
//   After  - the code that works in Geo SCADA 2025 R2 (and in earlier versions)
// On 2025 and earlier both steps should report OK. On 2025 R2 the Before step is expected to fail
// and the After step to work. Results are shown in a small form after each step (JScript has no
// MsgBox); an error in one step does not stop the run. A summary is shown at the end.
//
// Issue IDs match GeoSCADA-Script-Analyser.ps1 2.0.0 and later (JS = the JScript section of the Resource
// Centre (Knowledge Base) article "Adjustments for VBScript in Geo SCADA 2025 R2"; analyser 1.x used
// RN11 and RN12 for JS02 and JS03).
// The only database object used is the root group (Id 0), which always exists. Nothing is written.

var Results = "";

function Run_Examples() {
    Results = "";
    JS02_Before(); JS02_After();
    JS03_Before(); JS03_After();
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
// JS02 - Wrong case for a Geo SCADA name (Knowledge Base: Specific JScript Issues; release notes p24)
// The JScript runtime in 2025 R2 is case-sensitive for Geo SCADA objects, methods and properties.
// Before 2025 R2, names were matched regardless of case.
//----------------------------------------------------------------------------------------------
function JS02_Before() {
    try {
        Form.init("JS02");                                     // should be Form.Init
        var box = Form.AddEditBox(1, 1, 20, 1);
        box.value = "test";                                    // should be .Value
        var o = Server.LookUpObject(0);                        // should be LookupObject (0 = root group)
        Report("JS02", "Before", null,
            "Form.init / box.value / Server.LookUpObject\nExpected: OK before 2025 R2; error in 2025 R2.");
    } catch (e) {
        Report("JS02", "Before", e,
            "Form.init / box.value / Server.LookUpObject\nExpected: OK before 2025 R2; error in 2025 R2.");
    }
}

function JS02_After() {
    try {
        Form.Init("JS02");
        var box = Form.AddEditBox(1, 1, 20, 1);
        box.Value = "test";
        var o = Server.LookupObject(0);                        // 0 = root group
        Report("JS02", "After", null, "Form.Init / box.Value / Server.LookupObject\nExpected: OK in all versions.");
    } catch (e) {
        Report("JS02", "After", e, "Form.Init / box.Value / Server.LookupObject\nExpected: OK in all versions.");
    }
}

//----------------------------------------------------------------------------------------------
// JS03 - Number assigned to a Boolean property (Knowledge Base: Specific JScript Issues; release notes p24)
// The JScript runtime in 2025 R2 needs a real Boolean (true / false) for Boolean properties.
//----------------------------------------------------------------------------------------------
function JS03_Before() {
    try {
        Form.Init("JS03");
        var ok = Form.AddPushButton(5, 5, "OK");
        ok.Default = 1;
        Report("JS03", "Before", null, "ok.Default = 1\nExpected: OK before 2025 R2; error in 2025 R2.");
    } catch (e) {
        Report("JS03", "Before", e, "ok.Default = 1\nExpected: OK before 2025 R2; error in 2025 R2.");
    }
}

function JS03_After() {
    try {
        Form.Init("JS03");
        var ok = Form.AddPushButton(5, 5, "OK");
        ok.Default = true;
        Report("JS03", "After", null, "ok.Default = true\nExpected: OK in all versions.");
    } catch (e) {
        Report("JS03", "After", e, "ok.Default = true\nExpected: OK in all versions.");
    }
}
