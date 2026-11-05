# CLAUDE.md

Guidance for Claude Code working in this repository. `README.md` is the user-facing
documentation and the specification of the export file format. Read its "File structure" and
"Analyser" sections before changing either script.

## What this is

There are two independent Windows PowerShell scripts for Geo SCADA Expert **client-side
scripting**: mimic scripts and script libraries, in VBScript, JScript, JScript9 and JavaScript V8.
Logic programs are out of scope.

| File | Purpose |
|------|---------|
| `GeoSCADA-Script-Exporter.ps1` | Connects to a Geo SCADA server through the .NET Client API and writes one `.txt` file per mimic or script library. Each file has a parseable header and then the raw script. It also writes `_ScriptExport_Manifest.csv`. Read-only against the database. |
| `GeoSCADA-Script-Analyser.ps1` | Offline. Reads an export folder and reports what needs changing for Geo SCADA 2025 R2. Writes `_ScriptAnalysis.csv`, `_ScriptAnalysis_Actions.csv`, `_ScriptAnalysis_Instances.csv`, `_ScriptAnalysis_ExportCheck.csv` and a self-contained `_ScriptAnalysis.html`. No server connection. Exit code 3 when the export is incomplete. |
| `Examples/Rule-Examples-*.vbs / .js` | Before/after demo scripts for the version-dependent rules (RC, JS, RX), for a VBScript and a JScript test mimic. They may only use objects that always exist: the mimic running the script (`Server.ThisObject`) and the root group (Id 0). The one exception is RC01's `Test Analog` point, which the script looks for first and skips (after asking) when it's missing. |
| `Tests/` | Analyser regression tests. They use synthetic export files, so no server is needed. |
| `Tools/` (git-ignored) | Local test-system tools that **write** to the database, so they are kept out of the repository. `Build-ExampleMimics.ps1` builds the example mimics `Claude.Examples.Default` / `.JScript` by cloning `Claude.Test`'s document, so it can't run while `Claude.Test` is missing, and a rebuild loses buttons added by hand (such as RC12's). `Set-MimicScript.ps1` replaces the script of an existing `Claude.*` mimic and keeps everything else. Both are dry runs unless `-Write` is given. |
| `References/` (git-ignored) | Vendor documents: the 2025 R2 release notes PDF, printouts of the KB article, `ClientAPI.chm` and screenshots of field-found script issues. Not redistributable, so never commit them. |

The parameters and disclaimer are modelled on https://github.com/adamwoodland2/GeoSCADA-Mimic-Walker.
Remote: https://github.com/adamwoodland2/GeoSCADA-Script-Exporter (branch `main`).

## Environment

- **Runtime:** Windows PowerShell **5.1** on .NET Framework 4.8, not PowerShell 7. Both scripts must keep working in 5.1.
- **Geo SCADA:** 2025 R2 (build 88.9742.1). Client DLL: `C:\Program Files\Schneider Electric\ClearSCADA\ClearScada.Client.dll` (6.88.9742.1). Server: `127.0.0.1:5481`.
- **This machine is a copy of production.** The data is real and sensitive. The logon is in `CLAUDE.local.md`, which is not committed.
- **The server stops after a 2-hour runtime.** If a connection fails with the server not running, run `net start dbserver`. This is pre-approved.
- **Licence limit:** the database is capped at 1,000 objects, so new points can't be created here. The RC01 example needs a point called `Test Analog`, which has to be created on a licensed test system; without it the example asks whether to carry on without RC01.
- **Tools:** git and `gh` are installed; Python 3.14 is available but not used. There's no Node. For browser checks, use headless Edge at `C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`.

## Rules for working here

- **Never write to real database objects.** The only objects you may create or change are test objects under `Claude.*`, and only the ones the task is about. The group's contents change, so check what is there first (see Current state). Never point an example at a real tag.
- **The repository is read-only on the database.** Nothing committed may write to Geo SCADA: the exporter only reads, and the analyser and tests are offline. Anything that writes goes in the git-ignored `Tools/` folder.
- **Customer-specific work stays local.** Scripts written for this system live in its mimics. Notes about them go in git-ignored `*.local.md` files. Keep site, object and user names out of committed files.
- **Ask before running a full export.** It takes about 9 minutes and produces about 42,000 files; the user has stopped one before. Exporting single objects with `-Name` is fine.
- **Exported scripts and reports are sensitive.** Keep them out of git; `ScriptExport/`, `*.txt` and `*.csv` are ignored. Put working output in the session scratchpad or `%TEMP%`, never the repo.
- **Never commit credentials.** Test tools take `-Username` / `-Password` parameters.
- **Use the dedicated Edit tool for `README.md`.** A bulk regex replace once inserted a table row after every row in the file.
- **Commit and push only when the user asks.** Use the attribution lines from the system reminder.
- **Version every released change.**
  - Bump `$ExporterVersion` / `$AnalyserVersion` using semantic versioning: patch for fixes, minor for new rules / features / columns added at the end, major for breaking output or parameter changes.
  - Bump `FormatVersion` (and the analyser's `$SupportedFormat`) only when the export file layout changes.
  - Add an entry to `CHANGELOG.md` in the same commit.
- **ViewX COM automation** (`Serck.ViewXApplication.OpenFromServer`) hangs against the user's running ViewX. Don't use it, and don't try to run mimic scripts yourself. The user tests in ViewX.

## Common commands

```powershell
# Analyser and exporter-helper tests (no server needed). Exit code 0 = all passed.
.\Tests\Run-Tests.ps1

# Export one object or one group (fine to run); a full export needs the user's go-ahead
.\GeoSCADA-Script-Exporter.ps1 -OutputFolder <dir> -Username <u> -Password <p> -Name 'Claude.<Mimic>' -AcceptDisclaimer
.\GeoSCADA-Script-Exporter.ps1 -OutputFolder <dir> -Username <u> -Password <p> -Group 'Plant.Scripts' -AcceptDisclaimer

# Analyse an export folder (about 30 s for a full export of ~42k files)
.\GeoSCADA-Script-Analyser.ps1 -InputFolder <exportDir> -OutputFolder <reportDir>
#   -MinimumSeverity, -IncludeRule / -ExcludeRule RC01,GS03, -IncludeInstances, -SkipProgIdCheck, -NoHtml, -ManifestOnly
#   Exit code 3 = report written, but the export is incomplete (see _ScriptAnalysis_ExportCheck.csv)

# Local only (Tools/ is git-ignored): rebuild the example mimics after editing Examples\*.vbs / *.js.
# Needs Claude.Test; a rebuild loses hand-added buttons, so prefer Set-MimicScript for an existing mimic.
.\Tools\Build-ExampleMimics.ps1 -Username <u> -Password <p> -Write
```

Exporter parameters: `-OutputFolder -Server -Port -Username -Password -Group -Name -SkipMimics
-SkipLibraries -CheckMissingLibraries -ClientDllPath -FailOnError -AcceptDisclaimer`.

## Exporter: how it works

- **Object list:** `SELECT WITH TEMPLATES Id, FullName, ConfigVersion, ConfigTime, ConfigUser, ParentTemplateId, ParentInstanceId FROM CMimic` (and `CScriptLibrary`).
  - `WITH TEMPLATES` must come straight after `SELECT`. `FROM X WITH TEMPLATES` is a parse error.
  - Without it, objects inside template definitions are missing.
  - `InTemplate` is `ParentTemplateId > 0`.
- **Mimics:** `Server.GetWebDocumentContent` returns an int32 character count followed by UTF-16 XML, the same as the web `?content` XML. The script and pick actions are read from the XML:
  - **`<Script lang="{GUID}">`:** each `<Library name>` is preceded by a CRLF, with one more CRLF after the last. The mimic's own script is the text after the last `<Library>`, minus one CRLF.
  - **`<Action type>`:**
    - `scriptCall` calls a named function in the current mimic or its libraries.
    - `scriptCode` runs inline code in the current mimic.
    - `scriptTopLevel` runs inline code in the top-level mimic.
  - **`<EmbeddedMimic ref="SCX:////CMimic/.Name">`:** relative references, where each leading dot strips one level.
- **Libraries:** `InvokeMethod('GetScriptSource')` returns the code.
  - The library's language GUID comes from the binary document (`GetDocumentContent`).
  - That, and `-CheckMissingLibraries`, are the only places binary parsing is used. The user prefers standard API calls over parsing binary.
- **Language GUIDs:**
  - VBScript `{B54F3741-5B07-11CF-A4B0-00AA004A55E8}`, which is also the default when there is no `<Script>`.
  - JScript `{F414C260-6AC0-11CF-B6D1-00AA00BBBB58}`.
  - JScript9 `{16D51579-A30B-4C8B-A276-0FF4DC41E755}`.
  - JavaScriptV8 `{16BA25BB-FB03-48BA-AECA-C53BF6ABD2A1}`.
- **Template control:**
  - `EditableSource` is the object a person actually edits: an instance copy follows `GetTemplateObject` (cached) back to its template definition, unless the copy has its own document (`HasDocumentContentOverride`).
  - Also recorded: `IsControlledByInstance`.
  - Example: `Region.Site A.Pumps.Pump 1.Display.Tab Status` maps to `~Templates.Pump Station.Devices.Pump.Display.Tab Status`.
- **File format:** the header lines are comments (`'` or `//`) between `===== GEOSCADA-SCRIPT-EXPORT BEGIN/END =====` markers, followed by the raw script.
  - Current version: `FormatVersion: 2` (`ExporterVersion` 1.1.1; analyser 2.1.0). See `CHANGELOG.md`.
  - Format 2 adds `PICKACTION <n>` code sections for inline pick-action code, `PickActionCode`, `ScriptStartLine` / `ScriptEndLine`, 7-field `PickAction` lines (`N | Where | Action | Type | Context | Function | Args`) and `EmbeddedMimic`.
  - The format is a parsing contract. If it changes, bump `FormatVersion`, update the README spec, and keep the analyser reading every older format. `Split-PickAction` handles both 1 and 2.
- **File names:** the object's full name plus `.txt`. Invalid characters become `_`, and Windows device names (`CON`, `CON.Child`, `LPT1.x` …) get a leading `_`. Collisions get ` [Id]`. Paths of 248 characters or more use the `\\?\` prefix.
- **Manifest and CSV cells:** a text cell starting with `= + - @` (or a tab or CR) is written with a leading `'`, to stop spreadsheet formula injection. Both scripts have a `ConvertTo-SafeCsvRow` helper. The analyser undoes the `'` when it reads the manifest (`ConvertFrom-SafeCsvValue`), so keep the two in step.
- **Line numbering:** a script ending in a line break has an empty last line, which `ScriptEndLine` counts, as an editor does.
- **Progress:** `Write-Progress` costs about 2 s per call in this host. Always go through the throttled `Show-Progress` helper.
- **Failures:** everything per object (read, template lookup, header, write) sits inside a try that calls `Register-Failure`, so one bad object becomes a manifest `Error` row (exit 2), never the end of the run.
- **Arrays through `-File`:** `powershell -File … -Group 'A','B'` arrives as one string `A,B` and matches nothing. Names can contain commas, so the exporter does not split them; call it with `-Command "& '.\GeoSCADA-Script-Exporter.ps1' -Group 'A','B' …"` instead.

## Analyser: how it works

- **`$Rules`:** an ordered table near the top holding ID, Severity, Title, Fix and Ref, followed by the help-text rule list. All output derives from this table.
- **Export check:** this runs first, in the main loading loop. It compares the folder with the manifest and validates each opened header with `Get-HeaderProblem`, then writes `_ScriptAnalysis_ExportCheck.csv` and the red HTML note.
  - Issues: `MissingFile`, `ExportError`, `Malformed`, `DuplicateName`, `ManifestMismatch`, `UnsupportedFormat`, `NotInManifest` and `NoManifest`.
  - The first three mean the report is incomplete, so the analyser exits with 3.
  - Files not in the manifest are still analysed unless `-ManifestOnly` is given, because two scoped exports into one folder are legitimate.
  - The manifest is authoritative for names: exporter 1.0.0 trimmed trailing spaces in headers.
  - `$SupportedFormat` is the newest `FormatVersion` the analyser understands. Bump it when the exporter's format changes.
- **Header rules** use the exporter metadata (the GS rules: libraries, pick-action resolution, unused libraries).
- **Body rules** live in `Invoke-BodyRules`.
  - Each script or pick-action code section is scanned through `New-ScanContext`.
  - `Code` is the script with comments **and** string contents blanked to spaces; `NoComments` blanks comments only. Both keep their length and line breaks, so positions map straight back to the original.
  - In JS, a regex literal starts at `/` after punctuation, an operator, or a keyword such as `return`, `typeof` or `case` (`RegexKeywords`); after an identifier or a number, `/` is division.
  - A compiled C# helper, `GeoScadaScriptScan`, does the work: `ReadHeader`, `ReadBody`, `MaskVbs`, `MaskJs`, `MatchParen`, `CountArgs`, `Chains`. Its source is in a **single-quoted** `@'…'@` here-string, because a double-quoted one mangles backticks.
  - Shared helpers:
    - `Test-StatementCall`: is a call the whole statement, or a read inside an expression?
    - `Get-VbsDeclarations`: VBScript procedures, their parameters, and `Dim`/`Const` at procedure or script level.
- **Function resolution (GS03):** pick-action functions are looked up in the mimic's own script, its libraries and embedded mimics (`Get-MimicDefs` / `Get-LibraryDefs`), honouring the action's Current or TopLevel context.
  - `scriptCode` / `scriptTopLevel` actions are inline code, not function names.
  - JS definitions use ordinal (case-sensitive) dictionaries.
- **Batching:**
  - Findings are reported against `EditableSource`, with a role:
    - `Template`
    - `OverriddenInstance` (edit it individually)
    - `Instance`
    - `Standalone`
  - Each finding also carries `InstancesFixed`.
  - Instance copies are skipped unless `-IncludeInstances` is given; the manifest pre-skip records which instances each source covers.
- **HTML:** one template in a single-quoted here-string with a `{{JSON}}` placeholder; the data is serialised with `JavaScriptSerializer`.
  - Cast values explicitly to `[string]`, `[int]` or `[string[]]`, or the serialiser hits circular references.
  - The UI provides:
    - a by-rule table listing every rule in `$Rules`, in rule order, sortable by any column. A rule with no findings shows `0`; a rule that didn't run shows `—` and the reason (`x` in the rules JSON);
    - a multi-select rule filter with All / None / RC only / JS only / RX only / GS only shortcuts, whose unticked rules are saved in `localStorage` (key `…rulesOff.v2`). When that key is missing, the 1.x key (`…rulesOff`) is read once and its RN IDs translated through `OLD` (built from each rule's `o`, its 1.x ID, and `p`, its retired IDs). IDs read from the v2 key go through `OLD` too, so an unticked RX10 becomes RC16;
    - each rule's earlier IDs ("was RX10, RN21", from `was(k)`) in the by-rule table and the filter's tooltip; search also matches them;
    - rule order from `ruleCmp`: prefix order RC, JS, RX, GS (`PREFIX`), then a numeric compare, so RC06a < RC06b < RC07;
    - severity and role filters: multi-select drop-downs like the rule filter, built by `makePick` (All / None, a count per choice, not remembered between reports); a language `<select>`; and search. Every `details.pick` closes on an outside click or Escape and is shifted back inside the window when it opens.

### Rule conventions

- **Rule IDs (since 2.0.0) follow the Resource Centre (Knowledge Base) article** "Adjustments for VBScript in Geo SCADA 2025 R2" (`$KB` in the analyser, which records the date it was last checked; the latest printout is `References/ResCentre 2026-11-05.pdf`, items 1–16 and J1–J4). The article is updated over time, so re-check it for new items.
  - **RC*nn*:** article item *nn* (RC = Resource Centre), letter-suffixed when one item has several checks (RC06a row count / RC06b scrolling; RC14a known aggregate / RC14b every other GetOPCValue). Item 5 (form layout) has no rule. A new article item gets the next RC number.
  - **JS*nn*:** the article's "Specific JScript Issues" section, now numbered J1–J4 there (JS*nn* = J*n*): JS01 ScriptControl (also checked in VBScript), JS02 case, JS03 Booleans, JS04 pick-action arguments; JS05 JScript in use is the section's closing note.
  - **RX*nn*:** 2025 R2 changes not in the article: RX01–RX03 64-bit COM, RX04 debugger, RX05–RX09 V8. When the article adds one, move it to its RC number and **retire** the RX number: RX10 became RC16 in 2.1.0. The user treats a single move like this as **minor**, not breaking: no column changes, and the retired ID is still translated everywhere.
  - **GS*nn*:** exporter-metadata checks (unchanged since 1.x).
  - **Never reuse an ID.** The 1.x IDs **RN01–RN28** are retired; the user chose new prefixes so that no old ID means something new. `$OldRuleIds` (after `$Rules`) maps each 1.x ID straight to its **current** ID (RN21 → RC16, not via RX10), so translations never chain. `$RetiredRuleIds` maps IDs retired since 2.0.0 (RX10 → RC16). `-IncludeRule` / `-ExcludeRule` translate both (`Resolve-RuleIds`, with a console note), and the HTML shows both (`o` = 1.x ID, `p` = retired IDs). `Run-Tests.ps1` checks the 1.x map covers RN01–RN28 once each, points at real rules, and agrees with the README "1.x rule" column and the CHANGELOG 2.0.0 table (whose RX10 resolves through `$RetiredRuleIds`); and that retired IDs are not current IDs, point at real rules, and appear in the README ("was RX10") and a CHANGELOG table.
  - **Moving an RX rule to RC:** rename it in `$Rules`, `Invoke-BodyRules`, the help text, README, fixtures (tags) and examples; add it to `$RetiredRuleIds`; point its `$OldRuleIds` entry at the new ID; add a retired → now table to the CHANGELOG.
  - A wholesale renumbering (like 2.0.0's) is breaking (major version): the `Rule` CSV column changes throughout. Moving one RX rule to RC with its old ID retired and translated is minor (see RX above). Use new IDs, extend the old → new map, put a table in `CHANGELOG.md`, and do the rename in one pass with all IDs mapped at once, so renames can't chain. The 2.0.0 rename was checked by running 1.4.0 and 2.0.0 over both exports: every finding matched line-for-line under its new ID.
  - Refs cite the item number (`$KB #6 '…'`) plus the release-notes page where there is one.
  - **Reading the KB printouts:** they come from Microsoft Print to PDF, which draws most text as shapes, so text extraction gets almost nothing. Render the pages with the Windows `Windows.Data.Pdf` API from PowerShell 5.1 (`PdfDocument.LoadFromFileAsync` → `GetPage(i).RenderToStreamAsync` → PNG) and read the images.
  - **RC01:** flags `.Value(x)` only when the call is the whole statement (a set). Reads such as `y = (x.Interface.Value(0))` work.
  - **RC03:** flags any bare use of a variable `Set` from `Form.Add…` in the same procedure. Exempt:
    - `.Member`
    - `Set y = x`
    - `x Is Nothing`
    - `IsObject` / `TypeName`
    - passing `x` to the user's own sub or function
  - **RC10:** an **undeclared** VBScript variable named `ToString`, `GetType`, `GetHashCode` or `Equals` is an Error. The R2 VBScript runtime is .NET-based, so these resolve to .NET members.
    - `Dim` in the same procedure is OK (tested, reported as Info).
    - A script-level `Dim`, a parameter, or a Sub/Function of that name is untested, so a Warning.
  - **RC04:** a variable `Set` from `Server.FindObject` / `LookupObject` / `ThisObject` (optionally `.Parent`) used as a value. It shares `Get-SetVariables` and `Find-ValueUses` with RC03, so keep their exemptions in step.
  - **RC06b:** every `Form.AddListBox` is a Review: in R2 a list box with more items than rows scrolls horizontally. Contents are usually dynamic, so no attempt is made to count items. It's raised in the RC06a loop, so it always matches RC06a's list boxes. Treated as an expected R2 change, so the wording is neutral (no "bug" note).
  - **RC09:** `IsNull` / `IsEmpty` / `= Null` / `IsObject` used as a not-found test on a `FindObject`/`LookupObject` result.
    - Not an R2 change: the test never worked.
    - Skipped when the script or a linked library defines its own `isNull` etc. Real systems do this: on the test system, one library's own `isNull` would otherwise have given 28 false hits.
  - **RC11:** a date passed through `SetRegistry` / `GetRegistry` without `CLng(Int(CDbl(…)))` / `CDate`.
    - `$isDate` only matches a value that starts with a date, so `"Date: " & Now` (a string) is not flagged.
  - **RC12:** a `scriptCall` pick action that gives no value to some of its function's parameters. In the R2 release (88.9742.1) the function silently doesn't run. The KB says a future update will fix it, so the wording says it is a problem in the release version and to check the current release notes. Keep it an Error until a fixed version ships.
    - It lives in the pick-action loop, next to GS03, and uses the parameter names now stored by `Get-FunctionDefs` (name → `string[]`).
    - A parameter counts as given only if it appears as `name <= value` in the PickAction `Args`. In the XML, a blank parameter is either missing or an `<Arg>` with no `<Animation>`; the exporter writes the latter as a bare name.
    - Header finding, so it's tested per button in `Run-Tests.ps1`, not by tags.
  - **RC13:** `App.LogOn` / `LogOff` / `LogOnUser` / `LogOffUser` (all languages; the method name in any case, because the native `ViewXCore.dll` type library spells them `Logon` / `Logoff` while the article writes `LogOn`; in JS the object must be exactly `App`, since `app` is a script variable, and a wrong method case is also JS02's), and `.Logon` on a variable set from `CreateObject("Serck.ViewXApplication")` or chained on it. In R2 the script stops at the call.
    - `Test-CodeFollows` decides the severity: Warning when code follows in the same procedure, Info when the call is the last statement. In JS the procedure is the innermost enclosing `function … {` / `=> {` (`MatchBrace`).
    - Deliberately conservative: only whitespace, block closers (`End If`, `Next`, `Loop`, `}` …) and `Exit` / `return` / `break` are ignored, so an `Else` branch counts as code.
    - `Form.Logon` is not covered (not in the article). Severity is tested per procedure in `Run-Tests.ps1`.
  - **RC14a / RC14b:** `Server.GetOPCValue` can't read aggregates in the R2 release (a temporary limitation; wording says to check the current release notes). A tag string can't tell `Object.Property` from `Object.Aggregate.Property`, so RC14a (Warning) only fires when a string literal in the argument names one of `$AggregateNames` after a dot. RC14b (Info, the user's choice) lists every other call.
    - `$AggregateNames` is the 110 aggregate names from the 2025 R2 schema, read through `Connection.GetClass('CDBObject')` → `DerivedClasses` / `Aggregates` / `AggregateClasses` (read-only). `$AggregatesAlsoProperties` (AlarmSummary, PSTN, Shared) are also property names, so they need a following name. Re-read the schema on a new version.
    - On the two exports, RC14a found 2 per export (both real `.Control.Action1Desc` reads); RC14b about 1,600 / 400.
    - `GISLocation` exists on every `CDBObject` (it can't be disabled), so examples can use `Server.ThisObject.FullName & ".GISLocation.Latitude"`.
  - **RC15** (article #15): VBScript `Set y = … .Interface.<aggregate>` (statement ends at the aggregate name). Error when `y Is Nothing` follows in the same procedure (the article's pattern); Warning otherwise. Skips `GISLocation` and `$AggregatesAlsoProperties`. On the 2026-11-05 export: 1 Error, 12 Warnings, all `.Interface.Historic`.
  - **RC16** (article #16; RX10 in 2.x): `GetRegistry` without a default. Not reported for the article's fix: VBScript `a = GetRegistry("k")` with `IsNull(a)` later in the same procedure.
  - **RC12 parameters:** `Get-FunctionDefs` lets an empty `()` sit inside a VBScript parameter list, and `Get-ParamNames` strips it, so `Sub F(vals(), n)` gives `vals`, `n`.
- **GS01–GS08:** best-practice and consistency checks from exporter metadata. Some will deliberately never be fixed, which is why the HTML filter remembers unticked rules.
- **A new rule needs:**
  - an entry in `$Rules`;
  - the help-text list near the top of the analyser;
  - the README rules table, with `*new*` in its "1.x rule" column;
  - `EXPECT:` / `NOT:` tagged lines in `Tests/New-Fixtures*.ps1`, including the false-positive cases;
  - for a version-dependent rule, a Before/After pair in `Examples/Rule-Examples-*` (plus a rebuild).
- **False positives matter.** After changing a rule, re-analyse a full export and diff that rule's findings against the previous report.
- **Field facts behind the rules:**
  - Script error positions in ViewX are **0-based** (line, column).
  - A library runs inside the mimic, so its errors are reported under the mimic's name: the line number counts through the combined text, and the caption shows `[unknown]` rather than the library.

## Testing

- **`Tests\Run-Tests.ps1`:** must pass before every commit. It builds fixtures in `%TEMP%\GeoSCADA-Script-Analyser-Tests` and runs five groups of checks:
  - every `EXPECT:` / `NOT:` tag, by file line (currently 199 tags), plus per-case checks for the RC12 buttons, RC13 severities and RC15 severities;
  - template batching and the export check, with and without a manifest;
  - `-ManifestOnly`;
  - the exporter's helper functions;
  - old rule IDs: the `$OldRuleIds` and `$RetiredRuleIds` maps (read from the analyser's syntax tree) against `$Rules`, the README and the CHANGELOG, and `-IncludeRule` / `-ExcludeRule` with a 1.x ID and a retired ID.

  The fixture and test scripts:
  - `New-Fixtures.ps1`: format 1, with the rule tags.
  - `New-FixturesV2.ps1`: format 2, with code sections and contexts.
  - `New-FixturesBatching.ps1`: a template, instances and an overridden instance; names starting with `-` (the CSV guard); and one case per export-check issue.
  - `Test-Exporter.ps1`: pulls exporter functions out of its syntax tree, without running it or connecting. It tests file names, the file layout and line numbers, relative names and the CSV guard.
  - Not covered by any test: XML extraction from live documents, template resolution, binary document parsing and failure handling. These need the server.
- **Proving a fixture catches a bug:** the C# types are compiled once per process. To compare old and new behaviour, run each version in a fresh `powershell -NoProfile -File …` process.
- **HTML changes:** write a copy of the report with an extra `<script>` that clicks controls and writes results into a `<pre>`. Run it with `msedge --headless=new --virtual-time-budget=5000 --dump-dom file:///…`, and use `--screenshot` to check the layout. Calling `msedge` directly from the PowerShell tool can print nothing; `Start-Process` with `--no-sandbox`, a fresh `--user-data-dir` and `-RedirectStandardOutput` works. To test saved filter choices, inject a `<script>` right after `<body>` that sets the `localStorage` key first.
- **Exporter changes:** export a few `Claude.*` objects and one real group, then check them. A full export needs the user's go-ahead first.

## PowerShell 5.1 pitfalls hit in this code

- **Single-element arrays unroll on return.** Wrap the call: `@(Func …)`.
- **Hashtables are case-insensitive.** Use `[Hashtable]::new([StringComparer]::Ordinal)` when case matters, e.g. for JS names.
- **`@($List)` on a generic `List[T]` can throw "Argument types do not match".** Use `.Count`.
- **CRLF text:** in regex, `$` doesn't match before `\r`, so use `\r?$`. Use `[ \t]` rather than `\s` when a match must not cross lines.
- **`ConvertTo-HeaderValue` doesn't trim on purpose.** Leading and trailing spaces in names are significant.
- **Hidden native errors:** avoid `2>&1` on native executables. Use try/catch with `-ErrorAction Stop` where a cmdlet failure should be caught.

## Current state (October 2026)

- **Pushed** to `main` (history rewritten on 2026-10-15 to remove customer object names; keep examples generic, e.g. `Plant.*`, `Region.Site A.*`). Asking GitHub Support to purge the two old commit IDs is left to the user. Commit only when asked.
- **GS03 is provisional** until a full format-2 export has been reviewed. A format-2 export now exists (a git-ignored `ScriptExport_*` folder in the repo, 44,598 files, exporter 1.1.0, made by the user on 2026-11-05) but its GS03 findings have not been reviewed yet. Ask before running a new full export.
- **New machine (2026-11-04):** the repo was re-cloned on a different server; `Tools/` and `CLAUDE.local.md` stayed on the old one until the user copies them over. `gh` is not installed here.
- **Unconfirmed on R2 (2.1.0):** RC15's Warning case (a `Set` of an aggregate with no `Is Nothing` test) and the RC15 example pair are untested in ViewX.
- **Test mimics gone (checked 2026-10-29):** `Claude.Test`, `Claude.Default`, `Claude.ArrayTest`, `Claude.PerformanceTest` and `Claude.Examples.*` no longer exist on this server; the `Claude` group now holds other mimics. Ask before recreating any of them.
- **Unconfirmed on R2:** the example mimics haven't been run in ViewX. Their "(to confirm)" expectations, and RC10's script-level `Dim` and parameter cases, are untested.
