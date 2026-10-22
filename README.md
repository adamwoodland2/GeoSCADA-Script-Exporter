# GeoSCADA-Script-Exporter

Exports every **client-side script** in a **Geo SCADA** database to text files: mimic scripts
and Script Library code, in **VBScript, JScript, JScript9 and JavaScript V8** (new in Geo SCADA
2025 R2). Each file is named after the object's full name and starts with a structured header,
so a later step can parse the files to look for problems.

It uses the Geo SCADA **.NET Client API** (`ClearScada.Client.dll`) and only reads; nothing in
the database is changed. Logic programs are out of scope.

The repository has two independent scripts:

| Script | What it does |
| ------ | ------------ |
| `GeoSCADA-Script-Exporter.ps1` | Connects to Geo SCADA and writes one `.txt` per script (this page). |
| `GeoSCADA-Script-Analyser.ps1` | Works offline on those files and reports what needs updating, chiefly for the 2025 R2 script engine change. See [Analyser](#analyser). |

A companion to [GeoSCADA-Mimic-Walker](https://github.com/adamwoodland2/GeoSCADA-Mimic-Walker).

## How it works

1. **Connects and logs on** through the .NET Client API (`ClearScada.Client.Simple.Connection`)
   and lists the `CMimic` and `CScriptLibrary` objects with SQL. The query uses
   `SELECT WITH TEMPLATES`, so objects inside template definitions are included; a plain `SELECT`
   leaves them out. Template mimics, which are the ones you edit, are therefore exported as well.
2. **Mimics:** `IObjectConfiguration.GetWebDocumentContent()` returns the mimic definition as
   XML (the same XML the web server serves for `/db/<Mimic>?content`). From it the script reads:
   - `<Script lang="{GUID}">` for the script language;
   - the text of `<Script>`, which is the mimic's own code;
   - `<Library name="...">` for each script library the mimic links to;
   - each pick action that runs script. There are three kinds:

     | XML `type` | ViewX pick action mode | What it holds | Runs in |
     | ---------- | ---------------------- | ------------- | ------- |
     | `scriptCall` | Call Specific Function | A function name (plus arguments) | This mimic: its script and linked libraries |
     | `scriptCode` | Execute Code Fragment | Inline script code | This mimic |
     | `scriptTopLevel` | Execute Code Fragment on Top Level Mimic | Inline script code | The top-level mimic on display |

     Inline code is scripting in its own right, so it is written out in full in a section of its
     own after the mimic script (see [File structure](#file-structure-formatversion-2));
   - `<EmbeddedMimic ref="...">` for each embedded mimic, with relative references such as
     `SCX:////CMimic/.Test` resolved to a full name.

   A file is written if the mimic has its own code, links to a library, or has a script pick
   action. That covers mimics with no script at all. A function call from such a mimic is an issue
   in itself; inline code is not.
3. **Script libraries:** the library's own server method `GetScriptSource()` returns the code.
   No API call returns a library's language, so it is read from the language GUID at a fixed
   position in the library's stored record (`GetDocumentContent()`).
4. **Finds where the script is edited.** An object inside an instance usually cannot be edited:
   its content comes from the matching object in the instance's template. The script follows
   `IObjectHierarchy.GetTemplateObject` from template to template until it reaches an object that
   is not controlled by an instance, or one whose content is overridden locally
   (`HasDocumentContentOverride`). That object is recorded as `EditableSource`. The chain can have
   several hops when a template contains instances of other templates. For example:

   ```
   FullName:       Region.Site A.Pumps.Pump 1.Display.Tab Status
   EditableSource: ~Templates.Pump Station.Devices.Pump.Display.Tab Status
   ```
5. **Writes** one `.txt` per object and a manifest, `_ScriptExport_Manifest.csv`.

### Languages

| Language | GUID |
| -------- | ---- |
| VBScript (also the default when none is stored) | `{B54F3741-5B07-11CF-A4B0-00AA004A55E8}` |
| JScript | `{F414C260-6AC0-11CF-B6D1-00AA00BBBB58}` |
| JScript9 | `{16D51579-A30B-4C8B-A276-0FF4DC41E755}` |
| JavaScriptV8 | `{16BA25BB-FB03-48BA-AECA-C53BF6ABD2A1}` |

Any other GUID is exported as `Language: Unknown`, with the GUID recorded.

### Missing script libraries (`-CheckMissingLibraries`)

The mimic XML leaves out library links whose library **does not exist**, for example after a
library was renamed or moved. With `-CheckMissingLibraries` the mimic's stored document
(`GetDocumentContent()`) is also decoded to read the complete list of links. Any link the XML
dropped is reported as `LibraryMissing`. The stored script text is also compared with the XML
text as a cross-check, and any difference is reported in the manifest.

This relies on an undocumented binary layout, so it is **off by default**. If decoding fails for a
mimic, that mimic is still exported from the XML and its header says the check was incomplete.

## Requirements

- Windows PowerShell 5.1. `ClearScada.Client.dll` targets .NET Framework 4.8.
- The Geo SCADA client or server installed on the machine; the script finds `ClearScada.Client.dll` itself.
- A Geo SCADA user with permission to view the objects you want to export.
- Tested against Geo SCADA 2025 R2 (`ClearScada.Client.dll` 6.88.9742.1).

## Usage

```powershell
# Everything on the local server, prompting for credentials
.\GeoSCADA-Script-Exporter.ps1 -OutputFolder C:\Export\Scripts

# A remote server, one part of the database, with the missing-library check
.\GeoSCADA-Script-Exporter.ps1 -Server scada01 -Credential (Get-Credential) `
    -Group 'Sewage.Site' -OutputFolder C:\Export\Scripts -CheckMissingLibraries

# Non-interactive, script libraries only
.\GeoSCADA-Script-Exporter.ps1 -OutputFolder C:\Export\Scripts -Username operator -Password 'secret' `
    -SkipMimics -AcceptDisclaimer
```

### Parameters

| Parameter                | Default                        | Description |
| ------------------------ | ------------------------------ | ----------- |
| `-OutputFolder`          | `ScriptExport` (by the script) | Folder for the `.txt` files and the manifest. Created if missing. |
| `-Server`                | `127.0.0.1`                    | Geo SCADA server host name or address. |
| `-Port`                  | `5481`                         | Geo SCADA client port. |
| `-Credential`            |                                | PSCredential for the logon. |
| `-Username`/`-Password`  |                                | Plain-text alternative to `-Credential`. If neither is given you are prompted. |
| `-Group`                 |                                | Only objects at or below these groups. Can be combined with `-Name`. |
| `-Name`                  |                                | Only these objects (full names). |
| `-SkipMimics`            | *(off)*                        | Do not export mimic scripts. |
| `-SkipLibraries`         | *(off)*                        | Do not export script libraries. |
| `-CheckMissingLibraries` | *(off)*                        | Decode stored mimic documents to report links to libraries that do not exist. |
| `-ClientDllPath`         | *(auto)*                       | Path to `ClearScada.Client.dll`. |
| `-FailOnError`           | *(off)*                        | Stop at the first object that cannot be exported; otherwise finish and exit 2. |
| `-AcceptDisclaimer`      | *(off)*                        | Accept the disclaimer non-interactively. |

## Output format

### File names

A file is named `<FullName>.txt`, for example `Sewage.Site.Display.Overview.txt`.

- Characters Windows does not allow in file names (`\ / : * ? " < > |` and control characters) become `_`.
- Trailing dots and spaces are removed.
- Windows device names get a leading `_`, also when followed by a dot or spaces: `CON`, `CON.Child`, `LPT1.Script` and `NUL .x` become `_CON.txt`, `_CON.Child.txt` and so on. The device names are `CON`, `PRN`, `AUX`, `NUL`, `COM0`–`COM9`, `LPT0`–`LPT9`, and `COM`/`LPT` followed by ¹, ² or ³.
- If two objects would end up with the same file name, the second gets ` [<Id>]` appended.

The real full name is always in the header and the manifest.

Paths of 248 characters or more are written with the `\\?\` prefix. Use a short `-OutputFolder`
if your tools cannot open long paths.

### File structure (FormatVersion 2)

Files are UTF-8 without a BOM, with CRLF line endings. A file has these parts:

1. a `BEGIN` marker line;
2. header lines, each `Key: Value`;
3. an `END` marker line, followed immediately by the script exactly as stored (lines
   `ScriptStartLine` to `ScriptEndLine`);
4. for each pick action holding inline code, a `PICKACTION <n>` marker line followed by the code,
   on the lines given by that action's `PickActionCode` header entry.

Script line endings are normalised to CRLF; nothing else is changed. Every header and marker
line starts with the language's line-comment token: `'` for VBScript and for mimics with no script
(`Language: None`), and `//` for the JavaScript family and unknown languages. That keeps the whole
file valid for the language's parser or linter.

```
' ===== GEOSCADA-SCRIPT-EXPORT BEGIN =====
' FormatVersion: 2
' FullName: Plant.Tools.Audit.Audit Report
' Class: CMimic
' ObjectId: 12345
' Language: VBScript
' LanguageGuid: {B54F3741-5B07-11CF-A4B0-00AA004A55E8}
' ConfigVersion: 3
' ConfigTime: 2017-11-13T23:42:47Z
' ConfigUser: someuser
' Source: GetWebDocumentContent
' Server: 127.0.0.1
' ExportedAt: 2026-10-08T01:42:55Z
' Exporter: GeoSCADA-Script-Exporter 1.1.0
' InTemplate: False
' ControlledByInstance: False
' DocumentContentOverride: False
' EditableSource: Plant.Tools.Audit.Audit Report
' EditableSourceId: 12345
' EditInPlace: True
' HasOwnScript: True
' ScriptLines: 291
' ScriptChars: 13073
' LibraryCount: 2
' Library: Plant.Scripts.Common
' Library: Plant.Scripts.Navigation
' LibraryCheck: Performed
' LibraryMissing: Plant.Scripts.Old Reports
' PickActionCount: 2
' PickAction: 1 | New Layer[Layer]/Btn_3[Button] | New Action | scriptCall | Current | RunAudit | strPath <= ".Source Path.Value", strOutputFile <= ".Output File Name.Value"
' PickAction: 2 | New Layer[Layer]/Btn_4[Button] | New Action | scriptTopLevel | TopLevel | (code) | (none)
' EmbeddedMimicCount: 1
' EmbeddedMimic: Header[Layer]/M_2[EmbeddedMimic] | Plant.Symbols.Header
' PickActionCode: 2 | 330 | 331
' ScriptStartLine: 38
' ScriptEndLine: 328
' ===== GEOSCADA-SCRIPT-EXPORT END =====
'***********************************************************************************************
...script...
' ===== GEOSCADA-SCRIPT-EXPORT PICKACTION 2 =====
strParent = Server.ThisObject.Parent.FullName
VerifyControl strParent
```

**Parsing rules**

- The `BEGIN` marker is line 1. The header ends at the line containing `GEOSCADA-SCRIPT-EXPORT END`.
- For a header line, remove the comment token and one space, then split on the first `": "` to get the key and value.
- Values are single-line. Keys may repeat (`Library`, `LibraryMissing`, `PickAction`, `PickActionCode`, `EmbeddedMimic`).
- The script is file lines `ScriptStartLine` to `ScriptEndLine`. Script line *n* is at file line `ScriptStartLine + n - 1`, which lets issues found in the file be mapped back to the script. An empty script has `ScriptEndLine` = `ScriptStartLine - 1`.
- `PickActionCode: <n> | <first line> | <last line>` gives the file lines of pick action *n*'s inline code. Its marker line is the line before.
- Format 1 files, from exporter 1.0, have no `ScriptEndLine`, `PickActionCode` or `EmbeddedMimic` entries, and use five `PickAction` fields. The script runs to the end of the file. The analyser reads both formats.

**Header keys**

| Key | Applies to | Meaning |
| --- | ---------- | ------- |
| `FormatVersion` | all | `2`. Bumped if the layout changes. |
| `FullName`, `Class`, `ObjectId` | all | The object: `CMimic` or `CScriptLibrary`. |
| `Language`, `LanguageGuid` | all | `VBScript`, `JScript`, `JScript9`, `JavaScriptV8` or `Unknown`, plus the GUID. A mimic with no `<Script>` element at all has `Language: None`. |
| `LanguageNote` | all | Present when the language was defaulted or could not be read. |
| `ConfigVersion`, `ConfigTime`, `ConfigUser` | all | From the database; time in UTC. |
| `TemplateId`, `InstanceId` | all | Present when the object is part of a template or instance. |
| `InTemplate` | all | `True` if the object is inside a template definition (under the template itself, not an instance of it). |
| `ControlledByInstance` | all | `True` if the object belongs to an instance and so takes its content from a template. |
| `DocumentContentOverride` | all | `True` if this object holds its own copy of the content, overriding the template. |
| `TemplateSource` | all | One line per hop from this object to its editable source, in order. Absent when the object is edited in place. |
| `EditableSource`, `EditableSourceId` | all | Full name and ID of the object to edit to change this script. It is this object itself when `EditInPlace` is `True`. Use `EditableSource` to de-duplicate findings across thousands of instances. |
| `EditInPlace` | all | `True` if this object is where the script is edited. |
| `EditableSourceNote` | all | Present when a template object could not be resolved; `EditableSource` is then the last object reached. |
| `Source` | all | `GetWebDocumentContent` (mimics) or `GetScriptSource` (libraries). |
| `Server`, `ExportedAt`, `Exporter` | all | Where and when the file was exported, and by which version. |
| `HasOwnScript` | mimic | Whether the mimic has its own code, as opposed to only library links or pick actions. |
| `ScriptLines`, `ScriptChars` | all | Size of the script body. |
| `LibraryCount`, `Library` | mimic | Linked libraries that exist, in link order. |
| `LibraryCheck` | mimic | `Performed`, `NotPerformed (...)` or `Incomplete: ...`. |
| `LibraryMissing` | mimic | A linked library that does not exist (only with `-CheckMissingLibraries`). |
| `PickActionCount`, `PickAction` | mimic | Script pick actions, as seven ` \| `-separated fields: `N \| Where \| ActionName \| Type \| Context \| Function \| Args`. See the field notes below. |
| `PickActionCode` | mimic | `N \| first line \| last line` of pick action *N*'s inline code. |
| `EmbeddedMimicCount`, `EmbeddedMimic` | mimic | Embedded mimics, as `Where \| full name`. Relative references are resolved, and the name is `(animated)` when the reference is set by an animation. |
| `ReferencedByMimics` | library | How many exported mimics link to the library. Only counts mimics in scope when `-Group`/`-Name` is used. |
| `ScriptStartLine`, `ScriptEndLine` | all | See the parsing rules. |

**`PickAction` fields**

- **`N`** numbers the action within the file.
- **`Where`** is the path of named containers, such as `Layer[Layer]/Grp_1[Group]/Btn_3[Button]`, or `(mimic)` for the mimic background.
- **`Type`** is `scriptCall`, `scriptCode` or `scriptTopLevel`.
- **`Context`** is `Current` or `TopLevel`.
- **`Function`** is the function name for `scriptCall`, and `(code)` for inline code.
- **`Args`** is `(none)` or a comma-separated list where an animated argument appears as `name <= expression`. Split on at most the first six separators, because `Args` can contain ` \| `.

### Manifest

`_ScriptExport_Manifest.csv` has one row per exported file, per empty script library, and per
object that failed. Mimics with no scripting are not listed.

Columns: `FullName, Class, Id, Status, File, Language, LanguageGuid, ScriptLines, ScriptChars,
HasOwnScript, LibraryCount, Libraries, MissingLibraries, PickActionCount, ReferencedByMimics,
ConfigVersion, ConfigTime, ConfigUser, TemplateId, InstanceId, InTemplate, ControlledByInstance,
DocumentContentOverride, EditableSource, EditableSourceId, Message`.

`Status` is one of:

- `Exported`
- `Empty`: a library with no code, so no file is written.
- `Error`: the reason is in `Message`.

Existing files in the output folder are overwritten, never deleted. The script warns when the
folder contains `.txt` files it did not write in this run, for example left over from an earlier export.
The analyser uses the manifest to tell which files belong to the current export (see
[Export check](#export-check)).

**Spreadsheet safety.** A text cell starting with `=`, `+`, `-` or `@` (or a tab or carriage return) is
written with a leading `'`, so Excel shows it as text instead of evaluating it as a formula. Names
and messages come from the database, so this stops a crafted object name from running as a formula
when the CSV is opened. A program that reads the manifest should remove a leading `'` when the next
character is one of those; the analyser does. The analyser's own CSV files use the same convention.

## Analyser

`GeoSCADA-Script-Analyser.ps1` reads the exporter's output folder and reports what needs to be
updated. It works **offline**: it does not connect to Geo SCADA.

```powershell
.\GeoSCADA-Script-Analyser.ps1 -InputFolder C:\Export\Scripts
.\GeoSCADA-Script-Analyser.ps1 -InputFolder C:\Export\Scripts -MinimumSeverity Warning -OutputFolder C:\Export\Report
```

It writes five files:

- `_ScriptAnalysis_Actions.csv`: **the work list**, one row per object to edit. It is ordered with the most severe objects first, and among those, the objects whose edit fixes the most instance copies first.
- `_ScriptAnalysis_Instances.csv`: the instance copies each object's edit fixes, one row each.
- `_ScriptAnalysis.csv`: one row per finding.
- `_ScriptAnalysis_ExportCheck.csv`: problems with the export itself, such as missing, left-over or broken files. See [Export check](#export-check).
- `_ScriptAnalysis.html`: a self-contained interactive report. It has:
  - a disclaimer at the top: the report is guidance only, and there is no guarantee that every issue has been identified;
  - a summary of **every** rule, in rule order (the same order as the rule filter). Click a column heading to sort by it, and click again to reverse. A rule with nothing found shows `0`. A rule that did not run shows `—` and the reason: left out by `-IncludeRule` / `-ExcludeRule`, RN06/RN07 under `-SkipProgIdCheck`, or a severity below `-MinimumSeverity`. So the report always shows what was checked;
  - the objects to edit, each with its role, "fixes N instance copies", the action to take, its findings, and an expandable list of the instances it covers;
  - search and filters for severity, role, rule and language.
  - The rule filter is a multi-select list of every rule, in rule order (RN01…, then GS01…), with its finding count or "(not run)". It has All / None / RN only / GS only shortcuts. Clicking a rule in the summary table shows only that rule.
  - Unticked rules are remembered in that browser, so rules you have decided not to fix stay hidden in later reports. Rules that are new to a report start ticked.
  - To leave rules out of the CSV files as well, run the analyser with `-ExcludeRule`, for example `-ExcludeRule GS03,GS04`.
  - Problems found by the export check appear in a note at the top. The note is red when the export is incomplete.

### Export check

A report is only as complete as the export it reads. Before the rules run, the analyser checks the
folder against the exporter's manifest, and checks the header of every file it opens:

| Issue | Meaning | Effect |
| ----- | ------- | ------ |
| `MissingFile` | The manifest says the object was exported, but its file is not in the folder. | Not in the report. **Incomplete.** |
| `ExportError` | The exporter could not export the object (a manifest `Error` row). | Not in the report. **Incomplete.** |
| `Malformed` | The file starts like an exporter file but its header is broken: no END marker (truncated), no `FullName`, an invalid `FormatVersion`, or line numbers outside the file. | Not analysed. **Incomplete.** |
| `DuplicateName` | Two files have the same `FullName` in their header. | Both are analysed. |
| `ManifestMismatch` | The file's header names a different object (`FullName` / `ObjectId`) from the manifest row for that file. | Analysed. |
| `UnsupportedFormat` | The file's `FormatVersion` is newer than the analyser knows. | Analysed; findings may be incomplete. |
| `MixedExports` | The files come from more than one export run (`ExportedAt`) or exporter version / `FormatVersion`. | Information only. |
| `NotInManifest` | The file is not in the manifest, for example left over from an earlier export or from a run with a different `-Group` scope. | Analysed, unless `-ManifestOnly` is given. |
| `NoManifest` | The folder has no manifest. | Missing files and left-over files cannot be detected. |

The problems are listed in `_ScriptAnalysis_ExportCheck.csv`, in a note at the top of the HTML
report, and on the console. When there are `MissingFile`, `ExportError` or `Malformed` problems,
the analyser still writes the report but **ends with exit code 3**, so a calling script can tell
that the report is based on an incomplete export.

Instance copies that are counted from the manifest without being opened (see
[How it works](#how-it-works)) are only checked for being present, not for their header.

An object name ending in a space lost that space in headers written by exporter 1.0.0. The
analyser then uses the name from the manifest, so the object still matches its instance copies.

### Batching by the object to edit

Most mimics in a large system are instance copies of a template, and their script can only be
changed in the template. So each problem is reported **once, against the object you actually
edit** (the header's `EditableSource`), together with how many instance copies that one edit
fixes. Each object to edit has a **role**:

| Role | Meaning | What to do |
| ---- | ------- | ---------- |
| `Template` | The object is inside a template definition. | Edit it once; the change reaches all `InstancesFixed` copies. |
| `OverriddenInstance` | An instance that holds its own copy of the content (`DocumentContentOverride: True`). | Edit **this instance** directly. Changes to its template do not reach it, so it is called out on its own even when its template has the same problem. |
| `Instance` | An instance copy whose template is not in this export, for example because of `-Group` scope. It is scanned as is. | Make the change in its template, named by `EditableSource` in its header. |
| `Standalone` | Not part of a template or an instance. | Edit it. |

The exporter does not record which template an overridden instance overrides. So the report can
say that a template fix won't reach its overridden instances, but it can't list which ones
belong to which template.

### How it works

- **Inline pick-action code.** The inline code of `scriptCode` and `scriptTopLevel` pick actions is scanned with the same rules as the script, in the mimic's language; a mimic with no script is treated as VBScript. Findings in it show `Section` = `PickAction <n>`, or `PickAction <n> (top-level mimic)` for code that runs in the top-level mimic. That code really runs in the top-level mimic's language, which the export cannot know.
- **Function references.** Only `scriptCall` actions name a function (GS02–GS05). The function is looked for in the mimic's script and its linked libraries. When the action is set on an embedded-mimic object, it is also looked for in the embedded mimic's script and libraries, following instance copies to their editable source.
- **Header parsing.** Each file's header is parsed (see [Output format](#output-format)), and the script body is scanned with comments masked out. String contents are also masked, except where a rule needs the literal, such as a COM ProgID. In JScript and JavaScript, regular expression literals are masked like strings: a `/` starts one after an operator or punctuation, or after a keyword such as `return`, `typeof`, `case` or `in`. Findings give the script line and the file line.
- **Template de-duplication.** A file whose header says `EditInPlace: False` takes its script from `EditableSource`, so only the editable object is scanned. With the exporter's manifest present, those instance copies are counted from it without being opened, which saves most of the disk reads. See [Batching by the object to edit](#batching-by-the-object-to-edit). `-IncludeInstances` scans every file instead.
- **Case-sensitivity scope.** The case check (RN11) only looks at member names on Geo SCADA objects. That means chains starting at `Server`, `Mimic`, `Form`, `Menu`, `App`, `Window` or `Clipboard`, or at a variable assigned from one. Names are compared with the ViewX script type library, so a script's own JavaScript objects (for example `data.value`) are not flagged.
- **COM checks.** The checks for 32-bit and unregistered COM classes (RN06/RN07) look ProgIDs up in **this machine's** registry, in the 64-bit view. Run the analyser on a machine built like your ViewX clients, or pass `-SkipProgIdCheck`.

### Rules

Severities:

- **Error:** fails or misbehaves in 2025 R2.
- **Warning:** likely to misbehave; check it.
- **Review:** behaviour or layout changed; look at it.
- **Info:** advisory.

Edit the rule table at the top of the script to change a rule's severity.

| Rule | Severity | Finds | Source |
| ---- | -------- | ----- | ------ |
| RN01 | Error / Warning | `.Value(x)` with one argument used as a statement to set a value, such as `x.Interface.Value(3.14)`, `Call x.Value(1)` or `If a Then x.Value(1)`. Error for `Interface.Value(x)`; Warning otherwise. Reads such as `y = (x.Interface.Value(0))` and `.Value(row, col)` on a query control are not flagged. | Release notes p20 |
| RN02 | Warning | Return value of `Menu.Init`, `AddItem` or `Separator` used (now `Null`). | p20-21 |
| RN03 | Warning | VBScript Form control (a variable `Set` from `Form.Add…` in the same procedure) used as a value without `.Value`, e.g. `MsgBox field`, `ElseIf field Mod 500 = 0`, `y = field`, `field = 5`, `CStr(field)`. Not flagged: member use (`field.Value`), `Set y = field`, `field Is Nothing`, `IsObject` / `TypeName`, and passing the control to your own sub or function. | p21 |
| RN04 | Review | `Form.AddListBox`: the row count is now exact. | p21-22 |
| RN05 | Review | `Form.AddQuery`: the row count is now exact. | p22-23 |
| RN06 | Error | `CreateObject` / `ActiveXObject` of a COM class registered only as a 32-bit in-process server. | p16 |
| RN07 | Warning | COM class not registered on this machine. | p16 |
| RN08 | Review | COM object created from a computed (non-literal) ProgID. | p16 |
| RN09 | Error | Clipboard through `InternetExplorer.Application` or `clipboardData`. | p23 |
| RN10 | Error | `ScriptControl` COM object. | p23 |
| RN11 | Error | JScript / V8 wrong-case Geo SCADA name, e.g. `LookUpObject`, `Form.init`, `.value`, `.checked`, `.sql`. | p24 |
| RN12 | Error | JScript / V8 number assigned to a Boolean property, e.g. `ok.Default = 1`. | p24 |
| RN13 | Warning | JScript pick action passes arguments; this fails with JScript9Legacy on Windows 11 24H2. | p24-25 |
| RN14 | Info | JScript in use (not recommended; also has a Windows Server 2025 run-mode issue). | p25, p58 |
| RN15 | Error | `CreateObject` / `ActiveXObject` in JavaScript V8. | p26-27 |
| RN16 | Error | `MsgBox` / `InputBox` in JavaScript V8. | p26 |
| RN17 | Error | `LocalTimeToUTC` / `UTCToLocalTime` in JavaScript V8. | p27 |
| RN18 | Review | Form layout in JavaScript V8 (pixel units). | p27 |
| RN19 | Info | JavaScript V8 script: needs 2025 R2 clients; not supported by Original WebX. | p19, p25-26 |
| RN20 | Info | `Stop` / `debugger` statement; the Microsoft Script Debugger has been removed. | p6, p19 |
| RN21 | Error / Warning | `GetRegistry("Key")` with no default value: it returns Null or an empty string when the key is not set, so converting it fails. Error when the call is wrapped directly in `CLng`, `CInt`, `CDbl` and similar; Warning otherwise. The fix is to pass a default: `CLng(GetRegistry("Key", 0))`. | 2025 R2 script runtime (found in the field, not in the release notes) |
| RN22 | Error / Warning / Info | A VBScript variable called `ToString`, `GetType`, `GetHashCode` or `Equals` (any case). Undeclared, the 2025 R2 VBScript runtime resolves it to the .NET member of that name, so `toString = x` fails with "The object has no suitable property or field named 'ToString'". All four names have been tested. An undeclared use is an Error on every line that uses it, and the fix is to rename the variable or `Dim` it. Declared with `Dim` / `ReDim` / `Const` in the same Sub or Function, it works (tested), so it is one Info on the declaration. A script-level declaration, a parameter, or a Sub / Function with that name is one Warning, because those cases are untested. Member calls such as `x.ToString` and JScript are not flagged. | Found in the field; since confirmed by the Knowledge Base article |
| RN23 | Warning | A VBScript variable `Set` from `Server.FindObject`, `Server.LookupObject` or `Server.ThisObject` (optionally `.Parent`) used as a value, for example `CStr(obj)`, `MsgBox obj`, `"Id: " & obj` or `obj = x`. A database object no longer has a default property, so `CStr(obj)` no longer gives its Id. The same exemptions as RN03 apply, and an object reached through `.Interface` is not tracked. | KB: Object Default Properties |
| RN24 | Warning | A "not found" test that never works: `IsNull(obj)`, `IsEmpty(obj)`, `obj = Null` or `IsObject(obj)` on a `FindObject` / `LookupObject` result. These return `Nothing`, so use `If obj Is Nothing Then`. **Not a 2025 R2 change**: the test never worked, and the not-found case fails later with "Object required". Skipped when the script, or a library the mimic links, defines its own function of that name. | KB: Server.LookupObject returns Nothing |
| RN25 | Warning | A date passed through the registry without conversion. Flagged: `SetRegistry` whose value is a date (`Now`, `Date`, `Time`, `DateAdd(…)` and the other date functions, a `Form.AddDateTime` control's `.Value`, or a variable set from one in the same procedure); a `GetRegistry` value passed straight to a date function (`FormatDateTime`, `DateDiff`, `Year` …); and a `GetRegistry` value written to a date-time control. The fix is `SetRegistry "Key", CLng(Int(CDbl(d)))`, then `CDate(GetRegistry("Key", 0))` to read it back. | KB: No implicit date/time to number conversions |
| RN26 | Error / Warning | A Call Specific Function pick action that gives no value to one or more of the function's parameters. In 2025 R2 the function **does not run at all**, and no error is shown. Before R2 the missing parameters were passed as `Empty`. The function's parameters, found as for GS03, are compared with the pick action's `name <= value` entries; a parameter that is missing, or listed without a value, counts as unset. The fix is to give every parameter a value (`''` or `0` if it is not used), or to remove unused parameters. Error for VBScript, where it has been seen; Warning for JavaScript, which is untested. A call written with inline arguments is not checked. | Found in the field (reported to Schneider) |
| GS01 | Error | Link to a script library that does not exist. Needs the export run with `-CheckMissingLibraries`. | Exporter header |
| GS02 | Error | Call Specific Function pick action (`scriptCall`) on a mimic that has no script. | Exporter header |
| GS03 | Error | `scriptCall` pick action names a function not defined in the mimic script or its libraries. For an action on an embedded-mimic object, the embedded mimic's script and libraries are also checked. A single inline call (`FilterAlm 1, "x"`) is checked by its first name. | Exporter header |
| GS04 | Error | Pick action function name differs only in case (JScript / V8). | Exporter header |
| GS05 | Info | Pick action function could not be checked because a linked library is not in the export. | Exporter header |
| GS06 | Error | Mimic links a library written in a different language. | Exporter header |
| GS07 | Warning | Script syntax looks like a different language from the one set. | Script body |
| GS08 | Info | Script library not linked by any exported mimic. | Exporter header |

The rule groups are:

- **RN** rules: changes in the 2025 R2 script runtime. The page numbers refer to *Geo SCADA Expert 2025 R2 Release Notes* (88.9742.1). "KB" refers to the Geo SCADA Knowledge Base article [Adjustments for VBScript in Geo SCADA 2025 R2](https://community.se.com/t5/Geo-SCADA-Knowledge-Base/Adjustments-for-VBScript-in-Geo-SCADA-2025-R2/ba-p/555067). It repeats most of the release-note items and adds new ones as they are found (RN22–RN25). RN21 and RN26 were found in the field.
- **GS** rules: consistency and good-practice checks on what the export recorded.

The rules are pattern-based, so treat findings as pointers to review, not as proof.

### Parameters

| Parameter           | Default            | Description |
| ------------------- | ------------------ | ----------- |
| `-InputFolder`      | *(required)*       | The exporter's output folder. |
| `-OutputFolder`     | `-InputFolder`     | Where to write the report. |
| `-ReportName`       | `_ScriptAnalysis`  | Base name of the `.csv` / `.html` files. |
| `-MinimumSeverity`  | `Info`             | Lowest severity to report: `Error`, `Warning`, `Review` or `Info`. |
| `-IncludeRule`      |                    | Only run these rule IDs. |
| `-ExcludeRule`      |                    | Skip these rule IDs. |
| `-IncludeInstances` | *(off)*            | Scan instance copies too, instead of only each script's editable source. |
| `-SkipProgIdCheck`  | *(off)*            | Do not look COM ProgIDs up in this machine's registry. |
| `-NoHtml`           | *(off)*            | Write the CSV only. |
| `-ManifestOnly`     | *(off)*            | Analyse only the files the manifest lists as exported. Left-over files are still reported as `NotInManifest`. |

### Report columns (CSV)

**`_ScriptAnalysis_Actions.csv`**: `Priority, Severity, FullName, Role, Class, Language,
InstancesFixed, LinkedByMimics, Findings, Errors, Warnings, Reviews, Infos, Rules, Action, File`.

- `Severity` is the worst finding on the object.
- `Rules` lists the rules hit, for example `GS03 x5; RN21`.
- `LinkedByMimics` is set for script libraries.

**`_ScriptAnalysis_Instances.csv`**: `EditableSource, Priority, Instance`. There is one row for each instance copy covered by an object with findings.

**`_ScriptAnalysis_ExportCheck.csv`**: `Issue, FullName, File, Detail`. There is one row per problem found by the [export check](#export-check). The file has only its header line when the export is complete.

In all the analyser's CSV files, a text cell starting with `=`, `+`, `-` or `@` begins with `'`, as in the manifest.

**`_ScriptAnalysis.csv`**: `Severity, Rule, Title, FullName, Class, Language, Section, ScriptLine,
FileLine, Code, Detail, Fix, Reference, Role, InstancesFixed, AffectedObjects, File, ObjectId,
ExporterVersion, FormatVersion, AnalyserVersion`. The last three record which exporter version and file format
the object's file came from, and which analyser version made the report.

`Section` is one of:

- `Script`: the object's script. `ScriptLine` is the line within it.
- `PickAction <n>` (optionally `(top-level mimic)`): that pick action's inline code. `ScriptLine` is the line within the code.
- `Header`: findings about the header, such as missing libraries or pick-action functions. Use `FileLine` for these.

`FullName` is the object to edit. `InstancesFixed` is the number of instance copies that edit
also fixes. `AffectedObjects` is the object itself plus those copies.

## Examples

`Examples/` holds before-and-after demonstrations of the version-dependent RN issues, for testing a
system before and after upgrading to 2025 R2:

| File | Mimic language | Covers |
| ---- | -------------- | ------ |
| `RN-Examples-VBScript.vbs` | VBScript | RN01, RN02, RN03, RN04, RN05, RN06, RN10, RN21, RN23, RN24, RN25 |
| `RN-Examples-JScript.js` | JScript | RN11, RN12 |

**Setting up the mimic.**

1. Paste the script into a mimic of the matching language.
2. Add a pick action of type **Call Specific Function** → `Run_Examples`.
3. For RN01, create an Internal Analog Point called `Test Analog` in the same group as the mimic. The example writes `Interface.Value` to it, so never point it at a real tag. A third RN01 step, Read, checks that `y = (x.Interface.Value(0))` still works.

**What happens when it runs.**

- Each issue runs its **Before** code (as typically written for 2025 and earlier), then its **After** code, which works in every version.
- Each step reports OK or the error it raised, and one failing step does not stop the run. A summary is shown at the end.
- On 2025 and earlier, both steps should report OK.
- On 2025 R2, each Before step is expected to fail, or for RN04/RN05 to look wrong, while each After step works.
- Expectations marked *(to confirm)* come from the release notes or a field report and have not yet been checked side by side on both versions.

**Issues left out.** These can't show "works before, fails after":

- **RN15–RN19:** JavaScript V8 doesn't exist before 2025 R2.
- **RN09:** the fix, the `Clipboard` object, only exists in 2025 R2.
- **RN13:** depends on the Windows 11 24H2 JScript engine.
- **RN07/RN08:** depend on the machine.
- **GS rules:** these are broken on every version.
- **RN26:** this is about the pick action's configuration, not the script, so it needs a mimic set up by hand. Add `Public Sub DoNothing(One, Two)` with `MsgBox "One"` in it, and a Call Specific Function pick action to `DoNothing` that sets only `One`. Before R2 the message shows; in 2025 R2 nothing runs. Give `Two` a value and it runs in both.

RN24 is included even though it is not an R2 change: its Before step fails with "Object required" in **every** version, because `LookupObject` has always returned `Nothing`. It looks up object Id 99999999, which must not exist. RN25 writes the registry key `Claude Examples - Date` in ViewX's own registry area, and reads it back.

Scripts saved in 2025 R2 can't be used by older clients (release notes p19). To run these
examples on an older version, create or export the mimic in a format that version can read.

## Versions

The exporter and the analyser are versioned separately, using semantic versioning (major = breaking, minor =
new features or rules, patch = fixes). The export file layout has its own `FormatVersion`. See
[CHANGELOG.md](CHANGELOG.md) for the history and the rules.

- **Each exported file** records the exporter that wrote it and its format, e.g. `Exporter: GeoSCADA-Script-Exporter 1.1.0` and `FormatVersion: 2`, together with `ExportedAt`.
- **The analyser** reports which exporter versions, formats and export runs it read, and its own version. These appear on the console, in the HTML report header, and in the `ExporterVersion`, `FormatVersion` and `AnalyserVersion` columns of `_ScriptAnalysis.csv`.
- **Mixed folders:** a folder holding files from more than one export run or exporter version is listed as `MixedExports` in the export check.

## Security

- **Read-only.** The script lists objects, reads documents and calls `GetScriptSource`. It makes no configuration changes.
- **Credentials.** Prefer `-Credential` (or the prompt) over `-Password` on the command line. The connection logs off and disconnects in a `finally` block.
- **Output contents.** Exported scripts can contain sensitive information: system topology, control logic, file paths, and sometimes embedded credentials. The script prints a reminder at the end of every run. Store, share and dispose of the files in line with your data protection and security policies.
- **Partial exports.** If any object cannot be read or written, the script still finishes, prints `EXPORT INCOMPLETE` and exits with code **2**. `-FailOnError` stops at the first failure instead.

## Disclaimer

This script is provided **"AS IS"**, without warranty of any kind, express or implied. The author
accepts no liability for any damages arising from its use. It connects to the Geo SCADA server
through the .NET Client API to read mimic and script library definitions and writes text files.
It is **NOT** certified for production SCADA environments. Do **NOT** run it against a production or
safety-critical system without first reviewing the code and testing it on a representative
non-production system. You run it at your own risk and are responsible for compliance with your
own change-control and security policies.

You are prompted to accept this disclaimer before the script runs (or pass `-AcceptDisclaimer`).

## License

Copyright (c) 2026 Adam Woodland.

Licensed under the **MIT License** &mdash; see the [LICENSE](LICENSE) file, or
<https://opensource.org/licenses/MIT>. This is free software, and you are welcome to
redistribute it under those conditions; it comes with **ABSOLUTELY NO WARRANTY**.
