# GeoSCADA-Script-Exporter

Exports every **client-side script** in a **Geo SCADA** database to text files: mimic scripts
and Script Library code, in **VBScript, JScript, JScript9 and JavaScript V8** (new in Geo SCADA
2025 R2). Each file is named after the object's full name and starts with a structured header,
so a later step can parse the files to look for problems.

It uses the Geo SCADA **.NET Client API** (`ClearScada.Client.dll`) and only reads; nothing in
the database is changed. Logic programs are out of scope.

A companion to [GeoSCADA-Mimic-Walker](https://github.com/adamwoodland2/GeoSCADA-Mimic-Walker).

## How it works

1. **Connects and logs on** through the .NET Client API (`ClearScada.Client.Simple.Connection`)
   and lists the `CMimic` and `CScriptLibrary` objects with SQL. SQL does not return objects
   inside template definitions, so each template (`CTemplate`) is also asked for its objects with
   `IObjectHierarchy.ListObjects`. Template mimics, which are the ones you edit, are therefore
   exported as well.
2. **Mimics:** `IObjectConfiguration.GetWebDocumentContent()` returns the mimic definition as
   XML (the same XML the web server serves for `/db/<Mimic>?content`). From it the script reads:
   - `<Script lang="{GUID}">` for the script language;
   - the text of `<Script>`, which is the mimic's own code;
   - `<Library name="...">` for each script library the mimic links to;
   - `<Action type="scriptCall">` for each pick action that calls a script function, recording
     the object, the function name and its arguments.

   A file is written if the mimic has its own code, links to a library, or has a pick action that
   calls a script. That last case covers mimics with no script at all, which is an issue in itself.
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
- Device names such as `CON` get a leading `_`.
- If two objects would end up with the same file name, the second gets ` [<Id>]` appended.

The real full name is always in the header and the manifest.

Paths of 248 characters or more are written with the `\\?\` prefix. Use a short `-OutputFolder`
if your tools cannot open long paths.

### File structure (FormatVersion 1)

Files are UTF-8 without a BOM, with CRLF line endings. A file has three parts:

1. a `BEGIN` marker line;
2. header lines, each `Key: Value`;
3. an `END` marker line, followed immediately by the script exactly as stored.

Script line endings are normalised to CRLF; nothing else is changed. Every header and marker
line starts with the language's line-comment token: `'` for VBScript, `//` for the
JavaScript family and unknown languages. That keeps the whole file valid for the language's
parser or linter.

```
' ===== GEOSCADA-SCRIPT-EXPORT BEGIN =====
' FormatVersion: 1
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
' Exporter: GeoSCADA-Script-Exporter 1.0.0
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
' PickActionCount: 1
' PickAction: New Layer[Layer]/Btn_3[Button] | New Action | scriptCall | RunAudit | strPath <= ".Source Path.Value", strOutputFile <= ".Output File Name.Value"
' ScriptStartLine: 27
' ===== GEOSCADA-SCRIPT-EXPORT END =====
'***********************************************************************************************
...script...
```

**Parsing rules**

- The `BEGIN` marker is line 1. The header ends at the line containing `GEOSCADA-SCRIPT-EXPORT END`.
- For a header line, remove the comment token and one space, then split on the first `": "` to get the key and value.
- Values are single-line. Keys may repeat (`Library`, `LibraryMissing`, `PickAction`).
- `ScriptStartLine` is the file line (1-based) holding script line 1. Script line *n* is at file line `ScriptStartLine + n - 1`, which lets issues found in the file be mapped back to the script.
- An empty script (for example a mimic that only links libraries) has no lines after the `END` marker.

**Header keys**

| Key | Applies to | Meaning |
| --- | ---------- | ------- |
| `FormatVersion` | all | `1`. Bumped if the layout changes. |
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
| `PickActionCount`, `PickAction` | mimic | Pick actions that call a script, as five ` \| `-separated fields: `Where \| ActionName \| Type \| Function \| Args`. `Where` is the path of named containers, such as `Layer[Layer]/Grp_1[Group]/Btn_3[Button]`, or `(mimic)` for the mimic background. `Args` is `(none)` or a comma-separated list where an animated argument appears as `name <= expression`. Split on at most the first four separators, because `Args` can contain ` \| `. |
| `ReferencedByMimics` | library | How many exported mimics link to the library. Only counts mimics in scope when `-Group`/`-Name` is used. |
| `ScriptStartLine` | all | See the parsing rules. |

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
