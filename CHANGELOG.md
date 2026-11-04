# Changelog

The exporter and the analyser are versioned separately, using [semantic versioning](https://semver.org/):

| Change | Version bump | Examples |
| ------ | ------------ | -------- |
| Breaking | **major** (2.0.0) | A CSV column removed or renamed, a parameter removed, a default changed, the export file format changed incompatibly |
| New, backwards-compatible | **minor** (1.3.0) | A new rule, report, parameter or CSV column added at the end |
| Fix only | **patch** (1.2.1) | A false positive fixed, a wording change, a crash fixed |

The export **file format** has its own number, `FormatVersion`, written in every file header. It changes only
when the layout of the files changes. The analyser reads every format up to its `$SupportedFormat`, and reports
newer ones as `UnsupportedFormat`.

Every exported file records the exporter that wrote it (`Exporter: GeoSCADA-Script-Exporter 1.1.0`) and its
`FormatVersion`. The analyser reports which exporter versions and formats it read, and its own version. These
appear on the console, in the HTML report header, and in the `ExporterVersion`, `FormatVersion` and
`AnalyserVersion` columns of `_ScriptAnalysis.csv`.

## GeoSCADA-Script-Analyser

### 2.0.0 (2026-11-04)

**Breaking: the rule IDs have changed.** The Resource Centre (Knowledge Base) article "Adjustments for VBScript in Geo SCADA 2025 R2" now numbers its items, and the rule IDs follow it. New prefixes are used, so no 1.x ID is reused with a different meaning:

- **RC*nn*** (Resource Centre) is article item *nn*. A letter is added when one item has several checks (RC06a, RC06b). Item 5, form layout, needs no code change, so there is no RC05.
- **JS*nn*** is the article's "Specific JScript Issues" section, in article order.
- **RX*nn*** is any other 2025 R2 change, not in the article: from the release notes, or found in the field.
- **GS01–GS08** are unchanged.
- The 1.x IDs **RN01–RN28** are retired and will not be used again.

This changes the `Rule` column of the CSV files and the rule references in the HTML report. To make the move easier:

- `-IncludeRule` / `-ExcludeRule` still accept the 1.x IDs. Each one is translated to its new ID, with a note on the console (`-ExcludeRule RN27 is the analyser 1.x ID of RC06b ...`).
- The HTML report shows each rule's 1.x ID ("was RN27") in the by-rule table and the rule filter, and the search box finds findings by their 1.x ID.
- The HTML report carries rule-filter choices saved by a 1.x report over to the new IDs, the first time a 2.0.0 report is opened in that browser.
- The README rules table has a "1.x rule" column.

Entries below 2.0.0 in this file use the 1.x IDs.

| Old | New | | Old | New | | Old | New |
| --- | --- | - | --- | --- | - | --- | --- |
| RN01 | RC01 | | RN11 | JS02 | | RN21 | RX10 |
| RN02 | RC02 | | RN12 | JS03 | | RN22 | RC10 |
| RN03 | RC03 | | RN13 | JS04 | | RN23 | RC04 |
| RN04 | RC06a | | RN14 | JS05 | | RN24 | RC09 |
| RN05 | RC07 | | RN15 | RX05 | | RN25 | RC11 |
| RN06 | RX01 | | RN16 | RX06 | | RN26 | RC12 |
| RN07 | RX02 | | RN17 | RX07 | | RN27 | RC06b |
| RN08 | RX03 | | RN18 | RX08 | | RN28 | RC13 |
| RN09 | RC08 | | RN19 | RX09 | | *new* | RC14a, RC14b |
| RN10 | JS01 | | RN20 | RX04 | | | |

- **New rules, RC14a (Warning) and RC14b (Info)**, from the article's 2026-11-04 update (item 14: "The Server.GetOPCValue method cannot read aggregates"). In the 2025 R2 release version, `Server.GetOPCValue` raises an error for an aggregate or a property inside one, such as `myPoint & ".Control.ControlMaximum"`. Schneider call it a temporary limitation that a future update will fix, so the text says to check the current release notes.
  - A tag is a dotted name, so `Object.Property` and `Object.Aggregate.Property` look alike. RC14a is a call whose tag text names one of the 110 aggregates in the 2025 R2 database schema, after a dot. `AlarmSummary`, `PSTN` and `Shared` are also property names, so they count only when another name follows.
  - RC14b lists every other `GetOPCValue`, so that tags built at run time can be reviewed. It can be hidden with the rule filter.
- **References** cite the article's item numbers ("KB #6"). RC06b (formerly RN27) now cites item 6, which includes the horizontal-scrolling note.
- **HTML report:**
  - The rule filter has RC only, JS only and RX only shortcuts, and rules are listed RC, JS, RX, then GS.
  - The severity and role filters are multi-select drop-downs like the rule filter, so several can be chosen at once (for example Error and Warning). Each choice shows its count, and All / None select or clear them all.
  - A drop-down that would open past the edge of a narrow window is moved back inside it.
- **Examples:** the example files are now `Examples/Rule-Examples-VBScript.vbs` and `Examples/Rule-Examples-JScript.js`. The before / after functions are renamed to the new IDs (for example `RC12_PickAction`, `RC13_Before`), and the VBScript examples have an RC14 pair that reads the mimic's own `GISLocation` aggregate.

### 1.4.0 (2026-11-03)

- **New rule, RN28 (Warning / Info):** a log on or log off stops the script in 2025 R2. It covers `App.LogOn`, `App.LogOff`, `App.LogOnUser`, `App.LogOffUser`, and a `Logon` through ViewX COM automation (`CreateObject("Serck.ViewXApplication")`). It is a Warning when code follows the call in the same procedure, and Info when the call is already the last statement. From the Knowledge Base article's 2026-11-03 update.
- **RN26:** now cites the Knowledge Base article, which has added the item. The fix text says the function fails to run in the 2025 R2 release version, that Schneider expect a future update to fix it, and to check the current release notes for the installed version.
- The Knowledge Base reference records the date the article was last checked (2026-11-03).

### 1.3.1 (2026-10-29)

- RN27 is worded as an expected 2025 R2 change: the "probable bug" note and its reference were removed. The HTML report no longer shows empty brackets for a rule without a reference.
- The HTML disclaimer now also says that some suggested changes may not work in earlier versions of Geo SCADA, or may change how scripts look or behave, and that every change should be tested with each Geo SCADA client version in use.

### 1.3.0 (2026-10-28)

- **New rule, RN27 (Review):** every `Form.AddListBox`. In 2025 R2 a list box holding more items than its rows shows a horizontal scroll bar instead of a vertical one. List contents are usually filled at run time, so every list box is listed to check. Found in the field and reported to Schneider as a probable bug; the Knowledge Base article also notes it.

### 1.2.2 (2026-10-22)

- Fixed false positives in the COM checks. RN06–RN08 (and RN09, RN10, RN15) matched any `CreateObject(`, including method calls on an object such as `objServer.CreateObject(pointClass, parent)`, a Geo SCADA method that creates a database item. These were reported as RN08 (computed ProgID), and with a literal class name would have been reported as RN07 (not registered). Only the built-in `CreateObject` / `ActiveXObject` are checked now.

### 1.2.1 (2026-10-22)

- The HTML report opens with a disclaimer: the report is guidance only, the checks are pattern-based, and there is no guarantee that every issue has been identified.

### 1.2.0 (2026-10-21)

- **New rules.**
  - RN23: a database object used as a value.
  - RN24: an object-not-found test using `IsNull` or `= Null` instead of `Is Nothing`.
  - RN25: a date passed through the registry without conversion.
  - RN26: a pick action that leaves function parameters without a value.
- **Version reporting.**
  - The console and HTML header show the exporter version(s), file format(s) and export run(s) read, and the analyser version.
  - `_ScriptAnalysis.csv` gains three columns at the end: `ExporterVersion`, `FormatVersion` and `AnalyserVersion`.
  - A new export-check issue, `MixedExports`, is raised when the files come from more than one export run or exporter version. It is information only.
- **HTML report.** The by-rule table and the rule filter list every rule. `0` means checked with nothing found; a rule that did not run shows why.
- **References.** RN22 now cites the Knowledge Base article "Adjustments for VBScript in Geo SCADA 2025 R2".
- **Note.** RN23–RN25 and the full rule list were first pushed in commit `6a3e237` under version 1.1.0. They belong to 1.2.0.

### 1.1.0 (2026-10-15)

First published version.

- **Rules:** RN01–RN22 for the 2025 R2 script runtime (release notes and field findings) and GS01–GS08 consistency checks.
- **Batching:** findings are reported against the object to edit (template, overridden instance, instance, standalone), with the number of instance copies each edit fixes.
- **Export check:** the folder is checked against the exporter's manifest (missing, broken, duplicate, mismatched and left-over files). Exit code 3 when the export is incomplete; `-ManifestOnly` analyses only the listed files.
- **Output:** findings, actions, instances and export-check CSVs, with a spreadsheet formula guard, plus an interactive HTML report.

## GeoSCADA-Script-Exporter

### 1.1.0 (2026-10-15), FormatVersion 2

- **File format 2:**
  - inline pick-action code in its own `PICKACTION <n>` sections;
  - pick-action context (current or top-level mimic);
  - embedded mimics;
  - `ScriptEndLine`;
  - 7-field `PickAction` lines.
- **Templates:** each object's editable source (template resolution) is recorded in the header and the manifest.
- **File names:** Windows device names with an extension (`CON.Child`) get a leading `_`.
- **Manifest:** cells that would start like a spreadsheet formula get a leading `'`.
- **Errors:** friendlier connection and logon errors.

### 1.0.0 (2026-10-08), FormatVersion 1

- First release: mimic scripts and script library code (VBScript, JScript, JScript9, JavaScript V8) exported to one `.txt` per object, with a parseable header and a manifest CSV.
