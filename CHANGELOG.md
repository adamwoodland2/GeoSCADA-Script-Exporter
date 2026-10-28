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
