# GeoSCADA-Script-Analyser.ps1
# Copyright (c) 2026  Adam Woodland
#
# Licensed under the MIT License. You may obtain a copy of the License in the LICENSE file
# distributed with this software, or at <https://opensource.org/licenses/MIT>.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
# BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
# NONINFRINGEMENT.

<#
.SYNOPSIS
    Scans the script files written by GeoSCADA-Script-Exporter.ps1 and reports what needs to be
    updated - chiefly for the script engine change in Geo SCADA 2025 R2 - as a CSV and an
    interactive HTML report.

.DESCRIPTION
    Works offline on the exported .txt files; it does not connect to Geo SCADA.

    Each file's header (see the exporter's README) is parsed and the script body is checked
    against a set of rules. Most rules come from the Geo SCADA 2025 R2 release notes ("Mimic and
    Global Script Execution" and following sections); the rest use the exporter's metadata:

      Release notes (RNxx)
        RN01  .Value(x) used as a statement to set a value (not reads)   (raises an error now)
        RN02  Return value of Menu.Init / AddItem / Separator used       (now Null)
        RN03  Form control used without .Value (default property)        (raises an error now)
        RN04  Form.AddListBox row count                                  (now exact - check layout)
        RN05  Form.AddQuery row count                                    (now exact - check layout)
        RN06  CreateObject / ActiveXObject of a 32-bit-only COM class    (ViewX is now 64-bit)
        RN07  CreateObject / ActiveXObject of a class not registered here
        RN08  CreateObject / ActiveXObject with a computed ProgID        (cannot be checked)
        RN09  Clipboard through Internet Explorer                        (use the Clipboard object)
        RN10  ScriptControl COM object                                   (not recommended / 32-bit)
        RN11  JScript / V8 wrong-case Geo SCADA name (e.g. .value, LookUpObject, Form.init)
        RN12  JScript / V8 number assigned to a Boolean property         (e.g. ok.Default = 1)
        RN13  JScript pick action passes arguments                       (fails with JScript9Legacy)
        RN14  JScript in use                                             (not recommended)
        RN15  CreateObject / ActiveXObject in JavaScript V8              (not supported)
        RN16  MsgBox / InputBox in JavaScript V8                         (use Window.Alert/Confirm/Prompt)
        RN17  LocalTimeToUTC / UTCToLocalTime in JavaScript V8           (not supported)
        RN18  Form layout in JavaScript V8                               (pixel units)
        RN19  JavaScript V8 client compatibility                         (2025 R2 clients only)
        RN20  Stop / debugger statement                                  (script debugger removed)
        RN21  GetRegistry without a default value                        (Null / empty when unset)
        RN22  Undeclared ToString / GetType / Equals / GetHashCode       (clashes with .NET member)
        RN23  Database object variable used as a value                   (no default property now)
        RN24  Object not found tested with IsNull / = Null / IsObject    (use Is Nothing)
        RN25  Date stored in / read from the registry unconverted        (no implicit date <-> number)
        RN26  Pick action leaves function parameters without a value     (function does not run at all)

      Exporter metadata (GSxx)
        GS01  Link to a script library that does not exist
        GS02  Call Specific Function pick action but the mimic has no script
        GS03  Pick action calls a function that is not defined
        GS04  Pick action function name differs only in case (JScript / V8)
        GS05  Pick action function could not be verified (library not in the export)
        GS06  Mimic links a library written in a different language
        GS07  Script syntax does not match its language setting
        GS08  Script library not linked by any mimic

    Inline pick-action code (Execute Code Fragment, also on the top-level mimic) is scanned with the
    same rules as the script; findings give its Section (e.g. "PickAction 3"). Function references
    (GS02-GS05) apply to Call Specific Function actions only, and for an action on an embedded-mimic
    object the embedded mimic's script and libraries are searched too.

    RN21 is a 2025 R2 runtime issue found in the field rather than in the release notes:
    GetRegistry("Key") returns Null or an empty string when the key has not been set, so
    converting it (CLng ...) fails. It is an Error when the call is wrapped directly in a
    conversion function and a Warning otherwise.

    RN22 is another field finding: the 2025 R2 VBScript runtime treats an undeclared variable
    called ToString, GetType, GetHashCode or Equals (any case) as the .NET member, so assigning it
    fails with "The object has no suitable property or field named 'ToString'" (Error on each
    use). Dim in the same procedure avoids it (Info, once); a script-level declaration, parameter
    or Sub / Function of that name is untested (Warning, once). The Knowledge Base article
    "Adjustments for VBScript in Geo SCADA 2025 R2" has since confirmed it.

    RN23-RN25 come from that article. RN23 treats a variable Set from Server.FindObject /
    LookupObject / ThisObject like RN03 treats a Form control. RN24 is not a 2025 R2 change (the
    test never worked) but is listed in the article; it is skipped for IsNull / IsEmpty / IsObject
    when the script or a linked library defines its own function of that name. RN25 looks for a
    date (Now, Date, Time, DateAdd ..., a Form.AddDateTime control's .Value, or a variable set
    from one) written with SetRegistry, and GetRegistry values used as dates without CDate.

    RN26 is a field finding (reported to Schneider): in 2025 R2 a Call Specific Function pick
    action whose function has a parameter with no value in the pick action configuration does not
    run at all, with no error; before R2 the missing parameters were passed as Empty. It compares
    the function's parameters (found as for GS03) with the "name <= value" entries of the
    PickAction header line. Error for VBScript (seen); Warning for JavaScript (untested).

    Severities: Error (fails or misbehaves in 2025 R2), Warning (likely to misbehave - check),
    Review (behaviour or layout changed - look at it), Info (advisory).

    Batching by editable object: an object whose header says EditInPlace: False takes its script
    from EditableSource, so by default only the editable object is scanned and every finding is
    reported once, against the object to edit, with the number of instance copies that one edit
    fixes (InstancesFixed). Each object has a Role:
        Template            inside a template definition; fixing it fixes its instance copies
        OverriddenInstance  an instance holding its own copy of the content - edit it directly;
                            changes to its template do not reach it
        Instance            an instance whose template is not in the export (scanned as is)
        Standalone          not part of a template or instance
    _ScriptAnalysis_Actions.csv lists one row per object to edit; _ScriptAnalysis_Instances.csv
    lists the instance copies each one covers. -IncludeInstances scans every file instead.

    The RN06/RN07 COM checks read this machine's registry (64-bit view). Run the analyser on a
    machine built like your ViewX clients, or use -SkipProgIdCheck.

    Versions: the console, the HTML header and every row of the findings CSV (ExporterVersion,
    FormatVersion, AnalyserVersion) record which exporter version and file format the export was
    made with and which analyser version produced the report. See CHANGELOG.md.

    Export check: before the rules run, the folder is checked against the exporter's manifest
    (_ScriptExport_Manifest.csv) and every file that is opened is checked for a sound header. The
    problems are listed in _ScriptAnalysis_ExportCheck.csv, at the top of the HTML report and on
    the console:
        MissingFile       the manifest says the object was exported but its file is not there
        ExportError       the exporter could not export the object (the manifest's Error rows)
        Malformed         the file starts like an exporter file but its header is broken (no END
                          marker, no FullName, line numbers outside the file); it is not analysed
        DuplicateName     two files have the same FullName in their header
        ManifestMismatch  the file's header names a different object (FullName / ObjectId) from
                          the manifest row for that file
        UnsupportedFormat the file's FormatVersion is newer than this analyser knows
        MixedExports      the files come from more than one export run (ExportedAt) or exporter
                          version / FormatVersion; information only
        NotInManifest     the file is not in the manifest - left over from an earlier export or a
                          different scope; analysed unless -ManifestOnly
        NoManifest        there is no manifest, so completeness cannot be checked
    Instance copies counted from the manifest without being opened are only checked for being
    present. When MissingFile, ExportError or Malformed problems are found the analyser still
    writes its report but ends with exit code 3, so a calling script can tell the report is
    based on an incomplete export.

.PARAMETER InputFolder
    The exporter's output folder (the .txt files and _ScriptExport_Manifest.csv).

.PARAMETER OutputFolder
    Where to write the report. Defaults to -InputFolder.

.PARAMETER ReportName
    Base name of the report files. Defaults to _ScriptAnalysis (=> _ScriptAnalysis.csv / .html).

.PARAMETER MinimumSeverity
    Lowest severity to report: Error, Warning, Review or Info (default).

.PARAMETER IncludeRule
    Only run these rule IDs (e.g. RN01,RN11,GS01).

.PARAMETER ExcludeRule
    Skip these rule IDs.

.PARAMETER IncludeInstances
    Scan instance copies too, instead of only the editable source of each script.

.PARAMETER SkipProgIdCheck
    Do not look COM ProgIDs up in this machine's registry (RN06/RN07 are then not reported).

.PARAMETER NoHtml
    Write the CSV only.

.PARAMETER ManifestOnly
    Analyse only the files listed as exported in the manifest; leave out .txt files left over
    from earlier exports or other scopes (they are still listed as NotInManifest). Without a
    manifest this switch is ignored, with a warning.

.EXAMPLE
    .\GeoSCADA-Script-Analyser.ps1 -InputFolder C:\Export\Scripts

.EXAMPLE
    .\GeoSCADA-Script-Analyser.ps1 -InputFolder C:\Export\Scripts -MinimumSeverity Warning -OutputFolder C:\Export\Report

.NOTES
    DISCLAIMER
    This script is provided "AS IS", without warranty of any kind, express or implied. In no
    event shall the author be liable for any claim, damages or other liability arising from,
    out of or in connection with the script or its use. It reads exported script files and
    writes a report; it does not connect to or change Geo SCADA. Its rules are pattern-based:
    findings can be wrong or incomplete, so review each one and test changes on a
    representative non-production system before applying them. You are responsible for
    compliance with your own change-control and security policies.

    Requires Windows PowerShell 5.1 or later. Rule sources: Geo SCADA Expert 2025 R2 Release
    Notes (88.9742.1), pages 6, 16 and 18-28, and the ViewX script type library.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputFolder,
    [string]$OutputFolder,
    [string]$ReportName = '_ScriptAnalysis',
    [ValidateSet('Error', 'Warning', 'Review', 'Info')][string]$MinimumSeverity = 'Info',
    [string[]]$IncludeRule,
    [string[]]$ExcludeRule,
    [switch]$IncludeInstances,
    [switch]$SkipProgIdCheck,
    [switch]$NoHtml,
    [switch]$ManifestOnly
)

$ErrorActionPreference = 'Stop'
$AnalyserVersion = '1.2.1'      # see CHANGELOG.md; bump with every released change
$SupportedFormat = 2      # newest exporter FormatVersion this analyser understands

# =============================================================================================
#  RULES  -  severity, title, fix and source for each rule. Edit severities here.
# =============================================================================================
$RN = 'Geo SCADA 2025 R2 Release Notes'
$KB = 'Geo SCADA Knowledge Base: Adjustments for VBScript in Geo SCADA 2025 R2 (community.se.com, 2026-10-15)'
$Rules = [ordered]@{
    RN01 = @{ Severity = 'Error';   Title = '.Value(x) used to set a value'
              Fix = 'Assign instead: x.Interface.Value = 3.14. The Value property now takes precedence over the Value method, so .Value(x) raises a script error.'; Ref = "$RN p20 '[variable].Value and Similar Methods'" }
    RN02 = @{ Severity = 'Warning'; Title = 'Return value of a Menu method used'
              Fix = 'Menu.Init, Menu.AddItem and Menu.Separator now return Null (previously an empty string). Do not use the return value.'; Ref = "$RN p20-21 'Menu Method Return Values'" }
    RN03 = @{ Severity = 'Warning'; Title = 'Form control used without .Value'
              Fix = 'Form controls no longer have a default property. Use control.Value (e.g. MsgBox field.Value).'; Ref = "$RN p21 'Form Control Default Properties'" }
    RN04 = @{ Severity = 'Review';  Title = 'Form.AddListBox row count is now exact'
              Fix = 'The rows argument is now honoured exactly; reduce it if the list now overlaps buttons or other items. Empty lists show "No Items".'; Ref = "$RN p21-22 'Behavior Changes to Form.AddListBox'" }
    RN05 = @{ Severity = 'Review';  Title = 'Form.AddQuery row count is now exact'
              Fix = 'The rows argument is now honoured exactly; reduce it if the list now overlaps buttons or other items. Empty results show "No Query Results".'; Ref = "$RN p22-23 'Behavior Changes to Form.AddQuery'" }
    RN06 = @{ Severity = 'Error';   Title = 'COM class is 32-bit only'
              Fix = 'ViewX is now 64-bit and can only create 64-bit (or out-of-process) COM objects. Replace the component with a 64-bit version or remove it.'; Ref = "$RN p16 'ViewX and other Client Applications are now 64-bit'" }
    RN07 = @{ Severity = 'Warning'; Title = 'COM class not registered on this machine'
              Fix = 'Check that a 64-bit version of the component is installed on every ViewX client.'; Ref = "$RN p16 'ViewX and other Client Applications are now 64-bit'" }
    RN08 = @{ Severity = 'Review';  Title = 'COM object created from a computed ProgID'
              Fix = 'The ProgID is not a literal so it cannot be checked. Confirm the class is available as 64-bit on the ViewX clients.'; Ref = "$RN p16 'ViewX and other Client Applications are now 64-bit'" }
    RN09 = @{ Severity = 'Error';   Title = 'Clipboard accessed through Internet Explorer'
              Fix = 'Use the new Clipboard object: Clipboard.GetText(), Clipboard.SetText(text), Clipboard.Clear(), Clipboard.ContainsText().'; Ref = "$RN p23 'New Clipboard Object'" }
    RN10 = @{ Severity = 'Error';   Title = 'ScriptControl COM object'
              Fix = 'ScriptControl is 32-bit and not recommended. Replace MsgBox emulation with the Form object (JScript) or Window.Alert/Confirm/Prompt (V8).'; Ref = "$RN p23 'JScript Execution Issues - Calling Script Functions via COM'" }
    RN11 = @{ Severity = 'Error';   Title = 'Wrong case for a Geo SCADA name'
              Fix = 'JScript (and V8) are now case-sensitive for Geo SCADA objects, methods and properties. Use the exact name shown.'; Ref = "$RN p24 'JScript Execution Issues - Case Sensitivity'" }
    RN12 = @{ Severity = 'Error';   Title = 'Number assigned to a Boolean property'
              Fix = 'Use true / false, e.g. ok.Default = true;'; Ref = "$RN p24 'JScript Execution Issues - Type Conversions'" }
    RN13 = @{ Severity = 'Warning'; Title = 'JScript pick action passes arguments'
              Fix = 'JScript9Legacy (default on Windows 11 24H2) cannot pass parameters from a pick action. Remove the parameters, convert the script to JavaScript V8, or (at your own risk) set JScriptReplacement=0.'; Ref = "$RN p24-25 'JScript Execution Issues - Functions with Arguments'" }
    RN14 = @{ Severity = 'Info';    Title = 'JScript in use'
              Fix = 'Schneider Electric does not recommend JScript with Geo SCADA. Consider converting to JavaScript V8. On Windows Server 2025, JScript mimic scripts do not run after switching from Design to Run mode.'; Ref = "$RN p25, p58" }
    RN15 = @{ Severity = 'Error';   Title = 'COM object in JavaScript V8'
              Fix = 'CreateObject / ActiveXObject are not available in JavaScript V8.'; Ref = "$RN p26-27 'Translation Notes'" }
    RN16 = @{ Severity = 'Error';   Title = 'MsgBox / InputBox in JavaScript V8'
              Fix = 'Use Window.Alert, Window.Confirm or Window.Prompt, or the Form object.'; Ref = "$RN p26 'New Window API Methods for JavaScript V8'" }
    RN17 = @{ Severity = 'Error';   Title = 'LocalTimeToUTC / UTCToLocalTime in JavaScript V8'
              Fix = 'Not supported in V8. Use the JavaScript Date object local/UTC methods.'; Ref = "$RN p27 'LocalTimeToUTC and UTCToLocalTime'" }
    RN18 = @{ Severity = 'Review';  Title = 'Form layout in JavaScript V8 uses pixels'
              Fix = 'Form positions and sizes in V8 are pixel units, not text units. Check the Form.Add... coordinates.'; Ref = "$RN p27 'Translation Notes'" }
    RN19 = @{ Severity = 'Info';    Title = 'JavaScript V8 script'
              Fix = 'Runs only on Geo SCADA 2025 R2 or later ViewX / Virtual ViewX clients; not supported by the Original WebX client.'; Ref = "$RN p19, p25-26" }
    RN20 = @{ Severity = 'Info';    Title = 'Debugger statement'
              Fix = 'The Microsoft Script Debugger can no longer be used with VBScript / JScript. Remove the statement (a V8 debugger is available in the script editor).'; Ref = "$RN p6, p19 'No VBScript or JScript Debugger'" }
    RN21 = @{ Severity = 'Warning'; Title = 'GetRegistry without a default value'
              Fix = 'Pass a default, e.g. CLng(GetRegistry("Key", 0)). A key that has not been set returns Null / an empty string, which fails when converted (CLng, CInt, CDbl ...).'; Ref = '2025 R2 script runtime (field experience; not in the release notes)' }
    RN22 = @{ Severity = 'Error';   Title = 'VBScript name clashes with a .NET object member'
              Fix = 'Rename the variable (e.g. toString -> lastSite), or at least Dim it in the procedure. The 2025 R2 VBScript runtime resolves an undeclared ToString, GetType, GetHashCode or Equals to the .NET member, so assigning it raises "The object has no suitable property or field named ''ToString''".'; Ref = "$KB 'Using .Net reserved words toString, GetType, Equals and GetHashCode' (first found in the field)" }
    RN23 = @{ Severity = 'Warning'; Title = 'Database object used as a value (default property)'
              Fix = 'Name the property: CStr(obj.Id), MsgBox obj.FullName. A database object variable (Server.FindObject / LookupObject / ThisObject) no longer has a default property, so CStr(obj) no longer gives the object Id.'; Ref = "$KB 'Object Default Properties'" }
    RN24 = @{ Severity = 'Warning'; Title = 'Object not found tested without Is Nothing'
              Fix = 'Use: If obj Is Nothing Then. Server.FindObject / LookupObject return Nothing when the object is not found, so IsNull(obj), IsEmpty(obj) and obj = Null are never true and IsObject(obj) is always true. Not new in 2025 R2: the test never worked, so the not-found case fails later with "Object required".'; Ref = "$KB 'Server.LookupObject returns Nothing if the object is not found, not Null'" }
    RN25 = @{ Severity = 'Warning'; Title = 'Date stored in or read from the registry without conversion'
              Fix = 'Convert explicitly: SetRegistry "Key", CLng(Int(CDbl(dateValue))) to store it as a Long, and CDate(GetRegistry("Key", 0)) when reading it back. Dates and numbers are no longer converted implicitly.'; Ref = "$KB 'No implicit date/time to number conversions or vice versa'" }
    RN26 = @{ Severity = 'Error';   Title = 'Pick action leaves function parameters without a value'
              Fix = 'Give every parameter a value in the pick action configuration (an empty string '''' or 0 if it is not used), or remove the unused parameters from the function. In 2025 R2 the function does not run at all - with no error message - when any parameter has no value; before R2 the missing ones were passed as Empty.'; Ref = '2025 R2 script runtime (found in the field, reported to Schneider 2026-10; Resource Centre update expected)' }
    GS01 = @{ Severity = 'Error';   Title = 'Linked script library does not exist'
              Fix = 'Remove the link or point it at the library''s current location.'; Ref = 'Exporter header LibraryMissing' }
    GS02 = @{ Severity = 'Error';   Title = 'Function pick action but the mimic has no script'
              Fix = 'Add the function (or a link to a library defining it), or remove the pick action.'; Ref = 'Exporter header PickAction' }
    GS03 = @{ Severity = 'Error';   Title = 'Pick action function not defined'
              Fix = 'Define the function in the mimic script or a linked library, or correct the pick action.'; Ref = 'Exporter header PickAction' }
    GS04 = @{ Severity = 'Error';   Title = 'Pick action function name differs in case'
              Fix = 'JScript / V8 function names are case-sensitive. Match the pick action to the function name exactly.'; Ref = 'Exporter header PickAction' }
    GS05 = @{ Severity = 'Info';    Title = 'Pick action function could not be verified'
              Fix = 'A linked library is not in this export (out of scope?), so the function could not be checked.'; Ref = 'Exporter header PickAction' }
    GS06 = @{ Severity = 'Error';   Title = 'Linked library is in a different language'
              Fix = 'A mimic and the libraries it links must use the same language.'; Ref = 'Exporter headers Language / Library' }
    GS07 = @{ Severity = 'Warning'; Title = 'Script syntax does not match its language'
              Fix = 'The code looks like a different language from the one set on the object. Correct the language setting or the code.'; Ref = 'Script body' }
    GS08 = @{ Severity = 'Info';    Title = 'Script library not linked by any mimic'
              Fix = 'Possibly unused (only mimics are counted; libraries can also be used by other means). Consider removing it.'; Ref = 'Exporter header ReferencedByMimics' }
}
$SeverityRank = @{ Error = 0; Warning = 1; Review = 2; Info = 3 }
# ======================================  end of rules  ======================================

# Geo SCADA script API names (ViewX script type library, ViewXCore.dll, Geo SCADA 2025 R2) used by
# the case-sensitivity and Boolean rules. Members of the script objects (Server, Mimic, Form ...)
# and of mimic drawing objects, plus the V8 Window and Clipboard objects.
$ApiMembers = @'
Add,AddCheckBox,AddComboBox,AddDateTime,AddEditBox,AddGroupBox,AddItem,AddListBox,AddMenu,AddMenuItem,AddMenuSeparator,AddPushButton,
AddQuery,AddRadioButton,AddSeparator,AddSlider,AddStaticText,Alarms,Align,BackColour,CategoryDesc,ChangePassword,CheckAccess,Checked,
Children,Class,Clear,Close,ColumnCount,ColumnNames,Columns,Condition,ConfirmPassword,Connect,Controls,Cookie,Count,Default,DefaultFilter,
Delete,DeselectAndScrollToTop,Disabled,Disconnect,Display,DisplayHelp,DisplayObjectMenu,Editable,Error,ErrorMessage,Exit,FindChild,
FindObject,Format,FullName,GetFilterForSystem,GetOPCValue,GetParameter,GetRegistry,GoBack,GoForward,Group,Head,Height,HomePage,Id,Init,
Insert,Interface,IsGroup,IsSelected,Item,Layers,LeftText,Lines,Load,LocalTimeToUTC,Logoff,LogOffUser,Logon,LogOnUser,LookupObject,
Maximum,MaxLength,Message,Minimum,Multiple,Name,Navigate,NavigateOnHead,Object,OpenInset,OpenWindow,OpenWindowOnHead,Parameters,
Parent,Password,Print,Query,ReadOnly,Remove,RootObject,RowCount,Rows,ScrollToTopEnabled,SelectedAlarms,SelectedRow,Selection,
SelectText,Server,SetBusy,SetFilterForSystem,SetOPCValue,SetParamControl,SetRange,SetRegistry,SetSelected,Severity,SeverityDesc,Show,
ShowKeypad,Sleep,Sort,Source,State,SubCondition,System,SystemName,Text,ThisObject,TickFrequency,Time,Title,Translate,UserName,
UTCToLocalTime,ValidatePassword,ValidateUser,Value,Visible,VScroll,WebLinkWithinViewX,Width,WordWrap,WorkstationName,X,Y,
AddAction,AddButton,AddEllipse,AddEmbeddedAlarmList,AddEmbeddedGraph,AddEmbeddedList,AddEmbeddedMimic,AddHover,AddImage,AddLine,
AddMap,AddOLE,AddPieChart,AddPipe,AddPolyline,AddRectangle,AddRemoteImage,AddSVGFile,AddTag,AddText,Aggregate,AggregateId,AlarmBell,
AlarmBellSounding,AlarmLimitCount,AlarmLimitLabelPos,Alignment,Alpha,AltParam,Animations,AnnotationSize,AnnotationStyle,AutoTicks,
BackAlpha,BackgroundColour,BarWidth,Blink,Border,BringToFront,ButtonFill,ButtonFont,ButtonPen,Colour,Comment,Content,Continuous,
Control,CurrentColIdx,CurrentRowId,CurrentRowIdx,CurrentRowName,CurrentRowSource,DashCap,DashStyle,DataMode,DataPoints,Definition,
EnableColSize,EnableFilter,EnableSort,EndCap,EndCapSize,EndTime,Expression,ExtendedSource,ExtendToNow,Extension,Family,FastUpdates,
Figures,Fill,FillAlpha,FillColour,FillMode,Filter,Fixed,FlashAlpha,FlashColour,FlashFillAlpha,FlashFillColour,FlashMode,FlashOffset,
Font,GapInterval,GetAltArguments,GetArguments,GradAlpha,GradAngle,GradColour,Graph,GridColour,GridSize,H,Hatch,HeaderFont,Hint,
HorzAlignment,Hover,ImageAlignment,Index,Interval,Invert,JoinStyle,Jump,KeyColour,KeyFont,KeyPos,KeyStyle,KeyVisible,Label,LabelExpr,
LabelFont,LabelPos,LabelStyle,Latched,Line,LineColour,LineStyle,LineWidth,Logarithmic,MainFont,MaintainAspectRatio,MajorTicks,
MarkerColour,MarkerSize,MarkerStyle,MaxZoom,MimeType,MinorTicks,MinZoom,Mode,Moniker,Move,Multiplier,NextInterval,Offset,
OperatorControlled,Orientation,Param,Path,Pen,PickComment,PickDisabled,PickMenu,PickParam,PickParamAlt,PickType,PipeFill,PipeWidth,
PivotType,Point,PointCount,Position,PosMax,PosMin,PosVal,PrevInterval,Property,RefreshRate,Resize,Rotate,Rotation,Save,ScaleMaximum,
ScaleMinimum,ScaleType,Segments,Selected,SendToBack,Shared,ShowGrid,ShowKey,ShowMarkers,ShowTitle,SizeDir,SizeMax,SizeMin,SizeVal,
Sources,Split,Sql,StartCap,StartTime,States,Status,Style,Tag,TextColour,TextFill,TextFont,TextPen,TimeAxis,TitleColour,TitleExpr,
TitleFont,TitleVisible,TooltipText,Traces,TraceStyle,Type,UpdateMode,UpdateTime,Valid,VertAlignment,W,Warnings,X1,X2,XAxis,Y1,Y2,
YAxes,YAxis,Zoom,ZoomIn,ZoomOut,Alert,Confirm,Prompt,ContainsText,GetText,SetText
'@ -split '[,\s]+' | Where-Object { $_ }
$BoolProperties = @'
Checked,Default,Disabled,Editable,Group,LeftText,Multiple,Password,ReadOnly,SelectText,ShowKeypad,Sort,Visible,VScroll,WebLinkWithinViewX,
WordWrap,AlarmBell,AutoTicks,Continuous,EnableColSize,EnableFilter,EnableSort,ExtendToNow,FastUpdates,Fixed,Invert,Jump,KeyVisible,
Latched,Logarithmic,MaintainAspectRatio,OperatorControlled,PickDisabled,Selected,Shared,ShowGrid,ShowKey,ShowMarkers,ShowTitle,
TitleVisible
'@ -split '[,\s]+' | Where-Object { $_ }
$GlobalObjects   = @('Server', 'Mimic', 'Form', 'Menu', 'App', 'Window', 'Clipboard')
$GlobalFunctions = @('GetRegistry', 'SetRegistry', 'System', 'Sleep', 'DisplayObjectMenu', 'LocalTimeToUTC', 'UTCToLocalTime', 'SetBusy', 'Translate')
# Members that return Geo SCADA objects: a variable assigned from a chain ending in one of these is
# treated as a Geo SCADA object for the case check.
$ObjectReturning = @('Interface', 'FindObject', 'LookupObject', 'RootObject', 'ThisObject', 'Parent', 'Layers', 'Item', 'Query', 'Controls',
                     'Children', 'FindChild', 'AddMenu', 'Object', 'Server', 'AddCheckBox', 'AddComboBox', 'AddDateTime', 'AddEditBox',
                     'AddGroupBox', 'AddListBox', 'AddPushButton', 'AddQuery', 'AddRadioButton', 'AddSlider', 'AddStaticText')
# Lower-case JavaScript built-ins that clash with API names; not flagged on variables (only on the globals).
$JsBuiltinClash = @('add', 'clear', 'delete', 'filter', 'index', 'insert', 'item', 'move', 'remove', 'sort', 'split', 'name', 'type',
                    'value', 'text', 'count', 'error', 'message', 'source', 'time', 'state', 'status', 'style', 'line', 'position')

# --- Lookup sets ----------------------------------------------------------
$ApiByLower = @{}; foreach ($n in $ApiMembers + $GlobalObjects + $GlobalFunctions) { $ApiByLower[$n.ToLowerInvariant()] = $n }
$GlobalByLower = @{}; foreach ($n in $GlobalObjects) { $GlobalByLower[$n.ToLowerInvariant()] = $n }
$GlobalFnByLower = @{}; foreach ($n in $GlobalFunctions) { $GlobalFnByLower[$n.ToLowerInvariant()] = $n }
$ObjectReturningSet = New-Object 'System.Collections.Generic.HashSet[string]' ([string[]]$ObjectReturning), ([StringComparer]::Ordinal)
$JsBuiltinClashSet  = New-Object 'System.Collections.Generic.HashSet[string]' ([string[]]$JsBuiltinClash), ([StringComparer]::Ordinal)
$BoolPattern = '\.\s*(' + (($BoolProperties | ForEach-Object { [regex]::Escape($_) }) -join '|') + ')\s*=(?!=)\s*(-?\d+)\b'

# --- Text scanning helpers (compiled once) ---------------------------------
if (-not ([System.Management.Automation.PSTypeName]'GeoScadaScriptScan').Type) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
public sealed class GeoScadaExportHeader {
    public string Comment; public int End;
    public int FileLines;      // lines in the whole file, counted as an editor does
    public string Problem;     // set when the file starts like an exporter file but the header is broken
    public Dictionary<string, List<string>> Values = new Dictionary<string, List<string>>(StringComparer.Ordinal);
    public Dictionary<string, List<int>> Lines = new Dictionary<string, List<int>>(StringComparer.Ordinal);
}
public static class GeoScadaScriptScan {
    const string Begin = " ===== GEOSCADA-SCRIPT-EXPORT BEGIN =====";
    // Header of an exporter file: lines "<comment> Key: Value" between the BEGIN and END markers.
    // Returns null when the file does not start with a BEGIN marker (not an exporter file). Reads on to
    // the end of the file to count its lines, so the header's line numbers can be checked.
    public static GeoScadaExportHeader ReadHeader(string path) {
        using (var sr = new StreamReader(path, Encoding.UTF8)) {
            string first = sr.ReadLine();
            if (first == null) return null;
            int b = first.IndexOf(Begin, StringComparison.Ordinal);
            if (b <= 0 || first.Substring(0, b).IndexOf(' ') >= 0) return null;
            var h = new GeoScadaExportHeader { Comment = first.Substring(0, b) };
            int n = 1; string line;
            while ((line = sr.ReadLine()) != null) {
                n++;
                if (line.IndexOf("GEOSCADA-SCRIPT-EXPORT END", StringComparison.Ordinal) >= 0) {
                    h.End = n;
                    // Lines as an editor counts them (and as the exporter numbers them): a script that ends
                    // with a line break has an empty last line.
                    string rest = sr.ReadToEnd();
                    int breaks = 0; foreach (char ch in rest) if (ch == '\n') breaks++;
                    h.FileLines = n + breaks + (rest.Length > 0 ? 1 : 0);
                    return h;
                }
                string t = line.StartsWith(h.Comment, StringComparison.Ordinal) ? line.Substring(h.Comment.Length) : line;
                if (t.StartsWith(" ")) t = t.Substring(1);
                string k, v; int i = t.IndexOf(": ", StringComparison.Ordinal);
                if (i > 0) { k = t.Substring(0, i); v = t.Substring(i + 2); }
                else if (t.EndsWith(":")) { k = t.TrimEnd(':'); v = ""; }
                else continue;
                List<string> vl; if (!h.Values.TryGetValue(k, out vl)) { vl = new List<string>(); h.Values[k] = vl; h.Lines[k] = new List<int>(); }
                vl.Add(v); h.Lines[k].Add(n);
            }
            h.FileLines = n;
            h.Problem = "no header END marker (file truncated or not written completely?)";
            return h;
        }
    }
    // Lines startLine..endLine (1-based, inclusive) joined with CRLF; endLine < 1 means to the end of the file.
    public static string ReadBody(string path, int startLine, int endLine) {
        string s = File.ReadAllText(path, Encoding.UTF8);
        int pos = 0;
        for (int line = 1; line < startLine; line++) { int nl = s.IndexOf('\n', pos); if (nl < 0) return ""; pos = nl + 1; }
        if (endLine < 1) return s.Substring(pos);
        if (endLine < startLine) return "";
        int end = pos;
        for (int line = startLine; line <= endLine; line++) { int nl = s.IndexOf('\n', end); if (nl < 0) { end = s.Length; break; } end = nl + 1; }
        string r = s.Substring(pos, end - pos);
        return r.EndsWith("\r\n") ? r.Substring(0, r.Length - 2) : r.TrimEnd('\n');
    }
    // Replaces comments (and, if blankStrings, the contents of string literals) with spaces.
    // Length and line breaks are kept, so positions map straight back to the original text.
    public static string MaskVbs(string s, bool blankStrings) {
        var b = new StringBuilder(s);
        bool inStr = false; bool lineStart = true;
        for (int i = 0; i < s.Length; i++) {
            char c = s[i];
            if (c == '\r' || c == '\n') { inStr = false; lineStart = true; continue; }
            if (inStr) {
                if (c == '"') { if (i + 1 < s.Length && s[i + 1] == '"') { if (blankStrings) { b[i] = ' '; b[i + 1] = ' '; } i++; continue; } inStr = false; continue; }
                if (blankStrings) b[i] = ' ';
                continue;
            }
            if (c == '"') { inStr = true; lineStart = false; continue; }
            bool rem = lineStart && (c == 'R' || c == 'r') && i + 2 < s.Length && string.Compare(s, i, "rem", 0, 3, true) == 0 &&
                       (i + 3 >= s.Length || char.IsWhiteSpace(s[i + 3]));
            if (c == '\'' || rem) { while (i < s.Length && s[i] != '\r' && s[i] != '\n') { b[i] = ' '; i++; } i--; continue; }
            if (c == ':') { lineStart = true; continue; }
            if (!char.IsWhiteSpace(c)) lineStart = false;
        }
        return b.ToString();
    }
    // Keywords after which a '/' starts a regular expression literal rather than a division.
    static readonly HashSet<string> RegexKeywords = new HashSet<string>(StringComparer.Ordinal) {
        "return", "typeof", "case", "do", "else", "in", "instanceof", "new", "delete", "void", "throw", "yield", "await", "of" };
    public static string MaskJs(string s, bool blankStrings) {
        var b = new StringBuilder(s);
        char prev = '\0';   // previous significant character outside comments/strings
        string prevWord = null;   // the identifier that ended at prev, when prev ends one
        for (int i = 0; i < s.Length; i++) {
            char c = s[i];
            if (char.IsLetter(c) || c == '_' || c == '$') {
                int k = i + 1; while (k < s.Length && (char.IsLetterOrDigit(s[k]) || s[k] == '_' || s[k] == '$')) k++;
                prevWord = s.Substring(i, k - i); prev = 'a'; i = k - 1; continue;
            }
            if (c == '/' && i + 1 < s.Length && s[i + 1] == '/') { while (i < s.Length && s[i] != '\r' && s[i] != '\n') { b[i] = ' '; i++; } i--; continue; }
            if (c == '/' && i + 1 < s.Length && s[i + 1] == '*') {
                int e = s.IndexOf("*/", i + 2); if (e < 0) e = s.Length - 2;
                for (int k = i; k < e + 2 && k < s.Length; k++) if (s[k] != '\r' && s[k] != '\n') b[k] = ' ';
                i = e + 1; continue;
            }
            if (c == '"' || c == '\'' || c == '`') {
                int k = i + 1;
                while (k < s.Length && s[k] != c) {
                    if (s[k] == '\\') { if (blankStrings) b[k] = ' '; k++; if (k < s.Length && blankStrings && s[k] != '\r' && s[k] != '\n') b[k] = ' '; k++; continue; }
                    if ((s[k] == '\r' || s[k] == '\n') && c != '`') break;
                    if (blankStrings && s[k] != '\r' && s[k] != '\n') b[k] = ' ';
                    k++;
                }
                i = k; prev = 'a'; prevWord = null; continue;
            }
            if (c == '/' && (prev == '\0' || "(,=:[!&|?{};+-*%<>~^".IndexOf(prev) >= 0 ||
                             (prevWord != null && RegexKeywords.Contains(prevWord)))) {   // regular expression literal
                int k = i + 1; bool cls = false;
                while (k < s.Length && s[k] != '\r' && s[k] != '\n') {
                    if (s[k] == '\\') { if (blankStrings) b[k] = ' '; k++; if (k < s.Length && blankStrings && s[k] != '\r' && s[k] != '\n') b[k] = ' '; k++; continue; }
                    if (s[k] == '[') cls = true; else if (s[k] == ']') cls = false; else if (s[k] == '/' && !cls) break;
                    if (blankStrings) b[k] = ' ';
                    k++;
                }
                i = k; prev = 'a'; prevWord = null; continue;
            }
            if (!char.IsWhiteSpace(c)) { prev = c; prevWord = null; }
        }
        return b.ToString();
    }
    public static int MatchParen(string s, int open) {
        int d = 0;
        for (int i = open; i < s.Length; i++) { if (s[i] == '(') d++; else if (s[i] == ')') { d--; if (d == 0) return i; } }
        return -1;
    }
    public static int CountArgs(string s, int open, int close) {
        bool any = false; int n = 1, d = 0;
        for (int i = open + 1; i < close; i++) {
            char c = s[i];
            if (c == '(' || c == '[' || c == '{') d++; else if (c == ')' || c == ']' || c == '}') d--;
            else if (c == ',' && d == 0) n++;
            if (!char.IsWhiteSpace(c)) any = true;
        }
        return any ? n : 0;
    }
    // Member-access chains such as Mimic.Layers("x").Item("y").Sql: one entry per member, as
    // { root, rootIndex, member, memberIndex, position-in-chain }. Call arguments are skipped
    // (their own chains are reported separately).
    public static List<object[]> Chains(string s) {
        var toks = new List<int[]>(); // { kind, start, length }  kind: 0 ident, 1 dot, 2 '(', 3 ')', 4 other
        for (int i = 0; i < s.Length; i++) {
            char c = s[i];
            if (char.IsWhiteSpace(c)) continue;
            if (char.IsLetter(c) || c == '_' || c == '$') { int k = i + 1; while (k < s.Length && (char.IsLetterOrDigit(s[k]) || s[k] == '_' || s[k] == '$')) k++; toks.Add(new[] { 0, i, k - i }); i = k - 1; continue; }
            if (char.IsDigit(c)) { int k = i + 1; while (k < s.Length && (char.IsLetterOrDigit(s[k]) || s[k] == '.')) k++; toks.Add(new[] { 4, i, k - i }); i = k - 1; continue; }
            toks.Add(new[] { c == '.' ? 1 : c == '(' ? 2 : c == ')' ? 3 : 4, i, 1 });
        }
        var res = new List<object[]>();
        for (int t = 0; t < toks.Count; t++) {
            if (toks[t][0] != 0 || (t > 0 && toks[t - 1][0] == 1)) continue;
            string root = s.Substring(toks[t][1], toks[t][2]);
            int j = t + 1, pos = 0;
            while (j < toks.Count) {
                if (toks[j][0] == 1 && j + 1 < toks.Count && toks[j + 1][0] == 0) {
                    res.Add(new object[] { root, toks[t][1], s.Substring(toks[j + 1][1], toks[j + 1][2]), toks[j + 1][1], pos++ });
                    j += 2;
                } else if (toks[j][0] == 2) {
                    int d = 0; while (j < toks.Count) { if (toks[j][0] == 2) d++; else if (toks[j][0] == 3) { d--; if (d == 0) break; } j++; }
                    j++;
                } else break;
            }
        }
        return res;
    }
}
'@
}

# --- Helpers ---------------------------------------------------------------
$script:ProgressClock = [System.Diagnostics.Stopwatch]::StartNew()
function Show-Progress([string]$Activity, [string]$Status, [int]$Percent = -1) {
    <# At most one Write-Progress every 2 s of work: in some hosts each call itself takes seconds. #>
    if ($script:ProgressClock.ElapsedMilliseconds -lt 2000) { return }
    if ($Percent -ge 0) { Write-Progress -Activity $Activity -Status $Status -PercentComplete $Percent }
    else { Write-Progress -Activity $Activity -Status $Status }
    $script:ProgressClock.Restart()      # after the call, so a slow call does not count as work time
}

function Get-LongPath([string]$Path) { if ($Path.Length -ge 248 -and -not $Path.StartsWith('\\')) { return '\\?\' + $Path } return $Path }

# Spreadsheet formula-injection guard for the CSV files. Names, script lines and messages come from the
# database; a text cell starting with = + - @ (or tab / CR) would be evaluated by Excel, so it is written
# with a leading ' (shown as text). The same convention is used in the exporter's manifest, which
# ConvertFrom-SafeCsvValue undoes when the manifest is read.
$script:CsvTrigger = "=+-@`t`r"
function ConvertTo-SafeCsvRow {
    process {
        $row = $_; $copy = $null
        foreach ($p in $row.PSObject.Properties) {
            $v = $p.Value
            if ($v -is [string] -and $v.Length -and $script:CsvTrigger.IndexOf($v[0]) -ge 0) {
                if (-not $copy) { $copy = $row.PSObject.Copy() }    # the original keeps its value (used by the HTML)
                $copy.($p.Name) = "'" + $v
            }
        }
        if ($copy) { $copy } else { $row }
    }
}
function ConvertFrom-SafeCsvValue([string]$Value) {
    if ($Value.Length -gt 1 -and $Value[0] -eq "'" -and $script:CsvTrigger.IndexOf($Value[1]) -ge 0) { return $Value.Substring(1) }
    return $Value
}
function Export-SafeCsv([string]$Path, [object[]]$Rows, [string[]]$Columns) {
    <# Export-Csv with the formula guard; an empty list still gets a header line (and replaces an older file). #>
    if ($Rows.Count) { $Rows | ConvertTo-SafeCsvRow | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 }
    else { [System.IO.File]::WriteAllText($Path, ((@($Columns) | ForEach-Object { '"' + $_ + '"' }) -join ',') + "`r`n", (New-Object System.Text.UTF8Encoding($true))) }
}

function Read-ExportHeader([string]$Path) {
    <# Reads only the header: an object with Comment, Values (key -> list), Lines (key -> file line numbers) and End, or $null. #>
    return [GeoScadaScriptScan]::ReadHeader((Get-LongPath $Path))
}
function HV($Header, [string]$Key) { if ($Header.Values.ContainsKey($Key)) { return $Header.Values[$Key][0] } return '' }
function HL($Header, [string]$Key) { if ($Header.Values.ContainsKey($Key)) { return @($Header.Values[$Key]) } return @() }

function Read-ScriptBody($Rec) {
    <# Returns the script text: lines ScriptStartLine..ScriptEndLine (format 2), or to the end of the file (format 1). #>
    $start = [int]$Rec.ScriptStartLine
    if ($start -lt 1) { $start = $Rec.HeaderEnd + 1 }
    $end = if ($Rec.Header.Values.ContainsKey('ScriptEndLine')) { [int](HV $Rec.Header 'ScriptEndLine') } else { 0 }
    if ($Rec.Header.Values.ContainsKey('ScriptEndLine') -and $end -lt $start) { return '' }
    return [GeoScadaScriptScan]::ReadBody((Get-LongPath $Rec.Path), $start, $end)
}

function Split-PickAction([string]$Value, [int]$FormatVersion) {
    <#
      Format 2: N | Where | Action | Type | Context | Function | Args   (Function is "(code)" for inline code)
      Format 1: Where | Action | Type | Function | Args
    #>
    if ($FormatVersion -ge 2) {
        $p = $Value -split ' \| ', 7
        return [pscustomobject]@{ N = $p[0]; Where = $p[1]; Action = $p[2]; Type = $p[3]; Context = $p[4]
                                  Function = $(if ($p.Count -ge 6) { $p[5].Trim() } else { '' }); Args = $(if ($p.Count -ge 7) { $p[6].Trim() } else { '(none)' }) }
    }
    $p = $Value -split ' \| ', 5
    return [pscustomobject]@{ N = ''; Where = $p[0]; Action = $p[1]; Type = $p[2]; Context = ''
                              Function = $(if ($p.Count -ge 4) { $p[3].Trim() } else { '' }); Args = $(if ($p.Count -ge 5) { $p[4].Trim() } else { '(none)' }) }
}

function Test-IsJs([string]$Language) { return $Language -in @('JScript', 'JScript9', 'JavaScriptV8', 'Unknown') }

function Get-ParamNames([string]$List) {
    <# "ByVal a, b = 1, ...rest" -> a, b, rest (parameter names only). #>
    return , ([string[]]@($List -split ',' | ForEach-Object { ($_ -replace '(?i)^\s*(ByVal|ByRef)\s+', '' -replace '=.*$', '' -replace '^\s*\.\.\.', '').Trim() } | Where-Object { $_ }))
}

function Get-FunctionDefs([string]$Code, [string]$Language) {
    <# Function names defined in a script -> hashtable name -> parameter names. Code should have comments masked.
       JavaScript names are case-sensitive, so the table is too (VBScript names are stored lower-case). #>
    $defs = New-Object System.Collections.Hashtable ([StringComparer]::Ordinal)
    if (Test-IsJs $Language) {
        foreach ($m in [regex]::Matches($Code, '\bfunction\s+([A-Za-z_$][\w$]*)\s*\(([^)]*)\)')) { $defs[$m.Groups[1].Value] = Get-ParamNames $m.Groups[2].Value }
        foreach ($m in [regex]::Matches($Code, '\b(?:var|let|const)?\s*([A-Za-z_$][\w$]*)\s*=\s*function\s*\(([^)]*)\)')) { $defs[$m.Groups[1].Value] = Get-ParamNames $m.Groups[2].Value }
        foreach ($m in [regex]::Matches($Code, '\b(?:var|let|const)\s+([A-Za-z_$][\w$]*)\s*=\s*(?:async\s*)?\(([^)]*)\)\s*=>')) { $defs[$m.Groups[1].Value] = Get-ParamNames $m.Groups[2].Value }
    } else {
        foreach ($m in [regex]::Matches($Code, '(?im)^[ \t]*(?:(?:Public|Private)\s+)?(?:Default\s+)?(?:Sub|Function)\s+([A-Za-z_]\w*)\s*(?:\(([^)]*)\))?')) {
            $defs[$m.Groups[1].Value.ToLowerInvariant()] = Get-ParamNames $m.Groups[2].Value
        }
    }
    return $defs
}

$script:ProgIdCache = @{}
function Get-ProgIdStatus([string]$ProgId) {
    <# 'OK' (64-bit in-process or out-of-process server), '32Only', or 'NotRegistered' - from this machine's registry. #>
    if ($script:ProgIdCache.ContainsKey($ProgId)) { return $script:ProgIdCache[$ProgId] }
    $status = 'NotRegistered'
    $views = @{}
    foreach ($v in 'Registry64', 'Registry32') { $views[$v] = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]$v) }
    try {
        $clsid = @{}
        foreach ($v in $views.Keys) {
            $hk = $views[$v]; $id = $null
            $k = $hk.OpenSubKey("$ProgId\CLSID")
            if (-not $k) { $cv = $hk.OpenSubKey("$ProgId\CurVer"); if ($cv) { $k = $hk.OpenSubKey("$($cv.GetValue(''))\CLSID") } }
            if ($k) { $id = [string]$k.GetValue('') }
            $clsid[$v] = $id
        }
        $has = { param($v, $sub) $id = $clsid[$v]; if (-not $id) { return $false }; return [bool]$views[$v].OpenSubKey("CLSID\$id\$sub") }
        if ((& $has 'Registry64' 'InprocServer32') -or (& $has 'Registry64' 'LocalServer32') -or (& $has 'Registry32' 'LocalServer32')) { $status = 'OK' }
        elseif (& $has 'Registry32' 'InprocServer32') { $status = '32Only' }
    } finally { foreach ($v in $views.Values) { $v.Dispose() } }
    $script:ProgIdCache[$ProgId] = $status
    return $status
}

# --- Findings ----------------------------------------------------------------
$findings = New-Object System.Collections.Generic.List[object]
$minRank = $SeverityRank[$MinimumSeverity]
function Test-RuleOn([string]$Id) {
    if ($IncludeRule -and $IncludeRule -notcontains $Id) { return $false }
    if ($ExcludeRule -and $ExcludeRule -contains $Id) { return $false }
    return $true
}
function Add-Finding {
    param([string]$Rule, $Rec, [int]$ScriptLine, [int]$FileLine, [string]$Code, [string]$Detail, [string]$Severity, [string]$Section = 'Header')
    if (-not (Test-RuleOn $Rule)) { return }
    $r = $Rules[$Rule]
    $sev = if ($Severity) { $Severity } else { $r.Severity }
    if ($SeverityRank[$sev] -gt $minRank) { return }
    if ($Code -and $Code.Length -gt 240) { $Code = $Code.Substring(0, 237) + '...' }
    $findings.Add([pscustomobject][ordered]@{
        Severity = $sev; Rule = $Rule; Title = $r.Title; FullName = $Rec.FullName; Class = $Rec.Class; Language = $Rec.Language
        Section = $Section
        ScriptLine = $(if ($ScriptLine -gt 0) { $ScriptLine } else { '' }); FileLine = $(if ($FileLine -gt 0) { $FileLine } else { '' })
        Code = $Code; Detail = $Detail; Fix = $r.Fix; Reference = $r.Ref; Role = $Rec.Role; InstancesFixed = ([int]$Rec.Affected - 1); AffectedObjects = $Rec.Affected
        File = $Rec.File; ObjectId = $Rec.ObjectId
        ExporterVersion = $Rec.ExporterVersion; FormatVersion = $Rec.FormatVersion; AnalyserVersion = $AnalyserVersion
    })
}

# Per-script context used by the body rules. A script is either the object's script ('Script') or the
# inline code of a pick action ('PickAction N'); FirstLine is the file line holding its first line.
function New-ScanContext($Rec, [string]$Body, [int]$FirstLine = 0, [string]$Section = 'Script', [string]$Language) {
    if (-not $Language) { $Language = $Rec.Language }
    if ($Language -eq 'None') { $Language = 'VBScript' }      # no script element: ViewX default
    if ($FirstLine -lt 1) { $FirstLine = $Rec.ScriptStartLine }
    $isJs = Test-IsJs $Language
    $noComments = if ($isJs) { [GeoScadaScriptScan]::MaskJs($Body, $false) } else { [GeoScadaScriptScan]::MaskVbs($Body, $false) }
    $codeOnly   = if ($isJs) { [GeoScadaScriptScan]::MaskJs($Body, $true) }  else { [GeoScadaScriptScan]::MaskVbs($Body, $true) }
    $starts = New-Object System.Collections.Generic.List[int]; $starts.Add(0)
    for ($i = 0; $i -lt $Body.Length; $i++) { if ($Body[$i] -eq "`n") { $starts.Add($i + 1) } }
    return @{ Rec = $Rec; Body = $Body; IsJs = $isJs; NoComments = $noComments; Code = $codeOnly; Starts = $starts.ToArray()
              Lines = $Body -split "`r`n"; FirstLine = $FirstLine; Section = $Section; Language = $Language }
}
function Get-LineOf($Ctx, [int]$Pos) {
    $i = [Array]::BinarySearch($Ctx.Starts, $Pos)
    if ($i -lt 0) { $i = (-bnot $i) - 1 }
    return $i + 1
}
function Add-BodyFinding([string]$Rule, $Ctx, [int]$Pos, [string]$Detail, [string]$Severity) {
    $ln = Get-LineOf $Ctx $Pos
    $code = if ($ln -le $Ctx.Lines.Count) { $Ctx.Lines[$ln - 1].Trim() } else { '' }
    Add-Finding -Rule $Rule -Rec $Ctx.Rec -ScriptLine $ln -FileLine ($Ctx.FirstLine + $ln - 1) -Code $code -Detail $Detail -Severity $Severity -Section $Ctx.Section
}

# True when the member call whose '.' is at $Dot and whose argument list ends at $Close is a whole
# statement (x.Interface.Value(3.14), Call x.Value(1), If a Then x.Value(1)), not part of an
# expression (y = x.Value(0), MsgBox x.Value(0)). $Code is the comment- and string-masked script.
function Test-StatementCall([string]$Code, [int]$Dot, [int]$Close, [bool]$IsJs) {
    $isIdent = { param($c) [char]::IsLetterOrDigit($c) -or $c -eq '_' -or $c -eq '$' }
    # walk back over the member chain (x, x.y, f(a).y, a[0].y) to its first character
    $start = $Dot; $p = $Dot
    while ($true) {
        $q = $p - 1
        while ($q -ge 0 -and ($Code[$q] -eq ' ' -or $Code[$q] -eq "`t")) { $q-- }
        while ($q -ge 0 -and ($Code[$q] -eq ')' -or $Code[$q] -eq ']')) {
            $d = 0
            for (; $q -ge 0; $q--) {
                if ($Code[$q] -eq ')' -or $Code[$q] -eq ']') { $d++ } elseif ($Code[$q] -eq '(' -or $Code[$q] -eq '[') { $d--; if ($d -eq 0) { break } }
            }
            $q--
            while ($q -ge 0 -and ($Code[$q] -eq ' ' -or $Code[$q] -eq "`t")) { $q-- }
        }
        if ($q -lt 0 -or -not (& $isIdent $Code[$q])) { $start = $p; break }    # leading '.' (With block)
        while ($q -ge 0 -and (& $isIdent $Code[$q])) { $q-- }
        $start = $q + 1
        while ($q -ge 0 -and ($Code[$q] -eq ' ' -or $Code[$q] -eq "`t")) { $q-- }
        if ($q -ge 0 -and $Code[$q] -eq '.') { $p = $q; continue }
        break
    }
    # what comes before the chain must start a statement
    $pre = $Code.Substring(0, $start).TrimEnd(' ', "`t")
    $startsStatement = $false
    if ($pre.Length -eq 0) { $startsStatement = $true }
    elseif ($pre[-1] -eq "`n" -or $pre[-1] -eq "`r") { $startsStatement = -not (-not $IsJs -and $pre.TrimEnd() -match '(?:^|[ \t])_$') }  # VBScript line continuation
    elseif ($IsJs) { $startsStatement = ($pre[-1] -in ';', '{', '}', ')') -or $pre -match '\b(?:else|do)$' }
    else { $startsStatement = $pre[-1] -eq ':' -or $pre -match '(?i)\b(?:Then|Else|Call)$' }
    if (-not $startsStatement) { return $false }
    # and the call must end it (an '=' after it is an assignment with an argument - also a set)
    $post = $Code.Substring($Close + 1).TrimStart(' ', "`t")
    if ($post.Length -eq 0 -or $post[0] -eq "`r" -or $post[0] -eq "`n") { return $true }
    if ($IsJs) { return $post[0] -in ';', '}' -or ($post[0] -eq '=' -and $post -notmatch '^==') }
    return $post[0] -eq ':' -or $post[0] -eq '=' -or $post -match '(?i)^Else\b'
}

# VBScript procedures (Start / End / Name / Kind) and declared names in the masked script. Each name is
# @{ Name (lower case); Pos; Proc (Start of its procedure, -1 for script level); Kind; Desc }, Kind being
# Local (Dim / ReDim / Const in a procedure), Parameter, Script (script-level declaration) or Procedure.
function Get-VbsDeclarations([string]$Code) {
    $procs = New-Object System.Collections.Generic.List[object]
    $names = New-Object System.Collections.Generic.List[object]
    $open = $null
    $procPat = '(?im)^[ \t]*(?:(?:Public|Private)[ \t]+)?(?:Default[ \t]+)?(Sub|Function|Property[ \t]+(?:Get|Let|Set))[ \t]+(\w+)[ \t]*(?:\(([^)\r\n]*)\))?|^[ \t]*End[ \t]+(?:Sub|Function|Property)\b'
    foreach ($m in [regex]::Matches($Code, $procPat)) {
        if ($m.Groups[1].Success) {
            $kind = ($m.Groups[1].Value -split '[ \t]+')[0]
            $open = @{ Start = $m.Index; End = $Code.Length; Name = $m.Groups[2].Value; Kind = $kind }
            $procs.Add($open)
            $names.Add(@{ Name = $m.Groups[2].Value.ToLowerInvariant(); Pos = $m.Groups[2].Index; Proc = -1; Kind = 'Procedure'; Desc = "the name of a $kind" })
            if ($m.Groups[3].Success) {
                foreach ($pm in [regex]::Matches($m.Groups[3].Value, '(?i)(?:^|,)\s*(?:(?:ByVal|ByRef)\s+)?(\w+)')) {
                    $names.Add(@{ Name = $pm.Groups[1].Value.ToLowerInvariant(); Pos = $m.Groups[3].Index + $pm.Groups[1].Index; Proc = $open.Start
                                  Kind = 'Parameter'; Desc = "a parameter of $kind $($open.Name)" })
                }
            }
        } elseif ($open) { $open.End = $m.Index + $m.Length; $open = $null }
    }
    foreach ($m in [regex]::Matches($Code, '(?im)(?:^|:)[ \t]*(Dim|ReDim(?:[ \t]+Preserve)?|Const|Private|Public)[ \t]+([^\r\n:]*)')) {
        $kw = $m.Groups[1].Value -replace '[ \t]+', ' '; $list = $m.Groups[2].Value; $at = $m.Groups[2].Index
        if ($list -match '(?i)^(?:Sub|Function|Property|Default)\b') { continue }
        if ($list -match '(?i)^Const[ \t]+') { $kw = "$kw Const"; $at += $Matches[0].Length; $list = $list.Substring($Matches[0].Length) }
        $list = [regex]::Replace($list, '\([^)]*\)', { param($x) ' ' * $x.Length })      # array bounds
        $proc = $procs | Where-Object { $m.Index -ge $_.Start -and $m.Index -le $_.End } | Select-Object -First 1
        foreach ($nm in [regex]::Matches($list, '(?:^|,)\s*(\w+)')) {
            $n = @{ Name = $nm.Groups[1].Value.ToLowerInvariant(); Pos = $at + $nm.Groups[1].Index }
            if ($proc) { $n.Proc = $proc.Start; $n.Kind = 'Local'; $n.Desc = "declared with $kw in $($proc.Kind) $($proc.Name)" }
            else { $n.Proc = -1; $n.Kind = 'Script'; $n.Desc = "declared at script level with $kw" }
            $names.Add($n)
        }
    }
    return @{ Procs = $procs; Names = $names }
}

# VBScript object variables: one entry per variable and procedure (script level: the whole script) for each
# match of $Pattern, whose group 1 is the variable. $Describe gets the match and returns a description of
# what the variable holds, or $null to ignore that match. Scope is the procedure holding the Set.
function Get-SetVariables([string]$Code, [string]$Pattern, $Procs, [scriptblock]$Describe) {
    $vars = @{}
    foreach ($m in [regex]::Matches($Code, $Pattern)) {
        $what = & $Describe $m
        if (-not $what) { continue }
        $proc = $Procs | Where-Object { $m.Index -ge $_.Start -and $m.Index -le $_.End } | Select-Object -First 1
        $from = if ($proc) { $proc.Start } else { 0 }; $to = if ($proc) { $proc.End } else { $Code.Length }
        $key = "$($m.Groups[1].Value.ToLowerInvariant())|$from"
        if (-not $vars.ContainsKey($key)) { $vars[$key] = @{ Name = $m.Groups[1].Value; What = $what; From = $from; To = $to } }
    }
    return @($vars.Values)
}

# Built-in functions and keywords that use the value of their argument (so an object passed to them is
# read through its default property); passing an object to the script's own sub or function is fine.
$script:ValueFunctions = @('MsgBox', 'CStr', 'CInt', 'CLng', 'CDbl', 'CSng', 'CBool', 'CDate', 'CByte', 'CCur', 'Len', 'Trim', 'LTrim', 'RTrim',
    'UCase', 'LCase', 'Left', 'Right', 'Mid', 'InStr', 'InStrRev', 'IsEmpty', 'IsNull', 'IsNumeric', 'IsDate', 'Replace', 'Split', 'Val', 'Abs',
    'Int', 'Fix', 'Round', 'FormatNumber', 'FormatDateTime', 'StrComp', 'Hex', 'Oct', 'Asc', 'Chr',
    'If', 'ElseIf', 'While', 'Until', 'Case', 'Not', 'And', 'Or', 'Xor', 'Mod', 'Then', 'Return')

# Positions in the masked VBScript $Code where object variable $Var (from Get-SetVariables) is used as a
# value - x Mod 5, If x > 1, CStr(x), MsgBox x, y = x, x = 5 ... - rather than as the object itself:
# x.Member, x(...), Set y = x, x Is Nothing, IsObject(x) / TypeName(x) / VarType(x), Dim x, With x, or
# passing x to the script's own sub or function. $Exclude lists value functions not to count (another
# rule reports them); with -SkipNullCompare, "x = Null" / "x <> Null" are not counted either.
function Find-ValueUses([string]$Code, $Var, [string[]]$Exclude = @(), [switch]$SkipNullCompare) {
    $fns = '^(?:' + (($script:ValueFunctions | Where-Object { $_ -notin $Exclude }) -join '|') + ')$'
    $found = New-Object System.Collections.Generic.List[int]
    foreach ($m in [regex]::new("(?i)(?<![\w.])$([regex]::Escape($Var.Name))(?!\w)").Matches($Code, $Var.From)) {
        if ($m.Index -ge $Var.To) { break }
        $i = $m.Index; $j = $i + $m.Length
        while ($j -lt $Code.Length -and ($Code[$j] -eq ' ' -or $Code[$j] -eq "`t")) { $j++ }
        $next = if ($j -lt $Code.Length) { $Code[$j] } else { "`n" }
        if ($next -eq '.' -or $next -eq '(') { continue }                         # x.Value, x(...)
        if ($j + 2 -le $Code.Length -and [string]::Compare($Code, $j, 'Is', 0, 2, $true) -eq 0 -and
            ($j + 2 -eq $Code.Length -or -not [char]::IsLetterOrDigit($Code[$j + 2]))) { continue }   # x Is Nothing
        if ($SkipNullCompare -and $Code.Substring($j, [Math]::Min(12, $Code.Length - $j)) -match '(?i)^(=|<>)\s*Null\b') { continue }
        $ls = [Math]::Max($Code.LastIndexOf("`n", [Math]::Max(0, $i - 1)), $Code.LastIndexOf(':', [Math]::Max(0, $i - 1))) + 1
        if ($ls -gt $i) { $ls = 0 }
        $pre = $Code.Substring($ls, $i - $ls)
        if ($pre -match '(?i)^\s*(?:Dim|ReDim|Private|Public|Const|With|Erase)\b' -or $pre -match '(?i)^\s*Set\s+(?:[\w.]+\s*=\s*)?$' -or
            $pre -match '(?i)\bIs\s*$' -or $pre -match '(?i)\b(?:IsObject|TypeName|VarType)\s*\(\s*$') { continue }
        if ($SkipNullCompare -and $pre -match '(?i)\bNull\s*(=|<>)\s*$') { continue }
        # a whole argument of a call: only a problem for the built-in value functions
        $pt = $pre.TrimEnd(); $prev = if ($pt.Length) { $pt[-1] } else { '' }
        if (($prev -eq '(' -or $prev -eq ',') -and ($next -in ')', ',', "`r", "`n", ':')) {
            $d = 0; $fn = $null
            for ($k = $pt.Length - 1; $k -ge 0; $k--) {
                if ($pt[$k] -eq ')') { $d++ } elseif ($pt[$k] -eq '(') { if ($d -eq 0) { if ($pt.Substring(0, $k) -match '(\w+)\s*$') { $fn = $Matches[1] }; break }; $d-- }
            }
            if (-not $fn -and $pt -match '^\s*(?:Call\s+)?(\w+)') { $fn = $Matches[1] }
            if ($fn -and $fn -notmatch "(?i)$fns") { continue }
        } elseif ($pt -match '^\s*(?:Call\s+)?(\w+)$' -and $Matches[1] -notmatch "(?i)$fns") { continue }   # DoThing x
        $found.Add($i)
    }
    return , $found.ToArray()
}

# Database-object variables: Set x = Server.FindObject(...) / Server.LookupObject(...) / Server.ThisObject,
# optionally followed by .Parent (still a database object). Returns 'Server.FindObject' etc., or $null.
function Get-DbObjectSource($Code, $Match) {
    $p = $Match.Index + $Match.Length; $call = $Match.Groups[2].Value
    if ($call -ine 'ThisObject') {
        while ($p -lt $Code.Length -and ($Code[$p] -eq ' ' -or $Code[$p] -eq "`t")) { $p++ }
        if ($p -ge $Code.Length -or $Code[$p] -ne '(') { return $null }
        $p = [GeoScadaScriptScan]::MatchParen($Code, $p); if ($p -lt 0) { return $null }; $p++
    }
    $rest = $Code.Substring($p, [Math]::Min(200, $Code.Length - $p))
    if ($rest -notmatch '(?i)^(\s*\.\s*Parent\b)*[ \t]*(?:$|[\r\n:]|''|Else\b)') { return $null }   # .Interface etc. is not a database object
    return "Server.$call"
}

# --- Body rules ----------------------------------------------------------------
function Invoke-BodyRules($Ctx) {
    $code = $Ctx.Code; $isJs = $Ctx.IsJs; $lang = $Ctx.Language
    $ci = if ($isJs) { 'None' } else { 'IgnoreCase' }
    $opt = [System.Text.RegularExpressions.RegexOptions]$ci
    $isV8 = $lang -eq 'JavaScriptV8'

    # RN01 .Value(x) used as a statement (a set). Reads such as y = x.Interface.Value(0) still work.
    foreach ($m in [regex]::Matches($code, '\.\s*Value\s*\(', $opt)) {
        $open = $m.Index + $m.Length - 1
        $close = [GeoScadaScriptScan]::MatchParen($code, $open)
        if ($close -lt 0 -or [GeoScadaScriptScan]::CountArgs($code, $open, $close) -ne 1) { continue }
        if (-not (Test-StatementCall $code $m.Index $close $isJs)) { continue }
        $before = $code.Substring([Math]::Max(0, $m.Index - 30), [Math]::Min(30, $m.Index))
        if ($before -match '(?i)Interface\s*$') { Add-BodyFinding 'RN01' $Ctx $m.Index 'Interface.Value(...) with one argument' }
        else { Add-BodyFinding 'RN01' $Ctx $m.Index '.Value(...) with one argument - an error if this object''s Value is a property' 'Warning' }
    }

    # RN02 Menu return value used
    foreach ($m in [regex]::Matches($code, '(?:[=(,&+]|\bIf[ \t]|\breturn[ \t])[ \t]*Menu[ \t]*\.[ \t]*(Init|AddItem|Separator|AddSeparator)\b', $opt)) {
        if ($code.Substring([Math]::Max(0, $m.Index - 1), 2) -match '[=!<>]=') { continue }   # ==, !=, <=, >= comparisons are uses too, but keep simple
        Add-BodyFinding 'RN02' $Ctx $m.Index "Return value of Menu.$($m.Groups[1].Value) used"
    }

    # RN03 Form control default property (VBScript)
    # RN23 (database objects) works the same way: every bare use of the variable reads its default property.
    $vbsProcs = @(); $formVars = @(); $usesRegistry = $false; $hasObjVars = $false
    if (-not $isJs) {
        $hasObjVars = $code -match '(?i)\bSet\s+\w+\s*=\s*(Form|Server)\s*\.'
        $usesRegistry = $code -match '(?i)\b(?:Set|Get)Registry\b'
        if ($hasObjVars -or $usesRegistry) { $vbsProcs = (Get-VbsDeclarations $code).Procs.ToArray() }
        $formVars = Get-SetVariables $code '(?i)\bSet\s+(\w+)\s*=\s*Form\s*\.\s*Add(?!GroupBox)(\w+)' $vbsProcs { param($m) "Form.Add$($m.Groups[2].Value)" }
        $seen = @{}
        foreach ($v in $formVars) {
            foreach ($i in (Find-ValueUses $code $v)) {
                $ln = Get-LineOf $Ctx $i
                if ($seen.ContainsKey($ln)) { continue }; $seen[$ln] = 1
                Add-BodyFinding 'RN03' $Ctx $i "'$($v.Name)' ($($v.What)) used as a value; use $($v.Name).Value"
            }
        }
    }

    # RN23 / RN24 database object variables (VBScript)
    if (-not $isJs -and $hasObjVars) {
        $dbVars =Get-SetVariables $code '(?i)\bSet\s+(\w+)\s*=\s*Server\s*\.\s*(FindObject|LookupObject|ThisObject)\b' $vbsProcs { param($m) Get-DbObjectSource $code $m }
        # RN23: used as a value - no default property any more (IsNull / IsEmpty / = Null are RN24's)
        $seen = @{}
        foreach ($v in $dbVars) {
            foreach ($i in (Find-ValueUses $code $v -Exclude 'IsNull', 'IsEmpty' -SkipNullCompare)) {
                $ln = Get-LineOf $Ctx $i
                if ($seen.ContainsKey($ln)) { continue }; $seen[$ln] = 1
                Add-BodyFinding 'RN23' $Ctx $i "'$($v.Name)' ($($v.What)) used as a value; name the property, e.g. $($v.Name).Id or $($v.Name).FullName"
            }
        }
        # RN24: "not found" tested with IsNull / IsEmpty / = Null / IsObject. FindObject and LookupObject return
        # Nothing, so these tests never see it. Skipped when the script or its libraries define their own function.
        $own = @{}
        foreach ($fn in 'isnull', 'isempty', 'isobject') { if ($code -match "(?im)^[ \t]*(?:(?:Public|Private)\s+)?Function\s+$fn\b") { $own[$fn] = 1 } }
        if ($Ctx.Rec.Class -eq 'CMimic') {
            foreach ($lib in @(HL $Ctx.Rec.Header 'Library')) { $ld = Get-LibraryDefs $lib; if ($ld) { foreach ($fn in 'isnull', 'isempty', 'isobject') { if ($ld.Defs.ContainsKey($fn)) { $own[$fn] = 1 } } } }
        }
        foreach ($v in ($dbVars | Where-Object { $_.What -ne 'Server.ThisObject' })) {
            $e = [regex]::Escape($v.Name)
            foreach ($m in [regex]::new("(?i)\b(IsNull|IsEmpty|IsObject)\s*\(\s*$e\s*\)|(?<![\w.])$e\s*(=|<>)\s*Null\b|\bNull\s*(=|<>)\s*$e(?!\w)").Matches($code, $v.From)) {
                if ($m.Index -ge $v.To) { break }
                $fn = $m.Groups[1].Value
                if ($fn -and $own.ContainsKey($fn.ToLowerInvariant())) { continue }
                $test = if ($fn) { "$fn($($v.Name))" } else { "$($v.Name) = Null" }
                $why = if ($fn -ieq 'IsObject') { 'IsObject is True for Nothing too' } else { "$($v.What) returns Nothing, not Null or Empty" }
                Add-BodyFinding 'RN24' $Ctx $m.Index "$test used to test whether the object was found - $why. Use: If $($v.Name) Is Nothing Then"
            }
        }
    }

    # RN25 implicit date <-> number conversion through the registry (VBScript): SetRegistry with a date value,
    # GetRegistry read straight into a date function or a date-time control. Dates: Now / Date / Time, the date
    # functions, a Form.AddDateTime control's .Value, and variables assigned from those in the same procedure.
    if (-not $isJs -and $usesRegistry) {
        $dtVars = @($formVars | Where-Object { $_.What -eq 'Form.AddDateTime' })
        $dateAlt = '(?:Now|Date|Time|DateAdd|DateSerial|TimeSerial|CDate|DateValue|TimeValue)\b(?!\s*\.)'
        foreach ($v in $dtVars) { $dateAlt += '|' + [regex]::Escape($v.Name) + '\s*\.\s*Value\b' }
        $isDate = "(?i)^\s*(?:$dateAlt)"     # the value starts with a date (Now - 1 is still a date)
        $converted = '(?i)^\s*(?:CDbl|CLng|CInt|CSng|CCur|CStr|Int|Fix|Round|FormatDateTime|Year|Month|Day|Hour|Minute|Second|Weekday|DatePart|DateDiff|Timer)\b'
        # variables holding a date: x = Now / x = DateAdd(...) / x = dt.Value, per procedure
        $dateVars = @{}
        foreach ($m in [regex]::Matches($code, '(?im)^[ \t]*(\w+)[ \t]*=[ \t]*([^\r\n:]+)')) {
            if ($m.Groups[2].Value -match $isDate -and $m.Groups[2].Value -notmatch $converted) {
                $proc = $vbsProcs | Where-Object { $m.Index -ge $_.Start -and $m.Index -le $_.End } | Select-Object -First 1
                $dateVars["$($m.Groups[1].Value.ToLowerInvariant())|$(if ($proc) { $proc.Start } else { 0 })"] = 1
            }
        }
        foreach ($m in [regex]::Matches($code, '(?i)(?<![\w.])(?:App[ \t]*\.[ \t]*)?SetRegistry\b[ \t]*\(?')) {
            # the value is the second argument: from the first top-level comma to the end of the statement
            $p = $m.Index + $m.Length; $d = 0; $comma = -1
            for ($k = $p; $k -lt $code.Length -and $code[$k] -ne "`r" -and $code[$k] -ne "`n"; $k++) {
                $c = $code[$k]
                if ($c -eq '(') { $d++ } elseif ($c -eq ')') { if ($d -eq 0) { break }; $d-- } elseif ($c -eq ':' -and $d -eq 0) { break } elseif ($c -eq ',' -and $d -eq 0) { $comma = $k; break }
            }
            if ($comma -lt 0) { continue }
            $q = $comma + 1; $d = 0
            for ($k = $q; $k -lt $code.Length -and $code[$k] -ne "`r" -and $code[$k] -ne "`n"; $k++) {
                $c = $code[$k]
                if ($c -eq '(') { $d++ } elseif ($c -eq ')') { if ($d -eq 0) { break }; $d-- } elseif (($c -eq ':' -or $c -eq ',') -and $d -eq 0) { break }
            }
            $value = $code.Substring($q, $k - $q).Trim()
            if (-not $value -or $value -match $converted) { continue }
            $proc = $vbsProcs | Where-Object { $m.Index -ge $_.Start -and $m.Index -le $_.End } | Select-Object -First 1
            $scope = if ($proc) { $proc.Start } else { 0 }
            if ($value -match $isDate -or ($value -match '^(\w+)$' -and $dateVars.ContainsKey("$($Matches[1].ToLowerInvariant())|$scope"))) {
                Add-BodyFinding 'RN25' $Ctx $m.Index "SetRegistry with a date value ($value) - store CLng(Int(CDbl($value))) and read it back with CDate"
            }
        }
        foreach ($m in [regex]::Matches($code, '(?i)\b(DateAdd|DateDiff|DatePart|FormatDateTime|Year|Month|Day|Hour|Minute|Second|Weekday)\s*\(([^()]*\([^()]*\))*[^()]*?(?<![\w.])GetRegistry\b')) {
            Add-BodyFinding 'RN25' $Ctx $m.Index "GetRegistry value passed to $($m.Groups[1].Value) - convert it with CDate(GetRegistry(...)) first"
        }
        if ($dtVars.Count) {
            foreach ($v in $dtVars) {
                foreach ($m in [regex]::new("(?i)(?<![\w.])$([regex]::Escape($v.Name))\s*\.\s*Value\s*=\s*GetRegistry\b").Matches($code, $v.From)) {
                    if ($m.Index -ge $v.To) { break }
                    Add-BodyFinding 'RN25' $Ctx $m.Index "GetRegistry value written to a date-time control ($($v.Name).Value) - convert it with CDate first"
                }
            }
        }
    }

    # RN04 / RN05 list and query row counts
    foreach ($m in [regex]::Matches($code, '\bForm\s*\.\s*(AddListBox|AddQuery)\b', $opt)) {
        $rule = if ($m.Groups[1].Value -ieq 'AddListBox') { 'RN04' } else { 'RN05' }
        Add-BodyFinding $rule $Ctx $m.Index "Form.$($m.Groups[1].Value)"
    }

    # RN06-RN10, RN15 COM objects
    $literal = @{}
    $comPattern = if ($isJs) { '\b(?:new\s+)?(ActiveXObject|CreateObject)\s*\(\s*(["''])(.*?)\2' } else { '(?i)\b(CreateObject)\s*\(\s*(")(.*?)"' }
    foreach ($m in [regex]::Matches($Ctx.NoComments, $comPattern)) {
        # Strings are not masked in NoComments, so make sure the call itself is code, not text inside a string.
        $fnGroup = $m.Groups[1]
        if ($code.Substring($fnGroup.Index, $fnGroup.Length) -ne $fnGroup.Value) { continue }
        $literal[$m.Index] = 1
        $progId = $m.Groups[3].Value.Trim()
        $fn = $m.Groups[1].Value
        if ($isV8) { Add-BodyFinding 'RN15' $Ctx $m.Index "$fn(""$progId"")"; continue }
        if ($progId -match '(?i)^InternetExplorer\.Application') { Add-BodyFinding 'RN09' $Ctx $m.Index "$fn(""$progId"")"; continue }
        if ($progId -match '(?i)ScriptControl') { Add-BodyFinding 'RN10' $Ctx $m.Index "$fn(""$progId"")"; continue }
        if ($SkipProgIdCheck) { continue }
        switch (Get-ProgIdStatus $progId) {
            '32Only'        { Add-BodyFinding 'RN06' $Ctx $m.Index "'$progId' is registered only as a 32-bit in-process server on $env:COMPUTERNAME" }
            'NotRegistered' { Add-BodyFinding 'RN07' $Ctx $m.Index "'$progId' is not registered on $env:COMPUTERNAME" }
        }
    }
    foreach ($m in [regex]::Matches($code, '\b(?:new\s+)?(ActiveXObject|CreateObject)\s*\(', $(if ($isJs) { 'None' } else { 'IgnoreCase' }))) {
        if ($literal.ContainsKey($m.Index)) { continue }
        if ($isV8) { Add-BodyFinding 'RN15' $Ctx $m.Index "$($m.Groups[1].Value)(...)" } else { Add-BodyFinding 'RN08' $Ctx $m.Index "$($m.Groups[1].Value) with a computed ProgID" }
    }
    if ($isJs) { foreach ($m in [regex]::Matches($code, '\bclipboardData\b')) { Add-BodyFinding 'RN09' $Ctx $m.Index 'clipboardData' } }

    # RN11 case / RN12 Boolean types (JScript, JScript9, V8)
    if ($isJs) {
        $declared = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($m in [regex]::Matches($code, '\b(?:var|let|const)\s+([A-Za-z_$][\w$]*)')) { [void]$declared.Add($m.Groups[1].Value) }
        foreach ($m in [regex]::Matches($code, '\bfunction\s*([A-Za-z_$][\w$]*)?\s*\(([^)]*)\)')) {
            if ($m.Groups[1].Success) { [void]$declared.Add($m.Groups[1].Value) }
            foreach ($p in ($m.Groups[2].Value -split ',')) { if ($p.Trim()) { [void]$declared.Add($p.Trim()) } }
        }
        $chains = [GeoScadaScriptScan]::Chains($code)
        # members by root position, to find what a variable is assigned from
        $byRoot = @{}
        foreach ($c in $chains) { $k = [int]$c[1]; if (-not $byRoot.ContainsKey($k)) { $byRoot[$k] = New-Object System.Collections.Generic.List[object] }; $byRoot[$k].Add($c) }
        $tracked = New-Object 'System.Collections.Generic.HashSet[string]'
        for ($pass = 0; $pass -lt 3; $pass++) {
            foreach ($m in [regex]::Matches($code, '(?<![\w$.])([A-Za-z_$][\w$]*)\s*=(?![=>])\s*(?:new\s+)?(?=[A-Za-z_$])')) {
                $rhs = $m.Index + $m.Length
                if (-not $byRoot.ContainsKey($rhs)) { continue }
                $mem = $byRoot[$rhs]; $root = [string]$mem[0][0]
                $rootIsGeo = $GlobalObjects -ccontains $root -or $tracked.Contains($root)
                if ($rootIsGeo -and $ObjectReturningSet.Contains([string]$mem[$mem.Count - 1][2])) { [void]$tracked.Add($m.Groups[1].Value) }
            }
        }
        $reported = @{}
        foreach ($c in $chains) {
            $root = [string]$c[0]; $member = [string]$c[2]; $rootPos = [int]$c[1]
            $rl = $root.ToLowerInvariant()
            # A global object written in the wrong case (form.Init) - unless it is the script's own variable.
            if ([int]$c[4] -eq 0 -and $GlobalByLower.ContainsKey($rl) -and $GlobalByLower[$rl] -cne $root -and -not $declared.Contains($root) -and -not $reported.ContainsKey("r$rootPos")) {
                $reported["r$rootPos"] = 1
                Add-BodyFinding 'RN11' $Ctx $rootPos "'$root' should be '$($GlobalByLower[$rl])'"
            }
            $isGlobal = $GlobalByLower.ContainsKey($rl) -and -not $declared.Contains($root) -or $GlobalObjects -ccontains $root
            if (-not ($isGlobal -or $tracked.Contains($root))) { continue }
            $ml = $member.ToLowerInvariant()
            if (-not $ApiByLower.ContainsKey($ml) -or $ApiByLower[$ml] -ceq $member) { continue }
            if (-not $isGlobal -and $JsBuiltinClashSet.Contains($member) -and $member -cne 'value') { continue }
            $pos = [int]$c[3]
            if ($reported.ContainsKey("m$pos")) { continue }; $reported["m$pos"] = 1
            Add-BodyFinding 'RN11' $Ctx $pos "'.$member' should be '.$($ApiByLower[$ml])'"
        }
        # Global functions called in the wrong case (lookUpObject is a member, so covered above)
        foreach ($m in [regex]::Matches($code, '(?<![\w$.])([A-Za-z_$][\w$]*)\s*\(')) {
            $n = $m.Groups[1].Value; $nl = $n.ToLowerInvariant()
            if ($GlobalFnByLower.ContainsKey($nl) -and $GlobalFnByLower[$nl] -cne $n -and -not $declared.Contains($n)) {
                Add-BodyFinding 'RN11' $Ctx $m.Index "'$n' should be '$($GlobalFnByLower[$nl])'"
            }
        }
        foreach ($m in [regex]::Matches($code, $BoolPattern)) {
            $val = $m.Groups[2].Value
            Add-BodyFinding 'RN12' $Ctx $m.Index ".$($m.Groups[1].Value) = $val; use $(if ($val -eq '0') { 'false' } else { 'true' })"
        }
    }

    # RN16 / RN17 / RN18 JavaScript V8
    if ($isV8) {
        foreach ($m in [regex]::Matches($code, '(?<![\w$.])(MsgBox|InputBox)\s*\(', 'IgnoreCase')) { Add-BodyFinding 'RN16' $Ctx $m.Index "$($m.Groups[1].Value)(...)" }
        foreach ($m in [regex]::Matches($code, '\b(LocalTimeToUTC|UTCToLocalTime)\s*\(', 'IgnoreCase')) { Add-BodyFinding 'RN17' $Ctx $m.Index "$($m.Groups[1].Value)(...)" }
        foreach ($m in [regex]::Matches($code, '\bForm\s*\.\s*Init\s*\(')) { Add-BodyFinding 'RN18' $Ctx $m.Index 'Form layout' }
    }

    # RN21 GetRegistry(key) without a default: Null / empty when the key is not set
    foreach ($m in [regex]::Matches($code, '(?<![\w$])(?:App\s*\.\s*)?GetRegistry\s*\(', $opt)) {
        $open = $m.Index + $m.Length - 1
        $close = [GeoScadaScriptScan]::MatchParen($code, $open)
        if ($close -lt 0 -or [GeoScadaScriptScan]::CountArgs($code, $open, $close) -ne 1) { continue }
        $before = $code.Substring([Math]::Max(0, $m.Index - 40), [Math]::Min(40, $m.Index))
        if ($before -match '(?i)\b(CLng|CInt|CDbl|CSng|CBool|CDate|CByte|CCur|Int|Fix|Abs|DateValue|TimeValue|parseInt|parseFloat|Number)\s*\(\s*$') {
            Add-BodyFinding 'RN21' $Ctx $m.Index "GetRegistry with no default, converted with $($Matches[1])()" 'Error'
        } else {
            Add-BodyFinding 'RN21' $Ctx $m.Index 'GetRegistry with no default value'
        }
    }

    # RN22 VBScript identifier named like a .NET object member (ToString, GetType, GetHashCode, Equals - all tested).
    # Undeclared: an error on every use. Declared with Dim in the same procedure: works (tested), one Info.
    # Declared at script level, a parameter, or a Sub / Function of that name: untested, one Warning.
    if (-not $isJs) {
        $clash = [regex]::Matches($code, '(?i)(?<![\w.])(ToString|GetType|GetHashCode|Equals)(?!\w)')
        if ($clash.Count) {
            $decl = Get-VbsDeclarations $code
            $reported = @{}
            foreach ($m in $clash) {
                $p = $m.Index - 1
                while ($p -ge 0 -and ($code[$p] -eq ' ' -or $code[$p] -eq "`t")) { $p-- }
                if ($p -ge 0 -and $code[$p] -eq '.') { continue }        # x.ToString - a member call, not a name of our own
                $name = $m.Groups[1].Value; $key = $name.ToLowerInvariant()
                $proc = $decl.Procs | Where-Object { $m.Index -ge $_.Start -and $m.Index -le $_.End } | Select-Object -First 1
                $d = $null
                if ($proc) { $d = $decl.Names | Where-Object { $_.Name -eq $key -and $_.Proc -eq $proc.Start } | Select-Object -First 1 }
                if (-not $d) { $d = $decl.Names | Where-Object { $_.Name -eq $key -and $_.Proc -lt 0 } | Select-Object -First 1 }
                if (-not $d) {
                    Add-BodyFinding 'RN22' $Ctx $m.Index "'$name' used as a variable without a declaration" 'Error'
                    continue
                }
                $once = "$key|$($d.Pos)"
                if ($reported.ContainsKey($once)) { continue }
                $reported[$once] = 1
                if ($d.Kind -eq 'Local') { Add-BodyFinding 'RN22' $Ctx $d.Pos "'$name' is $($d.Desc) - works in 2025 R2, rename when convenient" 'Info' }
                else { Add-BodyFinding 'RN22' $Ctx $d.Pos "'$name' is $($d.Desc) - untested in 2025 R2" 'Warning' }
            }
        }
    }

    # RN20 debugger statements
    $dbg = if ($isJs) { '\bdebugger\b' } else { '(?im)^[ \t]*Stop[ \t]*\r?$' }
    foreach ($m in [regex]::Matches($code, $dbg)) { Add-BodyFinding 'RN20' $Ctx $m.Index $(if ($isJs) { 'debugger' } else { 'Stop' }) }

    # GS07 syntax looks like the other language
    $vbs = [regex]::Matches($code, '(?im)^[ \t]*(?:Dim\s+\w|Set\s+\w+\s*=|End\s+(?:Sub|Function|If|Select)\b|(?:Public\s+|Private\s+)?Sub\s+\w|If\b.*\bThen[ \t]*\r?$|Next[ \t]*\r?$|Wend\b)')
    $js  = [regex]::Matches($code, '(?m)(?:\bfunction\s*\w*\s*\(|\b(?:var|let|const)\s+[A-Za-z_$]|;[ \t]*\r?$|^[ \t]*[{}][ \t]*\r?$)')
    if ($isJs -and $vbs.Count -ge 3 -and $vbs.Count -gt $js.Count) {
        Add-BodyFinding 'GS07' $Ctx $vbs[0].Index "Language is $lang but the code looks like VBScript ($($vbs.Count) VBScript statements, $($js.Count) JavaScript)"
    } elseif (-not $isJs -and $js.Count -ge 3 -and $js.Count -gt $vbs.Count) {
        Add-BodyFinding 'GS07' $Ctx $js[0].Index "Language is $lang but the code looks like JavaScript ($($js.Count) JavaScript statements, $($vbs.Count) VBScript)"
    }
}

# --- Main ----------------------------------------------------------------------
$InputFolder = [System.IO.Path]::GetFullPath($InputFolder)
if (-not (Test-Path -LiteralPath $InputFolder -PathType Container)) { throw "Input folder '$InputFolder' not found." }
if ([string]::IsNullOrWhiteSpace($OutputFolder)) { $OutputFolder = $InputFolder }
$OutputFolder = [System.IO.Path]::GetFullPath($OutputFolder)
if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
foreach ($r in @($IncludeRule) + @($ExcludeRule)) { if ($r -and -not $Rules.Contains($r)) { throw "Unknown rule '$r'. Rules: $($Rules.Keys -join ', ')" } }

Write-Host "GeoSCADA-Script-Analyser $AnalyserVersion - reading headers in $InputFolder ..." -ForegroundColor Cyan
$sw = [System.Diagnostics.Stopwatch]::StartNew()

# 1. Read the manifest and the headers and index the files by object full name. When the exporter's
#    manifest is present, instance copies (whose script lives in an exported editable source) are
#    counted from it without opening them - with tens of thousands of files that saves most of the
#    disk reads. The folder is also checked against the manifest (export check, see the help).
$preSkip = @{}; $pendingAffected = @{}; $preSkipped = 0
$sourceOf = @{}       # object full name -> editable source full name (from the manifest), for embedded-mimic lookups
$checks = New-Object System.Collections.Generic.List[object]
function Add-CheckIssue([string]$Issue, [string]$FullName, [string]$File, [string]$Detail) {
    $checks.Add([pscustomobject][ordered]@{ Issue = $Issue; FullName = $FullName; File = $File; Detail = $Detail })
}
$manifestFiles = $null    # file name -> @(FullName, Id, EditableSource) of the manifest's Exported rows
$manifestPath = Join-Path $InputFolder '_ScriptExport_Manifest.csv'
if (Test-Path -LiteralPath $manifestPath) {
    $man = @(Import-Csv -LiteralPath $manifestPath)
    if ($man.Count -and -not ($man[0].PSObject.Properties['Status'] -and $man[0].PSObject.Properties['File'] -and $man[0].PSObject.Properties['FullName'])) {
        throw "'$manifestPath' is not an exporter manifest (no FullName / Status / File columns)."
    }
    $manifestFiles = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
    $hasSource = $man.Count -and $man[0].PSObject.Properties['EditableSource']
    $exported = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($row in $man) {
        # values that looked like spreadsheet formulas were written with a leading ' (see ConvertTo-SafeCsvRow)
        $fullName = $row.FullName; if ($fullName.StartsWith("'")) { $fullName = ConvertFrom-SafeCsvValue $fullName }
        $file = $row.File;         if ($file.StartsWith("'"))     { $file = ConvertFrom-SafeCsvValue $file }
        $src = if ($hasSource) { $row.EditableSource } else { '' }; if ($src.StartsWith("'")) { $src = ConvertFrom-SafeCsvValue $src }
        if ($row.Status -eq 'Error') { Add-CheckIssue 'ExportError' $fullName '' (ConvertFrom-SafeCsvValue $row.Message); continue }
        if ($row.Status -ne 'Exported' -or -not $file) { continue }
        $manifestFiles[$file] = @($fullName, $row.Id, $src)
        [void]$exported.Add($fullName)
    }
    foreach ($kv in $manifestFiles.GetEnumerator()) {
        $fullName = $kv.Value[0]; $src = $kv.Value[2]
        if (-not $src) { continue }
        $sourceOf[$fullName] = $src
        if (-not $IncludeInstances -and $src -ne $fullName -and $exported.Contains($src)) { $preSkip[$kv.Key] = @($src, $fullName) }
    }
    if ($preSkip.Count) { Write-Host "Manifest: $($preSkip.Count) instance copies will be counted without reading them." -ForegroundColor Cyan }
    $man = $null
} else {
    Add-CheckIssue 'NoManifest' '' '' 'No _ScriptExport_Manifest.csv in the folder, so missing files, export errors and left-over files cannot be detected.'
    if ($ManifestOnly) { Write-Warning '-ManifestOnly ignored: there is no manifest.' }
}

# Checks a header read from disk; returns a description of what is wrong, or $null.
function Get-HeaderProblem($Header) {
    if ($Header.Problem) { return $Header.Problem }
    if (-not (HV $Header 'FullName')) { return 'no FullName in the header' }
    $fv = 0; $fvText = HV $Header 'FormatVersion'
    if (-not [int]::TryParse($fvText, [ref]$fv) -or $fv -lt 1) { return "FormatVersion '$fvText' is not valid" }
    $lines = $Header.FileLines
    $startText = HV $Header 'ScriptStartLine'
    if ($startText) {
        $start = 0
        if (-not [int]::TryParse($startText, [ref]$start)) { return "ScriptStartLine '$startText' is not a number" }
        if ($start -le $Header.End) { return "ScriptStartLine $start is inside the header (which ends on line $($Header.End))" }
        if ($Header.Values.ContainsKey('ScriptEndLine')) {
            $end = 0; $endText = HV $Header 'ScriptEndLine'
            if (-not [int]::TryParse($endText, [ref]$end)) { return "ScriptEndLine '$endText' is not a number" }
            if ($end -lt $start - 1 -or $end -gt $lines) { return "ScriptEndLine $end is outside the file ($lines lines, script starts on line $start)" }
        } elseif ($start -gt $lines + 1) { return "ScriptStartLine $start is beyond the end of the file ($lines lines)" }
    }
    foreach ($pc in @(HL $Header 'PickActionCode')) {
        $p = $pc -split ' \| '; $last = 0
        if ($p.Count -lt 3 -or -not [int]::TryParse($p[2], [ref]$last)) { return "PickActionCode '$pc' is not 'n | first | last'" }
        if ($last -gt $lines) { return "PickActionCode '$pc' ends beyond the end of the file ($lines lines)" }
    }
    return $null
}
$records = New-Object System.Collections.Generic.List[object]
$byName  = @{}
$instancesOf = @{}    # editable source full name -> list of instance-copy full names it covers
function Add-InstanceOf([string]$Source, [string]$Instance) {
    if (-not $instancesOf.ContainsKey($Source)) { $instancesOf[$Source] = New-Object System.Collections.Generic.List[string] }
    $instancesOf[$Source].Add($Instance)
}
function Get-Role($Header) {
    <# Template / OverriddenInstance / Instance / Standalone, from the format-2 header (format 1 => Standalone). #>
    if ((HV $Header 'InTemplate') -eq 'True') { return 'Template' }
    if ((HV $Header 'ControlledByInstance') -eq 'True') {
        if ((HV $Header 'DocumentContentOverride') -eq 'True') { return 'OverriddenInstance' }
        return 'Instance'
    }
    return 'Standalone'
}
$skipped = 0; $leftOut = 0
$exportVersions = [ordered]@{}   # "exporter version|format version" -> files read
$exportRuns = @{}                # ExportedAt -> files read (one value per export run)
$onDisk = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$files = [System.IO.Directory]::EnumerateFiles($(if ($InputFolder.StartsWith('\\')) { $InputFolder } else { '\\?\' + $InputFolder }), '*.txt')
$n = 0
foreach ($f in $files) {
    $n++
    if ($n % 500 -eq 1) { Show-Progress 'Reading headers' "$n files" }
    $path = if ($f.StartsWith('\\?\')) { $f.Substring(4) } else { $f }
    $name = [System.IO.Path]::GetFileName($path)
    [void]$onDisk.Add($name)
    $pre = $preSkip[$name]
    if ($pre) {                                     # inlined: this runs once per instance copy (tens of thousands)
        $pendingAffected[$pre[0]] = 1 + [int]$pendingAffected[$pre[0]]
        $il = $instancesOf[$pre[0]]; if ($null -eq $il) { $il = New-Object System.Collections.Generic.List[string]; $instancesOf[$pre[0]] = $il }
        $il.Add($pre[1]); $preSkipped++; continue
    }
    $listed = if ($manifestFiles) { $manifestFiles[$name] } else { $null }
    if ($manifestFiles -and -not $listed -and $ManifestOnly) {
        Add-CheckIssue 'NotInManifest' '' $name 'Not in the manifest; not analysed (-ManifestOnly).'; $leftOut++; continue
    }
    $h = $null
    try { $h = Read-ExportHeader $path } catch { Add-CheckIssue 'Malformed' '' $name "Could not be read: $($_.Exception.Message)"; continue }
    if (-not $h) { $skipped++; continue }           # no BEGIN marker: not an exporter file
    $problem = Get-HeaderProblem $h
    if ($problem) { Add-CheckIssue 'Malformed' (HV $h 'FullName') $name "$problem. Not analysed."; continue }
    $fullName = HV $h 'FullName'
    if ([int](HV $h 'FormatVersion') -gt $SupportedFormat) {
        Add-CheckIssue 'UnsupportedFormat' $fullName $name "FormatVersion $(HV $h 'FormatVersion') is newer than this analyser ($SupportedFormat); findings may be incomplete."
    }
    # Exporter 1.0.0 trimmed header values, so a name ending in a space lost it in the header; the manifest
    # kept it. Use the manifest's name, or the object would not match its instance copies.
    if ($listed -and $listed[0] -ne $fullName -and $listed[0].TrimEnd() -eq $fullName.TrimEnd()) { $fullName = $listed[0] }
    if ($manifestFiles) {
        if (-not $listed) {
            Add-CheckIssue 'NotInManifest' $fullName $name 'Not in the manifest - left over from an earlier export or a different scope. Analysed anyway (use -ManifestOnly to leave it out).'
        } elseif ($listed[0] -ne $fullName -or ($listed[1] -and (HV $h 'ObjectId') -and $listed[1] -ne (HV $h 'ObjectId'))) {
            Add-CheckIssue 'ManifestMismatch' $fullName $name "The manifest lists this file as '$($listed[0])' (Id $($listed[1])); the header says '$fullName' (Id $(HV $h 'ObjectId'))."
        }
    }
    if ($byName.ContainsKey($fullName)) {
        Add-CheckIssue 'DuplicateName' $fullName $name "Same FullName as $($byName[$fullName].File); both are analysed."
    }
    $rec = [pscustomobject]@{
        Path = $path; File = [System.IO.Path]::GetFileName($path); Header = $h; HeaderEnd = $h.End
        FullName = $fullName; Class = HV $h 'Class'; Language = HV $h 'Language'; ObjectId = HV $h 'ObjectId'
        ScriptStartLine = [int](HV $h 'ScriptStartLine'); EditInPlace = (HV $h 'EditInPlace') -ne 'False'
        EditableSource = $(if ($listed -and $listed[2]) { $listed[2] } else { HV $h 'EditableSource' })   # the manifest keeps trailing spaces
        Affected = 1; Scan = $true; Role = (Get-Role $h)
        ExporterVersion = $(if (HV $h 'Exporter') { (HV $h 'Exporter') -replace '^GeoSCADA-Script-Exporter\s*', '' } else { '(unknown)' })
        FormatVersion = HV $h 'FormatVersion'
    }
    # versions and export runs the folder was made with (files counted from the manifest are not opened)
    $vk = "$($rec.ExporterVersion)|$($rec.FormatVersion)"; $exportVersions[$vk] = 1 + [int]$exportVersions[$vk]
    $ek = HV $h 'ExportedAt'; if (-not $ek) { $ek = '(unknown)' }; $exportRuns[$ek] = 1 + [int]$exportRuns[$ek]
    $records.Add($rec)
    if ($rec.FullName) { $byName[$rec.FullName] = $rec }
}
Write-Progress -Activity 'Reading headers' -Completed
if ($manifestFiles) {
    foreach ($kv in $manifestFiles.GetEnumerator()) {
        if (-not $onDisk.Contains($kv.Key)) { Add-CheckIssue 'MissingFile' $kv.Value[0] $kv.Key 'The manifest says this object was exported, but the file is not in the folder. Its script is not in this report.' }
    }
}
if (-not $records.Count) { throw "No exporter files found in '$InputFolder'." }
if ($skipped) { Write-Warning "$skipped .txt file(s) are not exporter files and were ignored." }
# Which exporter version(s), file format(s) and export run(s) the folder holds.
$exportSummary = (@($exportVersions.Keys | ForEach-Object { $p = $_ -split '\|', 2; "GeoSCADA-Script-Exporter $($p[0]) (FormatVersion $($p[1])): $($exportVersions[$_]) file(s)" }) -join '; ')
$runList = @($exportRuns.Keys | Sort-Object)
$exportedAt = if ($runList.Count -eq 1) { $runList[0] } else { "$($runList.Count) export runs, $($runList[0]) to $($runList[-1])" }
Write-Host "Export: $exportSummary; exported $exportedAt$(if ($preSkipped) { " (from the $($records.Count) files read; $preSkipped instance copies counted from the manifest were not opened)" })." -ForegroundColor Cyan
if ($exportVersions.Count -gt 1 -or $runList.Count -gt 1) {
    Add-CheckIssue 'MixedExports' '' '' "The folder holds files from $($runList.Count) export run(s) and $($exportVersions.Count) exporter version / format combination(s): $exportSummary. Runs: $($runList -join ', ')."
}
# Export check summary. MissingFile / ExportError / Malformed mean the report is based on an incomplete export.
$checkCounts = [ordered]@{}
foreach ($c in $checks) { $checkCounts[$c.Issue] = 1 + [int]$checkCounts[$c.Issue] }
$incomplete = @($checks | Where-Object { $_.Issue -in 'MissingFile', 'ExportError', 'Malformed' }).Count
foreach ($k in $checkCounts.Keys) {
    $first = @($checks | Where-Object Issue -eq $k | Select-Object -First 5 | ForEach-Object { if ($_.FullName) { $_.FullName } else { $_.File } })
    $more = if ($checkCounts[$k] -gt $first.Count) { ", ... ($($checkCounts[$k] - $first.Count) more)" } else { '' }
    $what = @{ MissingFile = 'listed in the manifest but missing from the folder'; ExportError = 'could not be exported (manifest Error rows)'
               Malformed = 'have a broken header and were not analysed'; DuplicateName = 'share a FullName with another file'
               ManifestMismatch = 'do not match their manifest row'; UnsupportedFormat = 'use a newer FormatVersion'
               MixedExports = 'files from more than one export run or exporter version'
               NotInManifest = $(if ($ManifestOnly) { 'are not in the manifest and were left out (-ManifestOnly)' } else { 'are not in the manifest (left over from an earlier export or another scope) but were analysed' })
               NoManifest = 'no manifest: completeness not checked' }[$k]
    $msg = if ($k -eq 'NoManifest') { "Export check: $what." }
           elseif ($k -eq 'MixedExports') { "Export check: $(($checks | Where-Object Issue -eq 'MixedExports' | Select-Object -First 1).Detail)" }
           else { "Export check: $($checkCounts[$k]) file(s) $what - $($first -join ', ')$more" }
    if ($k -in 'MissingFile', 'ExportError', 'Malformed') { Write-Warning $msg } else { Write-Host $msg -ForegroundColor Yellow }
}
if ($incomplete) { Write-Warning "EXPORT INCOMPLETE: this report does not cover $incomplete object(s) - see $ReportName`_ExportCheck.csv. Re-export before relying on it." }

Write-Verbose ('{0,-24} {1,7:n1}s' -f 'headers read', $sw.Elapsed.TotalSeconds)
# 2. De-duplicate template instances: scan the editable source, count the copies against it.
$instances = $preSkipped
$orphans = 0
foreach ($src in $pendingAffected.Keys) {
    if ($byName.ContainsKey($src)) { $byName[$src].Affected += $pendingAffected[$src] } else { $orphans += $pendingAffected[$src] }
}
if ($orphans) { Write-Warning "$orphans instance copies refer to an editable source that is no longer in the folder, so they were not scanned. Re-run with -IncludeInstances to scan them." }
if (-not $IncludeInstances) {
    foreach ($rec in $records) {
        if (-not $rec.EditInPlace -and $rec.EditableSource -and $byName.ContainsKey($rec.EditableSource) -and $byName[$rec.EditableSource] -ne $rec) {
            $rec.Scan = $false
            $byName[$rec.EditableSource].Affected++
            Add-InstanceOf $rec.EditableSource $rec.FullName
            $instances++
        }
    }
}
$toScan = @($records | Where-Object Scan)
Write-Host ("{0} files: {1} scanned, {2} instance copies counted against their template." -f $records.Count, $toScan.Count, $instances) -ForegroundColor Cyan

Write-Verbose ('{0,-24} {1,7:n1}s' -f 'de-duplicated', $sw.Elapsed.TotalSeconds)
# 3. Library function tables (for pick-action checks), loaded on demand.
$libFunctions = @{}
function Get-LibraryDefs([string]$Name) {
    if ($libFunctions.ContainsKey($Name)) { return $libFunctions[$Name] }
    $defs = $null
    if ($byName.ContainsKey($Name)) {
        $lr = $byName[$Name]
        $body = Read-ScriptBody $lr
        $masked = if (Test-IsJs $lr.Language) { [GeoScadaScriptScan]::MaskJs($body, $true) } else { [GeoScadaScriptScan]::MaskVbs($body, $true) }
        $defs = @{ Language = $lr.Language; Defs = (Get-FunctionDefs $masked $lr.Language) }
    }
    $libFunctions[$Name] = $defs
    return $defs
}

# Function tables of a mimic (its own script plus its libraries), by full name - used for pick actions
# on an embedded-mimic object, whose function can live in the embedded mimic. Instance copies resolve
# to their editable source. Returns $null when the mimic is not in the export.
$mimicFunctions = @{}
function Get-MimicDefs([string]$Name) {
    if ($mimicFunctions.ContainsKey($Name)) { return $mimicFunctions[$Name] }
    $result = $null
    $srcName = if ($sourceOf.ContainsKey($Name)) { $sourceOf[$Name] } elseif ($byName.ContainsKey($Name) -and $byName[$Name].EditableSource) { $byName[$Name].EditableSource } else { $Name }
    if (-not $byName.ContainsKey($srcName)) { $srcName = $Name }
    if ($byName.ContainsKey($srcName) -and $byName[$srcName].Class -eq 'CMimic') {
        $mr = $byName[$srcName]
        $lang = if ($mr.Language -eq 'None') { 'VBScript' } else { $mr.Language }
        $body = Read-ScriptBody $mr
        $masked = if (Test-IsJs $lang) { [GeoScadaScriptScan]::MaskJs($body, $true) } else { [GeoScadaScriptScan]::MaskVbs($body, $true) }
        $tables = @(@{ Name = $srcName; Defs = (Get-FunctionDefs $masked $lang) })
        foreach ($lib in @(HL $mr.Header 'Library')) { $ld = Get-LibraryDefs $lib; if ($ld) { $tables += @{ Name = $lib; Defs = $ld.Defs } } }
        $result = @{ Language = $lang; Tables = $tables }
    }
    $mimicFunctions[$Name] = $result
    return $result
}

# 4. Scan.
$i = 0
$libCheckNotDone = 0
foreach ($rec in $toScan) {
    $i++
    if ($i % 50 -eq 1 -or $i -eq $toScan.Count) {
        Show-Progress 'Analysing scripts' "$i of $($toScan.Count): $($rec.FullName)" ([int](100 * $i / $toScan.Count))
    }
    $h = $rec.Header
    try { $body = Read-ScriptBody $rec } catch { Write-Warning "Could not read '$($rec.File)': $($_.Exception.Message)"; continue }
    $ctx = New-ScanContext $rec $body

    if ($rec.Class -eq 'CMimic') {
        # GS01 missing libraries
        $missLines = if ($h.Lines.ContainsKey('LibraryMissing')) { $h.Lines['LibraryMissing'] } else { @() }
        $k = 0
        foreach ($lib in @(HL $h 'LibraryMissing')) { Add-Finding -Rule 'GS01' -Rec $rec -FileLine $missLines[$k] -Detail $lib -Code "LibraryMissing: $lib"; $k++ }
        if ((HV $h 'LibraryCheck') -like 'NotPerformed*') { $libCheckNotDone++ }

        # GS06 library language
        $libs = @(HL $h 'Library')
        $libLines = if ($h.Lines.ContainsKey('Library')) { $h.Lines['Library'] } else { @() }
        $k = 0
        foreach ($lib in $libs) {
            $ld = Get-LibraryDefs $lib
            if ($ld -and $rec.Language -ne 'None' -and $ld.Language -and $ld.Language -ne $rec.Language) {
                Add-Finding -Rule 'GS06' -Rec $rec -FileLine $libLines[$k] -Code "Library: $lib" -Detail "Mimic is $($rec.Language); library is $($ld.Language)"
            }
            $k++
        }

        # Pick actions: GS02 / GS03 / GS04 / GS05 / RN13. Only "Call Specific Function" actions (scriptCall)
        # name a function; inline code (scriptCode, scriptTopLevel) is scanned as a script below.
        $acts = @(HL $h 'PickAction')
        $fmt = [int]("0" + (HV $h 'FormatVersion'))
        $embedTarget = @{}
        foreach ($em in @(HL $h 'EmbeddedMimic')) { $ep = $em -split ' \| ', 2; if ($ep.Count -eq 2) { $embedTarget[$ep[0]] = $ep[1] } }
        if ($acts.Count) {
            $actLines = $h.Lines['PickAction']
            $own = Get-FunctionDefs $ctx.Code $rec.Language
            $isJs = Test-IsJs $rec.Language
            for ($a = 0; $a -lt $acts.Count; $a++) {
                $pa = Split-PickAction $acts[$a] $fmt
                if ($pa.Type -in 'scriptCode', 'scriptTopLevel') { continue }      # inline code, not a function reference (both formats record the type)
                # A single call can also carry arguments inline (Name args / Name(args) / Call Name(args));
                # anything else (assignments, several statements) is not a function reference.
                $call = $pa.Function; $fn = ''
                if ($call -match '^(?:Call\s+)?([A-Za-z_$][\w$]*)\s*$') { $fn = $Matches[1] }
                elseif ($call -match '^(?:Call\s+)?([A-Za-z_$][\w$]*)\s*(?:\(|\s+(?!=))' -and $call -notmatch '^(?:Call\s+)?[A-Za-z_$][\w$]*\s*=') { $fn = $Matches[1] }
                $where = $pa.Where
                if (-not $fn) { continue }
                if ($rec.Language -eq 'None') {
                    Add-Finding -Rule 'GS02' -Rec $rec -FileLine $actLines[$a] -Code "PickAction: $($acts[$a])" -Detail "$where calls '$fn'"
                    continue
                }
                $found = $false; $caseOnly = $null; $unverified = @()
                $key = if ($isJs) { $fn } else { $fn.ToLowerInvariant() }
                $tables = @(@{ Name = '(mimic script)'; Defs = $own; Lang = $rec.Language })
                foreach ($lib in $libs) { $ld = Get-LibraryDefs $lib; if ($ld) { $tables += @{ Name = $lib; Defs = $ld.Defs; Lang = $ld.Language } } else { $unverified += $lib } }
                # A pick action set on an embedded-mimic object can call a function of the embedded mimic.
                if ($where -match '\[EmbeddedMimic\]$' -and $embedTarget.ContainsKey($where) -and $embedTarget[$where] -ne '(animated)') {
                    $md = Get-MimicDefs $embedTarget[$where]
                    if ($md) { foreach ($t in $md.Tables) { $tables += @{ Name = $t.Name; Defs = $t.Defs; Lang = $md.Language } } }
                    else { $unverified += $embedTarget[$where] }
                }
                $params = $null; $paramLang = $null
                foreach ($t in $tables) {
                    $tblJs = Test-IsJs $t.Lang          # VBScript tables hold lower-case names; JavaScript ones are case-sensitive
                    $tk = if ($tblJs) { $fn } else { $fn.ToLowerInvariant() }
                    if ($t.Defs.ContainsKey($tk)) { $found = $true; $params = @($t.Defs[$tk]); $paramLang = $t.Lang; break }
                    if ($tblJs -and -not $caseOnly) { foreach ($dn in $t.Defs.Keys) { if ($dn -ieq $fn) { $caseOnly = $dn } } }
                }
                # RN26: every parameter of the called function needs a value in the pick action. The exporter writes
                # each parameter that has one as "name <= value"; a parameter left blank has no entry (or no value).
                # Only for a plain function name - a call written with inline arguments passes them itself.
                if ($found -and $params.Count -and $call -match '^(?:Call\s+)?[A-Za-z_$][\w$]*\s*$') {
                    $given = @{}
                    foreach ($gm in [regex]::Matches([string]$pa.Args, '(?:^|, )([A-Za-z_$][\w$]*) <= ')) { $given[$gm.Groups[1].Value.ToLowerInvariant()] = 1 }
                    $missing = @($params | Where-Object { -not $given.ContainsKey($_.ToLowerInvariant()) })
                    if ($missing.Count) {
                        $vbs = -not (Test-IsJs $paramLang)
                        Add-Finding -Rule 'RN26' -Rec $rec -FileLine $actLines[$a] -Code "PickAction: $($acts[$a])" `
                            -Detail ("$where calls '$fn($($params -join ', '))' with no value for: $($missing -join ', ')" + $(if ($vbs) { '' } else { ' (seen with VBScript; untested for JavaScript)' })) `
                            -Severity $(if ($vbs) { 'Error' } else { 'Warning' })
                    }
                }
                if (-not $found) {
                    $scope = 'the mimic script' + $(if ($libs.Count) { ' or its libraries' } else { '' }) + $(if ($where -match '\[EmbeddedMimic\]$' -and $embedTarget[$where]) { " or the embedded mimic $($embedTarget[$where])" } else { '' })
                    if ($caseOnly) { Add-Finding -Rule 'GS04' -Rec $rec -FileLine $actLines[$a] -Code "PickAction: $($acts[$a])" -Detail "$where calls '$fn' but the function is '$caseOnly'" }
                    elseif ($unverified.Count) { Add-Finding -Rule 'GS05' -Rec $rec -FileLine $actLines[$a] -Code "PickAction: $($acts[$a])" -Detail "'$fn' not in $scope; not exported: $($unverified -join '; ')" }
                    else { Add-Finding -Rule 'GS03' -Rec $rec -FileLine $actLines[$a] -Code "PickAction: $($acts[$a])" -Detail "$where calls '$fn', which is not defined in $scope" }
                }
                if ($rec.Language -in @('JScript', 'JScript9') -and $pa.Args -ne '(none)') {
                    Add-Finding -Rule 'RN13' -Rec $rec -FileLine $actLines[$a] -Code "PickAction: $($acts[$a])" -Detail "$where calls '$fn' with: $($pa.Args)"
                }
            }
        }
    } elseif ($rec.Class -eq 'CScriptLibrary') {
        $refs = HV $h 'ReferencedByMimics'
        if ($refs -eq '0') { Add-Finding -Rule 'GS08' -Rec $rec -FileLine $h.Lines['ReferencedByMimics'][0] -Code "ReferencedByMimics: $refs" -Detail 'No exported mimic links this library' }
    }

    # Per-language notes (one per object)
    if ($rec.Language -in @('JScript', 'JScript9') -and $body.Trim()) { Add-Finding -Rule 'RN14' -Rec $rec -FileLine $h.Lines['Language'][0] -Code "Language: $($rec.Language)" -Detail $rec.Language }
    if ($rec.Language -eq 'JavaScriptV8' -and $body.Trim()) { Add-Finding -Rule 'RN19' -Rec $rec -FileLine $h.Lines['Language'][0] -Code "Language: $($rec.Language)" -Detail 'JavaScript V8' }

    if ($body.Trim()) { Invoke-BodyRules $ctx }

    # Inline pick-action code (format 2): scanned with the same rules, in the mimic's language. Code run
    # on the top-level mimic actually runs in that mimic's language, which is not known here.
    $codeEntries = @(HL $h 'PickActionCode')
    $ctxOf = @{}
    if ($codeEntries.Count) { foreach ($pv in @(HL $h 'PickAction')) { $pp = $pv -split ' \| ', 6; if ($pp.Count -ge 5) { $ctxOf[$pp[0]] = $pp[4] } } }
    foreach ($pc in $codeEntries) {
        $cp = $pc -split ' \| '
        if ($cp.Count -lt 3) { continue }
        $start = [int]$cp[1]; $end = [int]$cp[2]
        if ($end -lt $start) { continue }
        try { $snippet = [GeoScadaScriptScan]::ReadBody((Get-LongPath $rec.Path), $start, $end) } catch { continue }
        if (-not $snippet.Trim()) { continue }
        $section = "PickAction $($cp[0])" + $(if ($ctxOf[$cp[0]] -eq 'TopLevel') { ' (top-level mimic)' } else { '' })
        Invoke-BodyRules (New-ScanContext $rec $snippet -FirstLine $start -Section $section)
    }
}
Write-Progress -Activity 'Analysing scripts' -Completed
Write-Verbose ('{0,-24} {1,7:n1}s' -f 'scripts analysed', $sw.Elapsed.TotalSeconds)

# --- Output --------------------------------------------------------------------
$sorted = @($findings | Sort-Object @{ Expression = { $SeverityRank[$_.Severity] } }, Rule, FullName, @{ Expression = { [int]("0" + $_.FileLine) } })
$csvPath = Join-Path $OutputFolder "$ReportName.csv"
Export-SafeCsv $csvPath $sorted @('Severity', 'Rule', 'Title', 'FullName', 'Class', 'Language', 'Section', 'ScriptLine', 'FileLine', 'Code', 'Detail', 'Fix', 'Reference', 'Role', 'InstancesFixed', 'AffectedObjects', 'File', 'ObjectId', 'ExporterVersion', 'FormatVersion', 'AnalyserVersion')
$checkPath = Join-Path $OutputFolder "$($ReportName)_ExportCheck.csv"
Export-SafeCsv $checkPath $checks.ToArray() @('Issue', 'FullName', 'File', 'Detail')

Write-Verbose ('{0,-24} {1,7:n1}s' -f 'findings CSV written', $sw.Elapsed.TotalSeconds)
# Actions: one row per object to edit, most severe first, then by how many instance copies the edit fixes.
$RoleAction = @{
    Template           = 'Edit this template object; the change reaches its {0} instance copies.'
    OverriddenInstance = 'Edit this instance directly: it holds its own copy of the content, so changes to its template do not reach it.'
    Instance           = 'Instance copy whose template is not in this export: make the change in its template ({1}).'
    Standalone         = 'Edit this object.'
}
$actions = @($sorted | Group-Object FullName | ForEach-Object {
    $g = $_.Group; $first = $g[0]
    $rec = $byName[$first.FullName]
    $worst = ($g | Sort-Object @{ Expression = { $SeverityRank[$_.Severity] } } | Select-Object -First 1).Severity
    $fixes = [int]$first.InstancesFixed
    $role = [string]$first.Role
    $action = [string]::Format($RoleAction[$(if ($RoleAction.ContainsKey($role)) { $role } else { 'Standalone' })], $fixes, $(if ($rec) { $rec.EditableSource } else { '' }))
    if ($role -eq 'Template' -and $fixes -eq 0) { $action = 'Edit this template object (no instance copies of it are in this export).' }
    if ($role -eq 'Instance' -and $rec -and $byName.ContainsKey($rec.EditableSource)) {
        $action = "Instance copy (scanned because of -IncludeInstances): make the change in its template ($($rec.EditableSource))."
    }
    if ($role -eq 'Standalone' -and $fixes -gt 0) { $action = "Edit this object; the change also reaches $fixes copies of it." }
    $linked = if ($rec -and $first.Class -eq 'CScriptLibrary') { HV $rec.Header 'ReferencedByMimics' } else { '' }
    [pscustomobject][ordered]@{
        Priority = 0; Severity = $worst; FullName = $first.FullName; Role = $role; Class = $first.Class; Language = $first.Language
        InstancesFixed = $fixes; LinkedByMimics = $linked; Findings = $g.Count
        Errors   = @($g | Where-Object Severity -eq 'Error').Count;  Warnings = @($g | Where-Object Severity -eq 'Warning').Count
        Reviews  = @($g | Where-Object Severity -eq 'Review').Count; Infos    = @($g | Where-Object Severity -eq 'Info').Count
        Rules    = (($g | Group-Object Rule | Sort-Object Name | ForEach-Object { if ($_.Count -gt 1) { "$($_.Name) x$($_.Count)" } else { $_.Name } }) -join '; ')
        Action   = $action; File = $first.File
    }
} | Sort-Object @{ Expression = { $SeverityRank[$_.Severity] } }, @{ Expression = { $_.InstancesFixed }; Descending = $true }, FullName)
for ($k = 0; $k -lt $actions.Count; $k++) { $actions[$k].Priority = $k + 1 }
$actionsPath = Join-Path $OutputFolder "$($ReportName)_Actions.csv"
Export-SafeCsv $actionsPath $actions @('Priority', 'Severity', 'FullName', 'Role', 'Class', 'Language', 'InstancesFixed', 'LinkedByMimics', 'Findings', 'Errors', 'Warnings', 'Reviews', 'Infos', 'Rules', 'Action', 'File')

Write-Verbose ('{0,-24} {1,7:n1}s' -f 'actions CSV written', $sw.Elapsed.TotalSeconds)
# Instance copies covered by each object to edit (only objects with findings).
$instancesPath = Join-Path $OutputFolder "$($ReportName)_Instances.csv"
$instRows = foreach ($a in $actions) {
    if ($instancesOf.ContainsKey($a.FullName)) {
        foreach ($n in ($instancesOf[$a.FullName] | Sort-Object)) { [pscustomobject][ordered]@{ EditableSource = $a.FullName; Priority = $a.Priority; Instance = $n } }
    }
}
Export-SafeCsv $instancesPath @($instRows) @('EditableSource', 'Priority', 'Instance')

Write-Verbose ('{0,-24} {1,7:n1}s' -f 'instances CSV written', $sw.Elapsed.TotalSeconds)
$byRule = $sorted | Group-Object Rule | ForEach-Object {
    $objs = @($_.Group | Group-Object FullName)
    [pscustomobject]@{ Rule = $_.Name; Severity = $_.Group[0].Severity; Title = $Rules[$_.Name].Title
                       Findings = $_.Count; ObjectsToEdit = $objs.Count
                       ObjectsAffected = ($objs | ForEach-Object { [int]$_.Group[0].AffectedObjects } | Measure-Object -Sum).Sum }
} | Sort-Object @{ Expression = { $SeverityRank[$_.Severity] } }, Rule

if (-not $NoHtml) {
    $htmlPath = Join-Path $OutputFolder "$ReportName.html"
    $meta = [ordered]@{
        generated = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm'); analyser = $AnalyserVersion; folder = $InputFolder
        export = [string]$exportSummary; exportedAt = [string]$exportedAt
        files = $records.Count + $preSkipped; scanned = $toScan.Count; instances = $instances; host = $env:COMPUTERNAME
        progIdCheck = (-not $SkipProgIdCheck); libCheckNotDone = $libCheckNotDone; minSeverity = $MinimumSeverity
        actionsCsv = [System.IO.Path]::GetFileName($actionsPath); instancesCsv = [System.IO.Path]::GetFileName($instancesPath)
        checkCsv = [System.IO.Path]::GetFileName($checkPath); incomplete = $incomplete; manifestOnly = [bool]$ManifestOnly
        # export-check problems: [issue, full name or file, detail], the serious ones first, at most 300
        checks = [object[]]@($checks | Sort-Object @{ Expression = { if ($_.Issue -in 'MissingFile', 'ExportError', 'Malformed') { 0 } else { 1 } } }, Issue |
                     Select-Object -First 300 | ForEach-Object { , [object[]]@([string]$_.Issue, [string]$(if ($_.FullName) { $_.FullName } else { $_.File }), [string]$_.Detail) })
        checkCounts = $(if ($true) { $cc = @{}; foreach ($k in $checkCounts.Keys) { $cc[[string]$k] = [int]$checkCounts[$k] }; $cc })
    }
    $rulesJson = @{}
    foreach ($k in $Rules.Keys) {
        # x: why the rule did not run (shown instead of a count, so 0 always means "checked, nothing found")
        $x = if (-not (Test-RuleOn $k)) { 'Not run: left out by -IncludeRule / -ExcludeRule' }
             elseif ($SkipProgIdCheck -and $k -in 'RN06', 'RN07') { 'Not run: -SkipProgIdCheck' }
             elseif ($SeverityRank[$Rules[$k].Severity] -gt $minRank) { "Not reported: its severity is below -MinimumSeverity $MinimumSeverity" }
             else { '' }
        $rulesJson[$k] = @{ s = $Rules[$k].Severity; t = $Rules[$k].Title; f = $Rules[$k].Fix; r = $Rules[$k].Ref; x = $x }
    }
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($f in $sorted) { $rows.Add([object[]]@([string]$f.Severity, [string]$f.Rule, [string]$f.FullName, [string]$f.Class, [string]$f.Language, [string]$f.ScriptLine, [string]$f.FileLine, [string]$f.Code, [string]$f.Detail, [int]$f.InstancesFixed, [string]$f.File, [string]$f.Section, [string]$f.Role)) }
    $objJson = @{}; $instJson = @{}
    foreach ($a in $actions) {
        $objJson[$a.FullName] = [object[]]@([int]$a.Priority, [string]$a.Role, [int]$a.InstancesFixed, [string]$a.Action, [string]$a.LinkedByMimics)
        if ($instancesOf.ContainsKey($a.FullName)) { $names = [string[]]$instancesOf[$a.FullName].ToArray(); [Array]::Sort($names, [StringComparer]::OrdinalIgnoreCase); $instJson[$a.FullName] = $names }
    }
    Add-Type -AssemblyName System.Web.Extensions
    $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    $ser.MaxJsonLength = [int]::MaxValue
    $json = $ser.Serialize(@{ meta = $meta; rules = $rulesJson; rows = $rows.ToArray(); objs = $objJson; inst = $instJson })
    $json = $json.Replace('</', '<\/')
    $template = @'
<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Geo SCADA Script Analysis</title>
<style>
:root{--bg:#f6f7f9;--fg:#1d2330;--muted:#5b6475;--card:#fff;--line:#dfe3ea;--err:#c62828;--warn:#b26a00;--rev:#1565c0;--info:#546e7a;--code:#f0f2f5;--tmpl:#2e7d32;--ovr:#8e24aa}
@media (prefers-color-scheme:dark){:root{--bg:#14171c;--fg:#e4e7ec;--muted:#9aa3b2;--card:#1d2128;--line:#2e343e;--err:#ef5350;--warn:#ffb74d;--rev:#64b5f6;--info:#90a4ae;--code:#262b33;--tmpl:#81c784;--ovr:#ce93d8}}
*{box-sizing:border-box}body{margin:0;font:14px/1.45 system-ui,Segoe UI,Roboto,sans-serif;background:var(--bg);color:var(--fg)}
header{padding:16px 20px;border-bottom:1px solid var(--line);background:var(--card)}h1{margin:0 0 4px;font-size:20px}.meta{color:var(--muted);font-size:12px}
main{padding:16px 20px;max-width:1500px}.cards{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px}
.card{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:10px 14px;min-width:120px}.card b{font-size:22px;display:block}
.Error{color:var(--err)}.Warning{color:var(--warn)}.Review{color:var(--rev)}.Info{color:var(--info)}.Template{color:var(--tmpl)}.OverriddenInstance{color:var(--ovr)}.Instance,.Standalone{color:var(--muted)}
table{border-collapse:collapse;width:100%;background:var(--card);border:1px solid var(--line);border-radius:8px;overflow:hidden}
th,td{text-align:left;padding:6px 8px;border-bottom:1px solid var(--line);vertical-align:top}th{font-size:12px;color:var(--muted);font-weight:600}
tr.rule:hover{background:var(--code);cursor:pointer}.pill{font-size:11px;font-weight:700;padding:1px 6px;border-radius:9px;border:1px solid currentColor;white-space:nowrap}
.controls{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin:16px 0 10px}.controls input[type=search]{flex:1;min-width:220px;padding:6px 8px;border:1px solid var(--line);border-radius:6px;background:var(--card);color:var(--fg)}
.controls select{padding:6px;border:1px solid var(--line);border-radius:6px;background:var(--card);color:var(--fg)}
details.obj{background:var(--card);border:1px solid var(--line);border-radius:8px;margin:6px 0}details.obj>summary{padding:8px 10px;cursor:pointer;display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.prio{font-weight:700;color:var(--muted);min-width:2.5em}.name{font-weight:600;word-break:break-all}.sub{color:var(--muted);font-size:12px}.f{border-top:1px solid var(--line);padding:8px 10px}
.fixes{font-weight:700;color:var(--tmpl)}.act{padding:8px 10px;border-top:1px solid var(--line);font-weight:600}
details.inst{padding:4px 10px 8px}details.inst ul{margin:4px 0 0;padding-left:20px;columns:2;font-size:12px}
code,pre{font-family:Consolas,ui-monospace,monospace;font-size:12.5px}pre{background:var(--code);padding:6px 8px;border-radius:6px;margin:4px 0;white-space:pre-wrap;word-break:break-all}
.fix{color:var(--muted);font-size:12px}.more{margin:10px 0}button{padding:6px 12px;border:1px solid var(--line);border-radius:6px;background:var(--card);color:var(--fg);cursor:pointer}
.note{background:var(--card);border-left:3px solid var(--warn);padding:8px 12px;margin:10px 0;font-size:13px}
.disclaimer{margin-top:10px;padding:8px 12px;border:1px solid var(--line);border-left:4px solid var(--rev);border-radius:6px;background:var(--bg);font-size:13px;max-width:1500px}
.note.bad{border-left:5px solid var(--err)}.note.bad>b{color:var(--err)}.note ul{margin:4px 0 0;padding-left:20px}.note details{margin-top:6px}.note summary{cursor:pointer;color:var(--muted)}
.chk{font-size:12px;max-height:320px;overflow:auto}.chk td{padding:3px 6px}
.pick{position:relative}.pick>summary{list-style:none;padding:6px 26px 6px 8px;border:1px solid var(--line);border-radius:6px;background:var(--card);cursor:pointer;white-space:nowrap;position:relative}
.pick>summary::-webkit-details-marker{display:none}.pick>summary::after{content:'\25BE';position:absolute;right:8px;color:var(--muted)}
.pickbox{position:absolute;z-index:5;top:100%;right:0;margin-top:4px;background:var(--card);border:1px solid var(--line);border-radius:8px;padding:8px 10px;max-height:60vh;overflow:auto;width:440px;max-width:calc(100vw - 32px);box-shadow:0 6px 18px rgba(0,0,0,.18)}
.pickbtns{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:6px}.pickbtns button{padding:3px 8px;font-size:12px}
.pickbox label{display:flex;gap:6px;align-items:baseline;padding:2px 0;cursor:pointer}.pickbox label b{min-width:3.2em}
tr.rule.off{opacity:.45}tr.rule.zero td:nth-child(n+4),tr.rule.notrun td:nth-child(n+4){color:var(--muted)}.notrunwhy{color:var(--warn);font-size:12px;font-weight:600}
th.sort{cursor:pointer;user-select:none;white-space:nowrap}th.sort:hover{color:var(--fg)}
th.sort::after{content:'\2195';opacity:.35;margin-left:4px}th[aria-sort=ascending]::after{content:'\25B2';opacity:1}th[aria-sort=descending]::after{content:'\25BC';opacity:1}
</style></head><body>
<header><h1>Geo SCADA Script Analysis</h1><div class="meta" id="meta"></div>
<div class="disclaimer" role="note"><b>Disclaimer: guidance only.</b> This report comes from pattern-based checks of the exported script text. It can include false positives, and there is no guarantee that every issue has been identified: some problems only show at run time, or depend on data, configuration or the client machine. Review each finding and test changes on a representative non-production system before applying them. The report is provided "as is", without warranty of any kind.</div></header>
<main>
<div id="checknote"></div>
<div class="cards" id="cards"></div>
<div id="notes"></div>
<h2 style="font-size:16px">By rule</h2>
<table id="rules"><thead><tr><th>Severity</th><th>Rule</th><th>Title</th><th>Findings</th><th>Objects to edit</th><th>Objects affected</th></tr></thead><tbody></tbody></table>
<h2 style="font-size:16px">Objects to edit</h2>
<div class="sub">One entry per object to change, most severe first, then by how many instance copies the change fixes. A template change reaches its instance copies; an overridden instance has to be edited on its own.</div>
<div class="controls">
<input type="search" id="q" placeholder="Search object, instance, code, detail...">
<select id="sev"><option value="">All severities</option><option>Error</option><option>Warning</option><option>Review</option><option>Info</option></select>
<select id="role"><option value="">All roles</option><option value="Template">Template</option><option value="OverriddenInstance">Overridden instance</option><option value="Standalone">Standalone</option><option value="Instance">Instance</option></select>
<details class="pick" id="rulepick"><summary id="rulesum">All rules</summary><div class="pickbox">
<div class="pickbtns"><button type="button" data-p="all">All</button><button type="button" data-p="none">None</button><button type="button" data-p="RN">RN only</button><button type="button" data-p="GS">GS only</button></div>
<div id="rulelist"></div></div></details>
<select id="lang"><option value="">All languages</option></select>
<span class="sub" id="count"></span>
</div>
<div id="list"></div>
<div class="more"><button id="morebtn" hidden>Show more</button></div>
</main>
<script type="application/json" id="data">{{JSON}}</script>
<script>
(function(){
var D=JSON.parse(document.getElementById('data').textContent),R=D.rules,M=D.meta,rows=D.rows,O=D.objs,I=D.inst;
var rank={Error:0,Warning:1,Review:2,Info:3},roleName={Template:'Template',OverriddenInstance:'Overridden instance',Instance:'Instance',Standalone:'Standalone'};
function el(t,c,x){var e=document.createElement(t);if(c)e.className=c;if(x!=null)e.textContent=x;return e;}
function obj(n){return O[n]||[0,'Standalone',0,'',''];}
document.getElementById('meta').textContent='Generated '+M.generated+' by GeoSCADA-Script-Analyser '+M.analyser+' from '+M.folder+'. Export: '+M.export+', exported '+M.exportedAt+' — '+M.files+' files, '+M.scanned+' scanned, '+M.instances+' instance copies counted against their template. Minimum severity: '+M.minSeverity+'. Also written: '+M.actionsCsv+', '+M.instancesCsv+', '+M.checkCsv+'.';
// Export check (the folder against the exporter's manifest): a red note when the export is incomplete.
(function(){var C=M.checkCounts||{},keys=Object.keys(C);if(!keys.length)return;
 var what={MissingFile:'listed in the manifest but missing from the folder',ExportError:'could not be exported (manifest Error rows)',Malformed:'have a broken header and were not analysed',
  DuplicateName:'share a FullName with another file',ManifestMismatch:'do not match their manifest row',UnsupportedFormat:'use a newer FormatVersion than this analyser',
  NotInManifest:M.manifestOnly?'are not in the manifest and were left out (-ManifestOnly)':'are not in the manifest (left over from an earlier export or another scope) but were analysed',
  NoManifest:'no manifest in the folder, so missing files cannot be detected',MixedExports:'files from more than one export run or exporter version (see the list)'};
 var n=el('div','note'+(M.incomplete?' bad':''));
 n.appendChild(el('b',null,M.incomplete?'Export incomplete: this report does not cover '+M.incomplete+' object(s). Re-export before relying on it.':'Export check'));
 var ul=el('ul');keys.forEach(function(k){ul.appendChild(el('li',null,k==='NoManifest'||k==='MixedExports'?what[k]:C[k]+' file(s) '+(what[k]||k)));});n.appendChild(ul);
 if(M.checks&&M.checks.length){var d=el('details');d.appendChild(el('summary',null,'Show the files'+(M.checks.length<keys.reduce(function(s,k){return s+C[k];},0)?' (first '+M.checks.length+'; all in '+M.checkCsv+')':' (also in '+M.checkCsv+')')));
  var box=el('div','chk'),t=el('table'),tb=el('tbody');M.checks.forEach(function(c){var tr=el('tr');tr.appendChild(el('td',null,c[0]));tr.appendChild(el('td','name',c[1]));tr.appendChild(el('td','sub',c[2]));tb.appendChild(tr);});
  t.appendChild(tb);box.appendChild(t);d.appendChild(box);n.appendChild(d);}
 document.getElementById('checknote').appendChild(n);})();
var cards=document.getElementById('cards'),bySev={Error:0,Warning:0,Review:0,Info:0},objs={};
rows.forEach(function(r){bySev[r[0]]++;objs[r[2]]=1;});
var names=Object.keys(objs),covered=0,ovr=0;names.forEach(function(n){covered+=obj(n)[2];if(obj(n)[1]==='OverriddenInstance')ovr++;});
[['Objects to edit',names.length,''],['Instance copies fixed by those edits',covered,'Template'],['Overridden instances (edit individually)',ovr,'OverriddenInstance']].concat(['Error','Warning','Review','Info'].map(function(s){return [s+' findings',bySev[s],s];})).forEach(function(c){var d=el('div','card');d.appendChild(el('b',c[2],String(c[1])));d.appendChild(el('span','sub',c[0]));cards.appendChild(d);});
var notes=document.getElementById('notes');
if(!M.progIdCheck)notes.appendChild(el('div','note','COM ProgID registry checks (RN06/RN07) were skipped.'));
else notes.appendChild(el('div','note','COM ProgID checks (RN06/RN07) used the registry of '+M.host+'. Results apply to machines built the same way as that one.'));
if(M.libCheckNotDone>0)notes.appendChild(el('div','note',M.libCheckNotDone+' mimic file(s) were exported without -CheckMissingLibraries, so links to missing libraries (GS01) are not known for them.'));
var rb=document.querySelector('#rules tbody'),langSel=document.getElementById('lang');
var agg={};rows.forEach(function(r){var a=agg[r[1]]||(agg[r[1]]={n:0,o:{}});a.n++;a.o[r[2]]=1;});
// Rule filter and by-rule table: every rule, in rule order (RN01.., then GS01..), so the complete picture is
// shown - 0 means checked with nothing found; a rule that did not run shows a dash and why (R[k].x).
// Unticked rules are remembered in this browser, so rules that will not be fixed stay hidden next time.
function ruleCmp(a,b){var pa=a.slice(0,2),pb=b.slice(0,2);if(pa!==pb)return pa==='RN'?-1:pb==='RN'?1:pa<pb?-1:1;return a.localeCompare(b,undefined,{numeric:true});}
var ruleIds=Object.keys(R).sort(ruleCmp);
ruleIds.forEach(function(k){if(!agg[k])agg[k]={n:0,o:{}};});
var ruleOn={},ruleBox={},ruleRow={},KEY='GeoSCADA-Script-Analyser.rulesOff',off=[];
try{off=JSON.parse(localStorage.getItem(KEY)||'[]')||[];}catch(e){off=[];}
ruleIds.forEach(function(k){ruleOn[k]=off.indexOf(k)<0;});
function saveRules(){try{var prev=JSON.parse(localStorage.getItem(KEY)||'[]')||[];
 var now=prev.filter(function(k){return !R[k];}).concat(ruleIds.filter(function(k){return !ruleOn[k];}));localStorage.setItem(KEY,JSON.stringify(now));}catch(e){}}
function syncRules(){var n=0;ruleIds.forEach(function(k){ruleBox[k].checked=ruleOn[k];ruleRow[k].classList.toggle('off',!ruleOn[k]);if(ruleOn[k])n++;});
 document.getElementById('rulesum').textContent=n===ruleIds.length?'All rules':n===0?'No rules':n===1?ruleIds.filter(function(k){return ruleOn[k];})[0]:n+' of '+ruleIds.length+' rules';}
function setRules(f){ruleIds.forEach(function(k){ruleOn[k]=f(k);});syncRules();saveRules();shown=200;render();}
// By-rule table: rule order by default (as the rule filter); click a column heading to sort by it, again to reverse.
var stat={};
ruleIds.forEach(function(k){
 var a=agg[k],x=R[k].x,tr=el('tr','rule'+(x?' notrun':a.n?'':' zero')),aff=0,on=Object.keys(a.o);on.forEach(function(o){aff+=1+obj(o)[2];});
 stat[k]=x?[rank[R[k].s],k,R[k].t,-1,-1,-1]:[rank[R[k].s],k,R[k].t,a.n,on.length,aff];     // not run sorts below 0
 var td=el('td');td.appendChild(el('span','pill '+R[k].s,R[k].s));tr.appendChild(td);tr.appendChild(el('td',null,k));
 var tt=el('td');tt.appendChild(el('div',null,R[k].t));if(x)tt.appendChild(el('div','notrunwhy',x));tt.appendChild(el('div','fix',R[k].f+' ['+R[k].r+']'));tr.appendChild(tt);
 [a.n,on.length,aff].forEach(function(v){tr.appendChild(el('td',null,x?'—':String(v)));});
 tr.title='Show only '+k;tr.addEventListener('click',function(){setRules(function(x){return x===k;});document.getElementById('q').scrollIntoView();});rb.appendChild(tr);ruleRow[k]=tr;});
var ths=document.querySelectorAll('#rules thead th'),sortCol=1,sortDir=1;
function sortRules(){
 var ids=ruleIds.slice().sort(function(a,b){var x=stat[a][sortCol],y=stat[b][sortCol],c;
  c=sortCol===1?ruleCmp(x,y):typeof x==='number'?x-y:x.localeCompare(y);
  return c*sortDir||ruleCmp(a,b);});
 ids.forEach(function(k){rb.appendChild(ruleRow[k]);});
 [].forEach.call(ths,function(th,i){th.setAttribute('aria-sort',i===sortCol?(sortDir>0?'ascending':'descending'):'none');});}
[].forEach.call(ths,function(th,i){th.className='sort';th.title='Sort by '+th.textContent;th.addEventListener('click',function(){
 if(sortCol===i)sortDir=-sortDir;else{sortCol=i;sortDir=i>=3?-1:1;}sortRules();});});
sortRules();
var rl=document.getElementById('rulelist');
ruleIds.forEach(function(k){var lb=el('label'),cb=el('input');cb.type='checkbox';
 cb.addEventListener('change',function(){ruleOn[k]=cb.checked;syncRules();saveRules();shown=200;render();});
 lb.appendChild(cb);lb.appendChild(el('b',R[k].s,k));lb.appendChild(el('span',null,R[k].t));lb.appendChild(el('span','sub',R[k].x?'(not run)':'('+agg[k].n+')'));if(R[k].x)lb.title=R[k].x;rl.appendChild(lb);ruleBox[k]=cb;});
document.querySelectorAll('.pickbtns button').forEach(function(b){b.addEventListener('click',function(){var p=b.getAttribute('data-p');
 setRules(function(k){return p==='all'?true:p==='none'?false:k.slice(0,2)===p;});});});
var pick=document.getElementById('rulepick');
document.addEventListener('click',function(e){if(pick.open&&!pick.contains(e.target))pick.open=false;});
document.addEventListener('keydown',function(e){if(e.key==='Escape')pick.open=false;});
syncRules();
var langs={};rows.forEach(function(r){langs[r[4]]=1;});Object.keys(langs).sort().forEach(function(l){var o=el('option',null,l);o.value=l;langSel.appendChild(o);});
var shown=200;
function instanceList(name){
 var list=I[name]||[];if(!list.length)return null;
 var d=el('details','inst');d.appendChild(el('summary','sub','Instance copies fixed by this change ('+list.length+')'));
 d.addEventListener('toggle',function(){if(!d.open||d.dataset.done)return;d.dataset.done=1;var ul=el('ul');
  list.slice(0,500).forEach(function(n){ul.appendChild(el('li',null,n));});d.appendChild(ul);
  if(list.length>500)d.appendChild(el('div','sub','… and '+(list.length-500)+' more in '+M.instancesCsv+'.'));});
 return d;}
function render(){
 var q=document.getElementById('q').value.toLowerCase(),sv=document.getElementById('sev').value,lg=langSel.value,ro=document.getElementById('role').value;
 var groups={},order=[];
 rows.forEach(function(r){if(sv&&r[0]!==sv)return;if(!ruleOn[r[1]])return;if(lg&&r[4]!==lg)return;if(ro&&obj(r[2])[1]!==ro)return;
  if(q&&(r[2]+' '+r[7]+' '+r[8]+' '+r[1]).toLowerCase().indexOf(q)<0&&!(I[r[2]]&&I[r[2]].some(function(n){return n.toLowerCase().indexOf(q)>=0;})))return;
  if(!groups[r[2]]){groups[r[2]]=[];order.push(r[2]);}groups[r[2]].push(r);});
 order.sort(function(a,b){return obj(a)[0]-obj(b)[0]||a.localeCompare(b);});
 var list=document.getElementById('list');list.textContent='';
 order.slice(0,shown).forEach(function(name){var g=groups[name],o=obj(name),d=el('details','obj'),s=el('summary');
  var worst=g.reduce(function(w,r){return rank[r[0]]<rank[w]?r[0]:w;},'Info');
  s.appendChild(el('span','prio','#'+o[0]));s.appendChild(el('span','pill '+worst,worst));s.appendChild(el('span','pill '+o[1],roleName[o[1]]||o[1]));
  s.appendChild(el('span','name',name));
  if(o[2]>0)s.appendChild(el('span','fixes','fixes '+o[2]+' instance cop'+(o[2]===1?'y':'ies')));
  if(o[4])s.appendChild(el('span','sub','linked by '+o[4]+' mimic(s)'));
  s.appendChild(el('span','sub',g[0][3]+' · '+g[0][4]+' · '+g.length+' finding(s) · '+g[0][10]));
  d.appendChild(s);
  d.addEventListener('toggle',function(){if(!d.open||d.dataset.done)return;d.dataset.done=1;
   d.appendChild(el('div','act',o[3]));var il=instanceList(name);if(il)d.appendChild(il);
   g.forEach(function(r){var f=el('div','f');
   var h=el('div');h.appendChild(el('span','pill '+r[0],r[0]));h.appendChild(document.createTextNode(' '+r[1]+' '+R[r[1]].t+(r[11]&&r[11]!=='Script'&&r[11]!=='Header'?' — '+r[11]:'')+(r[5]?' — line '+r[5]:'')+(r[6]?' (file line '+r[6]+')':'')));f.appendChild(h);
   if(r[7])f.appendChild(el('pre',null,r[7]));if(r[8])f.appendChild(el('div',null,r[8]));f.appendChild(el('div','fix',R[r[1]].f));d.appendChild(f);});});
  list.appendChild(d);});
 document.getElementById('count').textContent=order.length+' object(s) to edit';
 var mb=document.getElementById('morebtn');mb.hidden=order.length<=shown;mb.textContent='Show more ('+(order.length-shown)+' remaining)';
}
document.getElementById('morebtn').addEventListener('click',function(){shown+=200;render();});
['q','sev','role','lang'].forEach(function(id){document.getElementById(id).addEventListener('input',function(){shown=200;render();});});
render();
})();
</script></body></html>
'@
    [System.IO.File]::WriteAllText($htmlPath, $template.Replace('{{JSON}}', $json), (New-Object System.Text.UTF8Encoding($false)))
}

Write-Verbose ('{0,-24} {1,7:n1}s' -f 'HTML written', $sw.Elapsed.TotalSeconds)
# --- Summary -----------------------------------------------------------------
Write-Host ''
Write-Host ("Done in {0:n0}s. {1} finding(s) in {2} object(s) to edit; those edits also fix {3} instance copies." -f $sw.Elapsed.TotalSeconds, $sorted.Count, $actions.Count, (($actions | Measure-Object InstancesFixed -Sum).Sum)) -ForegroundColor Green
$ovrCount = @($actions | Where-Object Role -eq 'OverriddenInstance').Count
if ($ovrCount) { Write-Host "$ovrCount of them are overridden instances, which have to be edited individually (template changes do not reach them)." -ForegroundColor Magenta }
$byRule | Format-Table Severity, Rule, Title, Findings, ObjectsToEdit, ObjectsAffected -AutoSize | Out-String -Width 200 | Write-Host
Write-Host 'Top objects to edit:'
$actions | Select-Object -First 15 Priority, Severity, Role, InstancesFixed, Findings, FullName | Format-Table -AutoSize | Out-String -Width 220 | Write-Host
Write-Host "Findings CSV : $csvPath"
Write-Host "Actions CSV  : $actionsPath"
Write-Host "Instances CSV: $instancesPath"
Write-Host "Export check : $checkPath ($($checks.Count) problem(s))"
if (-not $NoHtml) { Write-Host "HTML         : $htmlPath" }
if ($libCheckNotDone) { Write-Warning "$libCheckNotDone mimic file(s) were exported without -CheckMissingLibraries; links to missing libraries (GS01) are unknown for them." }
if (-not $SkipProgIdCheck) { Write-Host "COM ProgID checks (RN06/RN07) used this machine's registry ($env:COMPUTERNAME)." -ForegroundColor DarkGray }
Write-Host 'Findings are pattern-based: review each one before changing anything.' -ForegroundColor Yellow
if ($incomplete) {
    Write-Warning "EXPORT INCOMPLETE: $incomplete object(s) are missing from or broken in the export (see $checkPath). Exit code 3."
    exit 3
}
