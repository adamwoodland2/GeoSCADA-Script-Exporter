param(
    [Parameter(Mandatory)][string]$Username,
    [Parameter(Mandatory)][string]$Password,
    [string]$Server = '127.0.0.1',
    [int]$Port = 5481,
    [switch]$Write
)
# Test-system tool, not part of the exporter / analyser. Builds Claude.Examples.Default (VBScript) and
# Claude.Examples.JScript from the Examples\RN-Examples-*.vbs / .js files, cloning Claude.Test's stored
# document (a VBScript mimic with one scriptCall pick action "RegistryTest") and replacing the script, the
# pick-action function name (same length: Run_Examples) and, for JScript, the language GUID.
# Without -Write it only builds and verifies the documents in memory. With -Write it updates the mimics
# (creating them in group Claude.Examples if needed), tries to create the RN01 test point, and reads back.
$ErrorActionPreference = 'Stop'
$repo = Join-Path (Split-Path -Parent $PSScriptRoot) 'Examples'
Add-Type -Path "C:\Program Files\Schneider Electric\ClearSCADA\ClearScada.Client.dll"

function Find-Bytes([byte[]]$Hay, [byte[]]$Needle, [int]$Start = 0) {
    for ($i = $Start; $i -le $Hay.Length - $Needle.Length; $i++) {
        $ok = $true; for ($j = 0; $j -lt $Needle.Length; $j++) { if ($Hay[$i + $j] -ne $Needle[$j]) { $ok = $false; break } }
        if ($ok) { return $i }
    }
    return -1
}
function Get-CStringBytes([string]$s) {          # MFC Unicode CString: FF FE FF + length + UTF-16
    $ms = New-Object IO.MemoryStream; $w = New-Object IO.BinaryWriter($ms)
    $w.Write([byte]0xFF); $w.Write([byte]0xFE); $w.Write([byte]0xFF)
    if ($s.Length -lt 0xFF) { $w.Write([byte]$s.Length) }
    elseif ($s.Length -lt 0xFFFF) { $w.Write([byte]0xFF); $w.Write([uint16]$s.Length) }
    else { $w.Write([byte]0xFF); $w.Write([uint16]0xFFFF); $w.Write([uint32]$s.Length) }
    $w.Write([Text.Encoding]::Unicode.GetBytes($s)); $w.Flush(); return , $ms.ToArray()
}
function Read-CStringLength([byte[]]$d, [int]$at) {   # at = index of FF FE FF; returns @(headerBytes, chars)
    $c = $at + 3; $n = [int]$d[$c]; $hdr = 4
    if ($n -eq 0xFF) { $n = [BitConverter]::ToUInt16($d, $c + 1); $hdr = 6; if ($n -eq 0xFFFF) { $n = [int][BitConverter]::ToUInt32($d, $c + 3); $hdr = 10 } }
    return @($hdr, $n)
}
function Get-Adler32([byte[]]$d) { [uint32]$a = 1; [uint32]$b = 0; foreach ($x in $d) { $a = ($a + $x) % 65521; $b = ($b + $a) % 65521 }; return [uint32](($b -shl 16) -bor $a) }
function Inflate([byte[]]$doc, [int]$z) {
    $ms = [IO.MemoryStream]::new($doc, [int]($z + 10), [int]($doc.Length - $z - 10))
    $ds = New-Object IO.Compression.DeflateStream($ms, [IO.Compression.CompressionMode]::Decompress)
    $o = New-Object IO.MemoryStream; $ds.CopyTo($o); return , $o.ToArray()
}

function New-MimicDocument([byte[]]$Template, [string]$Script, [string]$LangGuid, [string]$OldFn, [string]$NewFn) {
    $z = Find-Bytes $Template ([Text.Encoding]::ASCII.GetBytes('ZLIB'))
    $clen = [BitConverter]::ToUInt32($Template, $z - 4); $size = [BitConverter]::ToUInt32($Template, $z - 10)
    $inner = Inflate $Template $z
    if ($inner.Length -ne [BitConverter]::ToUInt32($Template, $z + 4)) { throw 'Template inflate length mismatch' }
    # 1. script: VBScript GUID + CString
    $vbGuid = ([Guid]'B54F3741-5B07-11CF-A4B0-00AA004A55E8').ToByteArray()
    $g = Find-Bytes $inner $vbGuid
    if ($g -lt 0 -or $inner[$g + 16] -ne 0xFF -or $inner[$g + 17] -ne 0xFE -or $inner[$g + 18] -ne 0xFF) { throw 'Script GUID + CString not found' }
    $len = Read-CStringLength $inner ($g + 16)
    $oldEnd = $g + 16 + $len[0] + 2 * $len[1]
    $newGuid = ([Guid]$LangGuid).ToByteArray()
    $out = New-Object IO.MemoryStream
    $out.Write($inner, 0, $g); $out.Write($newGuid, 0, 16)
    $cs = Get-CStringBytes $Script; $out.Write($cs, 0, $cs.Length)
    $out.Write($inner, $oldEnd, $inner.Length - $oldEnd)
    $inner2 = $out.ToArray()
    # 2. pick-action function name (same length, so no enclosing record changes)
    if ($OldFn.Length -ne $NewFn.Length) { throw 'Function names must be the same length' }
    #    stored as: 0D F0 AD 0B, uint32 byte length, UTF-16 text
    $nameBytes = [Text.Encoding]::Unicode.GetBytes($OldFn)
    $needle = [byte[]](@(0x0D, 0xF0, 0xAD, 0x0B) + [BitConverter]::GetBytes([uint32]$nameBytes.Length) + $nameBytes)
    $p = Find-Bytes $inner2 $needle
    if ($p -lt 0) { throw "Pick-action function '$OldFn' not found" }
    if ((Find-Bytes $inner2 $needle ($p + 1)) -ge 0) { throw "Pick-action function '$OldFn' found more than once" }
    $nb = [Text.Encoding]::Unicode.GetBytes($NewFn); [Array]::Copy($nb, 0, $inner2, $p + 8, $nb.Length)
    # 3. recompress: zlib header 78 9C + raw deflate + Adler-32 (big-endian)
    $dm = New-Object IO.MemoryStream
    $dfl = New-Object IO.Compression.DeflateStream($dm, [IO.Compression.CompressionMode]::Compress, $true)
    $dfl.Write($inner2, 0, $inner2.Length); $dfl.Close()
    $deflated = $dm.ToArray()
    $ad = Get-Adler32 $inner2
    $zl = New-Object IO.MemoryStream; $zl.WriteByte(0x78); $zl.WriteByte(0x9C); $zl.Write($deflated, 0, $deflated.Length)
    foreach ($s in 24, 16, 8, 0) { $zl.WriteByte([byte](($ad -shr $s) -band 0xFF)) }
    $zlib = $zl.ToArray()
    # 4. record: [size][ver][clen]["ZLIB"][ulen][zlib][rest of record] ; size covers ver..record end
    $tail = $size - (2 + 4 + $clen)                      # bytes in the record after the zlib block (8 in the template)
    $newClen = 8 + $zlib.Length
    $newSize = 2 + 4 + $newClen + $tail
    $doc = New-Object IO.MemoryStream; $bw = New-Object IO.BinaryWriter($doc)
    $bw.Write($Template, 0, $z - 10)
    $bw.Write([uint32]$newSize); $bw.Write([uint16][BitConverter]::ToUInt16($Template, $z - 6)); $bw.Write([uint32]$newClen)
    $bw.Write([Text.Encoding]::ASCII.GetBytes('ZLIB')); $bw.Write([uint32]$inner2.Length); $bw.Write($zlib)
    $restStart = $z + $clen
    $bw.Write($Template, $restStart, $Template.Length - $restStart); $bw.Flush()
    $result = $doc.ToArray()
    # verify: inflate the new document and compare
    $z2 = Find-Bytes $result ([Text.Encoding]::ASCII.GetBytes('ZLIB'))
    $check = Inflate $result $z2
    if ([Convert]::ToBase64String($check) -ne [Convert]::ToBase64String($inner2)) { throw 'Verification failed: rebuilt document does not inflate to the intended content' }
    return , $result
}

$conn = New-Object ClearScada.Client.Simple.Connection('Claude-Examples-Builder')
$conn.Connect((New-Object ClearScada.Client.ServerNode($Server, $Port)))
$conn.LogOn($Username, (ConvertTo-SecureString $Password -AsPlainText -Force))
try {
    $tmpl = $conn.GetObject('Claude.Test')
    $ms = New-Object IO.MemoryStream; $null = $conn.Server.GetDocumentContent($tmpl.Id, $null, $ms); $tbytes = $ms.ToArray()
    $vbs = ([IO.File]::ReadAllText("$repo\RN-Examples-VBScript.vbs") -replace "`r?`n", "`r`n").TrimEnd()
    $js  = ([IO.File]::ReadAllText("$repo\RN-Examples-JScript.js") -replace "`r?`n", "`r`n").TrimEnd()
    $targets = @(
        @{ Name = 'Claude.Examples.Default'; Script = $vbs; Guid = '{B54F3741-5B07-11CF-A4B0-00AA004A55E8}' }
        @{ Name = 'Claude.Examples.JScript'; Script = $js;  Guid = '{F414C260-6AC0-11CF-B6D1-00AA00BBBB58}' }
    )
    $docs = @{}
    foreach ($t in $targets) {
        $docs[$t.Name] = New-MimicDocument $tbytes $t.Script $t.Guid 'RegistryTest' 'Run_Examples'
        "built $($t.Name): $($docs[$t.Name].Length) bytes (verified)"
    }
    if (-not $Write) { 'Dry run only. Re-run with -Write to update the database.'; return }

    $group = $conn.GetObject('Claude.Examples')
    if (-not $group) { throw 'Group Claude.Examples not found' }
    foreach ($t in $targets) {
        $obj = $conn.GetObject($t.Name)
        if (-not $obj) { $obj = $conn.CreateObject('CMimic', $group.Id, $t.Name.Split('.')[-1]); "created $($t.Name)" }
        $warn = $null
        $ver = [int]$obj['ConfigVersion']
        $null = $conn.Server.SetDocumentContent($obj.Id, 'CMimic', (New-Object IO.MemoryStream(, $docs[$t.Name])), $ver, $false, [ref]$warn)
        "wrote $($t.Name) (was ConfigVersion $ver)" + $(if ($warn) { '; warnings: ' + (($warn | ForEach-Object { $_.ToString() }) -join ' | ') } else { '' })
    }
    $pt = $conn.GetObject('Claude.Examples.Test Analog')
    if (-not $pt) { try { $null = $conn.CreateObject('CPointAlgManual', $group.Id, 'Test Analog'); 'created Claude.Examples.Test Analog (Internal Analog Point)' } catch { "could not create Claude.Examples.Test Analog: $($_.Exception.InnerException.Message)" } }

    # read back through the web XML
    foreach ($t in $targets) {
        $o = $conn.GetObject($t.Name); $ms = New-Object IO.MemoryStream; $null = $conn.Server.GetWebDocumentContent($o.Id, $null, $ms); $b = $ms.ToArray()
        $x = New-Object Xml.XmlDocument; $x.PreserveWhitespace = $true; $x.LoadXml([Text.Encoding]::Unicode.GetString($b, 4, 2 * [BitConverter]::ToInt32($b, 0)).TrimEnd([char]0))
        $sc = $x.DocumentElement.SelectSingleNode('Script'); $act = $x.SelectSingleNode('//Action')
        $own = (($sc.ChildNodes | Where-Object { $_.NodeType -ne 'Element' } | ForEach-Object Value) -join '')
        "readback $($t.Name): lang=$($sc.GetAttribute('lang')) scriptMatches=$(($own -replace "`r?`n","`n") -eq ($t.Script -replace "`r?`n","`n")) action=$($act.GetAttribute('type')):$($act.InnerText.Trim()) ConfigVersion=$($o['ConfigVersion'])"
    }
} finally { $conn.Disconnect() }
