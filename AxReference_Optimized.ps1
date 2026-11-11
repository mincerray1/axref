# ============================================================================
#  AX 2012 R3 reference extractor
#
#  Put this whole file in your $PROFILE (or save it and dot-source it from
#  $PROFILE:   . C:\tools\AxReference.ps1 ), then in any folder run:
#
#       axref                       (asks everything interactively)
#       axref -TopFolder top -CusFolder cus -Keyword SalesLineGrouping -CallerDepth 4
#       axref -Keyword validateWrite -ObjectNameContains PurchAgree,SalesLine -CusObjectPrefix Fox
#       axref -ObjectNameContains PurchAgree -CusObjectPrefix Fox
#       axref -BuildIndex
#
#  Works on Windows PowerShell 5.1 and PowerShell 7.
#  Uses the Windows-provided winsqlite3.dll; no external SQLite package is needed.
# ============================================================================

$script:AxRefDir = Join-Path $HOME '.axref'
$script:AxRefRg = (Get-Command rg -ErrorAction SilentlyContinue | Select-Object -First 1).Source

$script:AxRefTypes = @(
    'Classes', 'Forms', 'Reports', 'Queries', 'Macros', 'Jobs',
    'Data Dictionary\Tables', 'Data Dictionary\Maps', 'Data Dictionary\Views',
    'Data Dictionary\Extended Data Types', 'Data Dictionary\Base Enums',
    'Data Dictionary\Table Collections', 'Data Dictionary\Configuration Keys',
    'Data Dictionary\License Codes',
    'Menus', 'Menu Items\Display', 'Menu Items\Output', 'Menu Items\Action',
    'Services', 'Service Groups',
    'Security\Roles', 'Security\Duties', 'Security\Privileges'
)

# Objects that nearly every method touches. They are never followed as a link
# (otherwise "called" and "caller" levels explode). Edit freely.
$script:AxRefNoise = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($n in @('Global', 'Args', 'Common', 'Info', 'Map', 'List', 'Set', 'Struct', 'Object', 'FormRun', 'Form',
                 'Query', 'QueryRun', 'QueryBuildDataSource', 'QueryBuildRange', 'QueryBuildFieldList', 'Dialog',
                 'DialogField', 'TreeNode', 'DictTable', 'DictField', 'DictEnum', 'DictClass', 'SysDictTable',
                 'SysDictClass', 'SysDictEnum', 'Box', 'Debug', 'Random', 'Array', 'Counter', 'Session', 'Connection',
                 'Statement', 'ResultSet', 'DateTimeUtil', 'Exception', 'NoYes', 'RecordInsertList',
                 'RecordSortedList', 'ClassFactory', 'Types')) { [void]$script:AxRefNoise.Add($n) }

Remove-Variable n -ErrorAction SilentlyContinue

# Words that are never treated as object or method references
$script:AxRefStop = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($n in @('if', 'else', 'while', 'for', 'switch', 'case', 'default', 'break', 'continue', 'return', 'try',
                 'catch', 'finally', 'throw', 'retry', 'int', 'str', 'real', 'boolean', 'date', 'utcdatetime',
                 'container', 'anytype', 'int64', 'void', 'public', 'private', 'protected', 'static', 'client',
                 'server', 'abstract', 'final', 'display', 'edit', 'select', 'from', 'where', 'join', 'order',
                 'group', 'by', 'firstonly', 'forupdate', 'next', 'print', 'super', 'this', 'true', 'false',
                 'null', 'ttsbegin', 'ttscommit', 'ttsabort', 'changecompany', 'delete_from', 'update_recordset',
                 'insert_recordset', 'sizeof', 'typeof')) { [void]$script:AxRefStop.Add($n) }

Remove-Variable n -ErrorAction SilentlyContinue

# Properties kept in the outline of non-class objects (tables, forms, EDTs, ...)
$script:AxRefKeep = [System.Collections.Generic.HashSet[string]]::new()
foreach ($n in @('extends', 'extendeddatatype', 'enumtype', 'table', 'query', 'datasource', 'datafield', 'datamethod',
                 'relatedtable', 'field', 'relatedfield', 'mandatory', 'stringsize', 'replacementkey', 'primaryindex',
                 'clusterindex', 'tablegroup', 'savedatapercompany', 'allowduplicates', 'type', 'arraylength',
                 'object', 'objecttype', 'enumvalue', 'class', 'method', 'parentkey', 'configurationkey',
                 'referencetable', 'joinsource', 'linktype', 'relations')) { [void]$script:AxRefKeep.Add($n) }

Remove-Variable n -ErrorAction SilentlyContinue

# ---- regexes -------------------------------------------------------------
$script:AxRefCodeRx = [regex]::new('@"[^"]*"|@''[^'']*''|"(?:\\.|[^"\\\r\n])*"|''(?:\\.|[^''\\\r\n])*''|//[^\r\n]*|/\*[\s\S]*?\*/', [System.Text.RegularExpressions.RegexOptions]::Compiled)
$script:AxRefSrcRx  = [regex]::new('(?ms)^[ \t]*SOURCE[ \t]+#(?<n>[^\r\n]*)\r?\n(?<b>.*?)^[ \t]*ENDSOURCE', [System.Text.RegularExpressions.RegexOptions]::Compiled)
$script:AxRefHeadRx = [regex]::new('(?m)^;\s*Microsoft Dynamics AX\s+(?<k>[^:\r\n]+?)\s*:\s*(?<n>\S+)\s+unloaded')
$script:AxRefNameRx = [regex]::new('(?m)^[ \t]*Name[ \t]+#(?<n>[^\r\n]*)')
$script:AxRefPropRx = [regex]::new('(?ms)^(?<prev>[^\r\n]*)\r?\n(?<ind>[ \t]*)PROPERTIES[ \t]*\r?\n(?<b>.*?)^[ \t]*ENDPROPERTIES')
$script:AxRefRefRxs = @(
    [regex]::new('\b(?:new|extends|implements)\s+([A-Za-z_]\w*)'),
    [regex]::new('(?<![\w.])([A-Za-z_]\w*)\s*::'),
    [regex]::new('\b\w*(?:Num|Str)\s*\(\s*([A-Za-z_]\w*)'),
    [regex]::new('(?<![\w.])([A-Za-z_]\w*)[ \t]+[a-z_]\w*[ \t]*[;,=)]')
)
$script:AxRefCallRx = [regex]::new('\b([A-Za-z_]\w*)\s*\(')

$script:AxRefEvalKeep = [System.Text.RegularExpressions.MatchEvaluator]{
    param($m)
    $v = $m.Value
    if ($v.StartsWith('//')) { return '' }
    if ($v.StartsWith('/*')) { if ($v.Contains("`n")) { return "`n" } else { return ' ' } }
    return $v
}
$script:AxRefEvalBlank = [System.Text.RegularExpressions.MatchEvaluator]{
    param($m)
    $v = $m.Value
    if ($v.StartsWith('//')) { return '' }
    if ($v.StartsWith('/*')) { return ' ' }
    return '""'
}

# ===========================================================================
#  small helpers
# ===========================================================================
function AxRef-Ask([string]$question, [string]$default) {
    $prompt = $question
    if ($default) { $prompt = "$question [$default]" }
    $a = Read-Host $prompt
    if ([string]::IsNullOrWhiteSpace($a)) { return $default }
    return $a.Trim()
}

function AxRef-MKey([string]$owner, [string]$name) {
    if ($owner) { return ($owner + '.' + $name).ToLowerInvariant() }
    return $name.ToLowerInvariant()
}

# removes comments (and, for analysis, string contents)
function AxRef-Strip([string]$code, [bool]$blankStrings) {
    if ($blankStrings) { return $script:AxRefCodeRx.Replace($code, $script:AxRefEvalBlank) }
    return $script:AxRefCodeRx.Replace($code, $script:AxRefEvalKeep)
}

# XPO source body -> plain code (drops the leading '#' that XPO puts on every line)
function AxRef-SourceToCode([string]$body) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($l in ($body -split "\r?\n")) {
        $i = $l.IndexOf('#')
        if ($i -ge 0 -and $l.Substring(0, $i).Trim().Length -eq 0) { $out.Add($l.Substring($i + 1)) }
    }
    return ($out -join "`r`n")
}

function AxRef-Tidy([string]$code) {
    $out = New-Object System.Collections.Generic.List[string]
    $blank = $true
    foreach ($l in ($code -split "\r?\n")) {
        $t = $l.TrimEnd()
        if ($t.Length -eq 0) {
            if (-not $blank) { $out.Add(''); $blank = $true }
        }
        else { $out.Add($t); $blank = $false }
    }
    while ($out.Count -gt 0 -and $out[$out.Count - 1].Length -eq 0) { $out.RemoveAt($out.Count - 1) }
    return ($out -join "`r`n")
}

function AxRef-TypeFromKind([string]$k) {
    $k = $k.ToLowerInvariant()
    if ($k -match 'table collection') { return 'Data Dictionary\Table Collections' }
    if ($k -match 'extended|user type') { return 'Data Dictionary\Extended Data Types' }
    if ($k -match 'enum')               { return 'Data Dictionary\Base Enums' }
    if ($k -match 'configuration')      { return 'Data Dictionary\Configuration Keys' }
    if ($k -match 'license')            { return 'Data Dictionary\License Codes' }
    if ($k -match 'class')              { return 'Classes' }
    if ($k -match 'table')              { return 'Data Dictionary\Tables' }
    if ($k -match 'view')               { return 'Data Dictionary\Views' }
    if ($k -match 'map')                { return 'Data Dictionary\Maps' }
    if ($k -match 'form')               { return 'Forms' }
    if ($k -match 'report')             { return 'Reports' }
    if ($k -match 'quer')               { return 'Queries' }
    if ($k -match 'macro')              { return 'Macros' }
    if ($k -match 'job')                { return 'Jobs' }
    if ($k -match 'display')            { return 'Menu Items\Display' }
    if ($k -match 'output')             { return 'Menu Items\Output' }
    if ($k -match 'action')             { return 'Menu Items\Action' }
    if ($k -match 'menu')               { return 'Menus' }
    if ($k -match 'service group')      { return 'Service Groups' }
    if ($k -match 'service')            { return 'Services' }
    if ($k -match 'role')               { return 'Security\Roles' }
    if ($k -match 'duty')               { return 'Security\Duties' }
    if ($k -match 'privilege')          { return 'Security\Privileges' }
    return 'Other'
}

function AxRef-ResolveType([string]$relDir, [string]$kind) {
    if ($relDir) {
        $d = $relDir.Replace('/', '\').Trim('\')
        foreach ($t in $script:AxRefTypes) { if ($t -ieq $d) { return $t } }
        return $d
    }
    return (AxRef-TypeFromKind $kind)
}

# ===========================================================================
#  XPO parsing
# ===========================================================================

# every SOURCE ... ENDSOURCE block (= one method / declaration)
function AxRef-GetSources([string]$text, [bool]$withOwner) {
    $list = New-Object System.Collections.Generic.List[object]
    $names = $null
    if ($withOwner) { $names = $script:AxRefNameRx.Matches($text) }
    $nameIndex = 0
    $owner = ''
    foreach ($m in $script:AxRefSrcRx.Matches($text)) {
        if ($withOwner) {
            while ($nameIndex -lt $names.Count -and $names[$nameIndex].Index -lt $m.Index) {
                $owner = $names[$nameIndex].Groups['n'].Value.Trim()
                $nameIndex++
            }
        }
        $list.Add([pscustomobject]@{
            Index = $m.Index
            Owner = $owner
            Name  = $m.Groups['n'].Value.Trim()
            Code  = (AxRef-SourceToCode $m.Groups['b'].Value)
        })
    }
    return , $list
}

# candidate object references and called method names of one method body
function AxRef-Extract([string]$analysisCode) {
    $refs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    foreach ($rx in $script:AxRefRefRxs) {
        foreach ($m in $rx.Matches($analysisCode)) {
            $t = $m.Groups[1].Value
            if (-not $script:AxRefStop.Contains($t)) { [void]$refs.Add($t) }
        }
    }
    $calls = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    foreach ($m in $script:AxRefCallRx.Matches($analysisCode)) {
        $t = $m.Groups[1].Value
        if (-not $script:AxRefStop.Contains($t)) { [void]$calls.Add($t.ToLowerInvariant()) }
    }
    $r = ','; $c = ','
    if ($refs.Count -gt 0)  { $r = ',' + [string]::Join(',', $refs) + ',' }
    if ($calls.Count -gt 0) { $c = ',' + [string]::Join(',', $calls) + ',' }
    return [pscustomobject]@{ Refs = $r; Calls = $c }
}

function AxRef-ParseText([string]$text, [string]$relDir, [string]$fileName) {
    $kind = ''
    $name = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    $hm = $script:AxRefHeadRx.Match($text)
    if ($hm.Success) { $kind = $hm.Groups['k'].Value; $name = $hm.Groups['n'].Value }
    $type = AxRef-ResolveType $relDir $kind
    $withOwner = ($type -eq 'Forms' -or $type -eq 'Reports')

    $methods = New-Object System.Collections.Generic.List[object]
    $declRefs = ','
    foreach ($s in (AxRef-GetSources $text $withOwner)) {
        $ana = AxRef-Strip $s.Code $true
        $x = AxRef-Extract $ana
        if ($s.Name -ieq 'classDeclaration') { $declRefs = $x.Refs; continue }
        $methods.Add([pscustomobject]@{ Owner = $s.Owner; Name = $s.Name; Calls = $x.Calls; Refs = $x.Refs })
    }
    return [pscustomobject]@{ Type = $type; Name = $name; DeclRefs = $declRefs; Methods = $methods }
}

# ===========================================================================
#  SQLite index cache (one database per top/cus pair)
# ===========================================================================
function AxRef-SqliteEnsureNative {
    if ('AxRefSqliteNative' -as [type]) { return }
    $dll = Join-Path $env:WINDIR 'System32\winsqlite3.dll'
    if (-not (Test-Path -LiteralPath $dll)) {
        throw "Windows SQLite library not found: $dll"
    }
    $source = @'
using System;
using System.Runtime.InteropServices;
public static class AxRefSqliteNative {
    [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl, CharSet=CharSet.Unicode, EntryPoint="sqlite3_open16")]
    public static extern int Open(string path, out IntPtr db);
    [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl, EntryPoint="sqlite3_close")]
    public static extern int Close(IntPtr db);
    [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl, EntryPoint="sqlite3_busy_timeout")]
    public static extern int BusyTimeout(IntPtr db, int milliseconds);
    [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl, CharSet=CharSet.Ansi, EntryPoint="sqlite3_exec")]
    public static extern int Exec(IntPtr db, string sql, IntPtr callback, IntPtr arg, out IntPtr errorMessage);
    [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl, EntryPoint="sqlite3_errmsg16")]
    public static extern IntPtr ErrorMessage(IntPtr db);
    [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl, EntryPoint="sqlite3_free")]
    public static extern void Free(IntPtr value);
    [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl, CharSet=CharSet.Unicode, EntryPoint="sqlite3_prepare16_v2")]
    public static extern int Prepare(IntPtr db, string sql, int bytes, out IntPtr statement, IntPtr tail);
    [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl, EntryPoint="sqlite3_finalize")]
    public static extern int Finalize(IntPtr statement);
    [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl, CharSet=CharSet.Unicode, EntryPoint="sqlite3_bind_text16")]
    public static extern int BindText(IntPtr statement, int index, string value, int bytes, IntPtr destructor);
    [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl, EntryPoint="sqlite3_bind_int64")]
    public static extern int BindInt64(IntPtr statement, int index, long value);
    [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl, EntryPoint="sqlite3_step")]
    public static extern int Step(IntPtr statement);
    [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl, EntryPoint="sqlite3_reset")]
    public static extern int Reset(IntPtr statement);
    [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl, EntryPoint="sqlite3_clear_bindings")]
    public static extern int ClearBindings(IntPtr statement);
    [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl, EntryPoint="sqlite3_column_int64")]
    public static extern long ColumnInt64(IntPtr statement, int column);
    [DllImport("winsqlite3.dll", CallingConvention=CallingConvention.Cdecl, EntryPoint="sqlite3_column_text16")]
    public static extern IntPtr ColumnText(IntPtr statement, int column);
}
'@
    Add-Type -TypeDefinition $source -ErrorAction Stop | Out-Null
}

function AxRef-SqliteOpen([string]$file) {
    AxRef-SqliteEnsureNative
    $db = [IntPtr]::Zero
    $rc = [AxRefSqliteNative]::Open($file, [ref]$db)
    if ($rc -ne 0) {
        $message = 'SQLite open failed with code ' + $rc
        if ($db -ne [IntPtr]::Zero) {
            $p = [AxRefSqliteNative]::ErrorMessage($db)
            if ($p -ne [IntPtr]::Zero) { $message = [Runtime.InteropServices.Marshal]::PtrToStringUni($p) }
            [void][AxRefSqliteNative]::Close($db)
        }
        throw $message
    }
    [void][AxRefSqliteNative]::BusyTimeout($db, 30000)
    return $db
}

function AxRef-SqliteExec($db, [string]$sql) {
    $errorMessage = [IntPtr]::Zero
    $rc = [AxRefSqliteNative]::Exec($db, $sql, [IntPtr]::Zero, [IntPtr]::Zero, [ref]$errorMessage)
    if ($rc -ne 0) {
        $message = 'SQLite command failed with code ' + $rc
        if ($errorMessage -ne [IntPtr]::Zero) {
            $message = [Runtime.InteropServices.Marshal]::PtrToStringAnsi($errorMessage)
            [AxRefSqliteNative]::Free($errorMessage)
        }
        throw $message
    }
}

function AxRef-SqlitePrepare($db, [string]$sql) {
    $statement = [IntPtr]::Zero
    $rc = [AxRefSqliteNative]::Prepare($db, $sql, -1, [ref]$statement, [IntPtr]::Zero)
    if ($rc -ne 0) {
        $p = [AxRefSqliteNative]::ErrorMessage($db)
        $message = 'SQLite prepare failed with code ' + $rc
        if ($p -ne [IntPtr]::Zero) { $message = [Runtime.InteropServices.Marshal]::PtrToStringUni($p) }
        throw $message
    }
    return $statement
}

function AxRef-SqliteBindText($statement, [int]$index, [string]$value) {
    if ($null -eq $value) { $value = '' }
    $rc = [AxRefSqliteNative]::BindText($statement, $index, $value, -1, [IntPtr](-1))
    if ($rc -ne 0) { throw "SQLite text binding failed with code $rc" }
}

function AxRef-SqliteBindInt64($statement, [int]$index, [int64]$value) {
    $rc = [AxRefSqliteNative]::BindInt64($statement, $index, $value)
    if ($rc -ne 0) { throw "SQLite integer binding failed with code $rc" }
}

function AxRef-SqliteRun($statement) {
    $rc = [AxRefSqliteNative]::Step($statement)
    if ($rc -ne 101) { throw "SQLite write failed with code $rc" }
    [void][AxRefSqliteNative]::Reset($statement)
    [void][AxRefSqliteNative]::ClearBindings($statement)
}

function AxRef-SqliteColumnText($statement, [int]$column) {
    $p = [AxRefSqliteNative]::ColumnText($statement, $column)
    if ($p -eq [IntPtr]::Zero) { return '' }
    return [Runtime.InteropServices.Marshal]::PtrToStringUni($p)
}

function AxRef-SqliteSchema($db) {
    AxRef-SqliteExec $db 'PRAGMA journal_mode=WAL; PRAGMA synchronous=NORMAL; PRAGMA foreign_keys=ON;
CREATE TABLE IF NOT EXISTS Meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS Files (
  layer TEXT NOT NULL, rel TEXT NOT NULL, ticks INTEGER NOT NULL, len INTEGER NOT NULL,
  type TEXT NOT NULL, name TEXT NOT NULL, decl_refs TEXT NOT NULL,
  PRIMARY KEY (layer, rel));
CREATE TABLE IF NOT EXISTS Methods (
  layer TEXT NOT NULL, rel TEXT NOT NULL, owner TEXT NOT NULL, name TEXT NOT NULL,
  calls TEXT NOT NULL, refs TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS Edges (
  source_layer TEXT NOT NULL, source_rel TEXT NOT NULL, source_owner TEXT NOT NULL,
  source_method TEXT NOT NULL, kind TEXT NOT NULL, target TEXT NOT NULL);
CREATE INDEX IF NOT EXISTS IX_Files_Name ON Files(name COLLATE NOCASE);
CREATE INDEX IF NOT EXISTS IX_Methods_File ON Methods(layer, rel);
CREATE INDEX IF NOT EXISTS IX_Edges_Target ON Edges(target COLLATE NOCASE, kind);
CREATE INDEX IF NOT EXISTS IX_Edges_Source ON Edges(source_layer, source_rel);'
}

function AxRef-SqliteCreateStatements($db) {
    return @{
        UpsertFile = (AxRef-SqlitePrepare $db 'INSERT INTO Files(layer,rel,ticks,len,type,name,decl_refs) VALUES(?1,?2,?3,?4,?5,?6,?7) ON CONFLICT(layer,rel) DO UPDATE SET ticks=excluded.ticks,len=excluded.len,type=excluded.type,name=excluded.name,decl_refs=excluded.decl_refs')
        DeleteMethods = (AxRef-SqlitePrepare $db 'DELETE FROM Methods WHERE layer=?1 AND rel=?2')
        DeleteEdges = (AxRef-SqlitePrepare $db 'DELETE FROM Edges WHERE source_layer=?1 AND source_rel=?2')
        InsertMethod = (AxRef-SqlitePrepare $db 'INSERT INTO Methods(layer,rel,owner,name,calls,refs) VALUES(?1,?2,?3,?4,?5,?6)')
        InsertEdge = (AxRef-SqlitePrepare $db 'INSERT INTO Edges(source_layer,source_rel,source_owner,source_method,kind,target) VALUES(?1,?2,?3,?4,?5,?6)')
        DeleteFile = (AxRef-SqlitePrepare $db 'DELETE FROM Files WHERE layer=?1 AND rel=?2')
    }
}

function AxRef-SqliteDisposeStatements($statements) {
    foreach ($statement in $statements.Values) {
        if ($statement -ne [IntPtr]::Zero) { [void][AxRefSqliteNative]::Finalize($statement) }
    }
}

function AxRef-SqliteStoreEntry($statements, $entry) {
    $layer = [string]$entry.Layer
    $rel = [string]$entry.Rel
    $deleteMethods = $statements.DeleteMethods
    AxRef-SqliteBindText $deleteMethods 1 $layer
    AxRef-SqliteBindText $deleteMethods 2 $rel
    AxRef-SqliteRun $deleteMethods
    $deleteEdges = $statements.DeleteEdges
    AxRef-SqliteBindText $deleteEdges 1 $layer
    AxRef-SqliteBindText $deleteEdges 2 $rel
    AxRef-SqliteRun $deleteEdges

    $file = $statements.UpsertFile
    AxRef-SqliteBindText $file 1 $layer
    AxRef-SqliteBindText $file 2 $rel
    AxRef-SqliteBindInt64 $file 3 $entry.Ticks
    AxRef-SqliteBindInt64 $file 4 $entry.Len
    AxRef-SqliteBindText $file 5 ([string]$entry.Type)
    AxRef-SqliteBindText $file 6 ([string]$entry.Name)
    AxRef-SqliteBindText $file 7 ([string]$entry.DeclRefs)
    AxRef-SqliteRun $file

    foreach ($method in $entry.Methods) {
        $insert = $statements.InsertMethod
        AxRef-SqliteBindText $insert 1 $layer
        AxRef-SqliteBindText $insert 2 $rel
        AxRef-SqliteBindText $insert 3 ([string]$method.Owner)
        AxRef-SqliteBindText $insert 4 ([string]$method.Name)
        AxRef-SqliteBindText $insert 5 ([string]$method.Calls)
        AxRef-SqliteBindText $insert 6 ([string]$method.Refs)
        AxRef-SqliteRun $insert
        foreach ($edgeSet in @(@('call', [string]$method.Calls), @('ref', [string]$method.Refs))) {
            foreach ($target in $edgeSet[1].Split(',')) {
                if (-not $target) { continue }
                $edge = $statements.InsertEdge
                AxRef-SqliteBindText $edge 1 $layer
                AxRef-SqliteBindText $edge 2 $rel
                AxRef-SqliteBindText $edge 3 ([string]$method.Owner)
                AxRef-SqliteBindText $edge 4 ([string]$method.Name)
                AxRef-SqliteBindText $edge 5 ([string]$edgeSet[0])
                AxRef-SqliteBindText $edge 6 $target
                AxRef-SqliteRun $edge
            }
        }
    }
    foreach ($target in ([string]$entry.DeclRefs).Split(',')) {
        if (-not $target) { continue }
        $edge = $statements.InsertEdge
        AxRef-SqliteBindText $edge 1 $layer
        AxRef-SqliteBindText $edge 2 $rel
        AxRef-SqliteBindText $edge 3 ''
        AxRef-SqliteBindText $edge 4 ''
        AxRef-SqliteBindText $edge 5 'decl'
        AxRef-SqliteBindText $edge 6 $target
        AxRef-SqliteRun $edge
    }
}

function AxRef-SqliteLoadEntries($db, [string]$topRoot, [string]$cusRoot, [bool]$includeMethods) {
    $entries = @{}
    $countStatement = AxRef-SqlitePrepare $db 'SELECT count(*) FROM Files'
    try {
        $countRc = [AxRefSqliteNative]::Step($countStatement)
        if ($countRc -ne 100) { throw "SQLite file count failed with code $countRc" }
        $fileTotal = [AxRefSqliteNative]::ColumnInt64($countStatement, 0)
    }
    finally { [void][AxRefSqliteNative]::Finalize($countStatement) }
    $statement = AxRef-SqlitePrepare $db 'SELECT layer,rel,ticks,len,type,name,decl_refs FROM Files'
    try {
        $loaded = 0
        while (($rc = [AxRefSqliteNative]::Step($statement)) -eq 100) {
            $loaded++
            $layer = AxRef-SqliteColumnText $statement 0
            $rel = AxRef-SqliteColumnText $statement 1
            $root = if ($layer -eq 'top') { $topRoot } else { $cusRoot }
            $entry = [pscustomobject]@{
                Layer = $layer; Rel = $rel; Path = (Join-Path $root $rel)
                Ticks = [AxRefSqliteNative]::ColumnInt64($statement, 2)
                Len = [AxRefSqliteNative]::ColumnInt64($statement, 3)
                Type = (AxRef-SqliteColumnText $statement 4)
                Name = (AxRef-SqliteColumnText $statement 5)
                DeclRefs = (AxRef-SqliteColumnText $statement 6)
                Methods = (New-Object System.Collections.Generic.List[object])
            }
            $entries[($layer + '|' + $rel)] = $entry
            if (($loaded % 1000) -eq 0 -or $loaded -eq $fileTotal) {
                $pct = [int](100 * $loaded / [Math]::Max(1, $fileTotal))
                Write-Progress -Activity 'Indexing XPO files' -Status ("Loading SQLite file index: {0} / {1}" -f $loaded, $fileTotal) -PercentComplete $pct
            }
        }
        if ($rc -ne 101) { throw "SQLite file read failed with code $rc" }
    }
    finally { [void][AxRefSqliteNative]::Finalize($statement) }

    if (-not $includeMethods) { return $entries }
    $countStatement = AxRef-SqlitePrepare $db 'SELECT count(*) FROM Methods'
    try {
        $countRc = [AxRefSqliteNative]::Step($countStatement)
        if ($countRc -ne 100) { throw "SQLite method count failed with code $countRc" }
        $methodTotal = [AxRefSqliteNative]::ColumnInt64($countStatement, 0)
    }
    finally { [void][AxRefSqliteNative]::Finalize($countStatement) }
    $statement = AxRef-SqlitePrepare $db 'SELECT layer,rel,owner,name,calls,refs FROM Methods'
    try {
        $loadedMethods = 0
        while (($rc = [AxRefSqliteNative]::Step($statement)) -eq 100) {
            $loadedMethods++
            $key = (AxRef-SqliteColumnText $statement 0) + '|' + (AxRef-SqliteColumnText $statement 1)
            if ($entries.ContainsKey($key)) {
                $entries[$key].Methods.Add([pscustomobject]@{
                    Owner = (AxRef-SqliteColumnText $statement 2)
                    Name = (AxRef-SqliteColumnText $statement 3)
                    Calls = (AxRef-SqliteColumnText $statement 4)
                    Refs = (AxRef-SqliteColumnText $statement 5)
                })
            }
            if (($loadedMethods % 5000) -eq 0 -or $loadedMethods -eq $methodTotal) {
                $pct = [int](100 * $loadedMethods / [Math]::Max(1, $methodTotal))
                Write-Progress -Activity 'Indexing XPO files' -Status ("Loading SQLite method index: {0} / {1}" -f $loadedMethods, $methodTotal) -PercentComplete $pct
            }
        }
        if ($rc -ne 101) { throw "SQLite method read failed with code $rc" }
    }
    finally { [void][AxRefSqliteNative]::Finalize($statement) }
    return $entries
}

function AxRef-SqliteMetaValue($db, [string]$key) {
    $statement = AxRef-SqlitePrepare $db 'SELECT value FROM Meta WHERE key=?1'
    try {
        AxRef-SqliteBindText $statement 1 $key
        $rc = [AxRefSqliteNative]::Step($statement)
        if ($rc -eq 100) { return (AxRef-SqliteColumnText $statement 0) }
        if ($rc -ne 101) { throw "SQLite metadata read failed with code $rc" }
        return ''
    }
    finally { [void][AxRefSqliteNative]::Finalize($statement) }
}

function AxRef-SqliteReadEdgeSources($statement, [string]$target, [string]$kind) {
    $sources = New-Object System.Collections.Generic.List[object]
    AxRef-SqliteBindText $statement 1 $target
    AxRef-SqliteBindText $statement 2 $kind
    while (($rc = [AxRefSqliteNative]::Step($statement)) -eq 100) {
        $sources.Add([pscustomobject]@{
            Layer = (AxRef-SqliteColumnText $statement 0)
            Type = (AxRef-SqliteColumnText $statement 1)
            Name = (AxRef-SqliteColumnText $statement 2)
            Owner = (AxRef-SqliteColumnText $statement 3)
            Method = (AxRef-SqliteColumnText $statement 4)
        })
    }
    if ($rc -ne 101) { throw "SQLite reverse-reference lookup failed with code $rc" }
    [void][AxRefSqliteNative]::Reset($statement)
    [void][AxRefSqliteNative]::ClearBindings($statement)
    return , $sources
}

function AxRef-ReadCache([string]$file) {
    $map = @{}
    try {
        $cur = $null
        foreach ($line in [System.IO.File]::ReadLines($file)) {
            if ($line.Length -lt 2) { continue }
            $f = $line.Split("`t")
            if ($f[0] -eq 'E') {
                $cur = [pscustomobject]@{
                    Layer = $f[1]; Rel = $f[2]; Path = ''; Ticks = [int64]$f[3]; Len = [int64]$f[4]
                    Type = $f[5]; Name = $f[6]; DeclRefs = $f[7]
                    Methods = (New-Object System.Collections.Generic.List[object])
                }
                $map[($f[1] + '|' + $f[2])] = $cur
            }
            elseif ($f[0] -eq 'M' -and $cur) {
                $cur.Methods.Add([pscustomobject]@{ Owner = $f[1]; Name = $f[2]; Calls = $f[3]; Refs = $f[4] })
            }
        }
    }
    catch {
        Write-Warning ("Could not read the XPO index cache '{0}'; it will be rebuilt. {1}" -f $file, $_.Exception.Message)
        return @{}
    }
    return $map
}

function AxRef-GetKeywordHitPaths([string[]]$roots, [string[]]$keywords) {
    if ($keywords.Count -eq 0) {
        return , ([System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase))
    }
    if (-not $script:AxRefRg) { return $null }

    $rgArgs = @('--files-with-matches', '--null', '--ignore-case', '--fixed-strings',
        '--hidden', '--no-ignore', '--text', '--glob-case-insensitive', '--glob', '*.xpo')
    foreach ($keyword in $keywords) {
        $rgArgs += @('--regexp', $keyword)
    }
    $rgArgs += '--'
    $rgArgs += $roots

    $output = @(& $script:AxRefRg @rgArgs)
    $exitCode = $LASTEXITCODE
    if ($exitCode -gt 1) {
        Write-Warning ("ripgrep failed with exit code {0}; using the PowerShell keyword-search fallback." -f $exitCode)
        return $null
    }

    $paths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if ($exitCode -eq 0) {
        $separator = [char[]]@([char]0)
        foreach ($path in ([string]::Join('', $output).Split($separator, [System.StringSplitOptions]::RemoveEmptyEntries))) {
            [void]$paths.Add([System.IO.Path]::GetFullPath($path))
        }
    }
    return , $paths
}

# Builds or incrementally updates the SQLite XPO and reverse-reference index.
function AxRef-BuildIndex([string]$topRoot, [string]$cusRoot, [string[]]$keywords, [bool]$rebuild, [bool]$indexOnly = $false) {
    New-Item -ItemType Directory -Force -Path $script:AxRefDir | Out-Null
    $md5 = [System.Security.Cryptography.MD5]::Create()
    $hash = [System.BitConverter]::ToString($md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes("$topRoot|$cusRoot"))).Replace('-', '')
    $dbFile = Join-Path $script:AxRefDir "index_$hash.sqlite3"
    $legacyFile = Join-Path $script:AxRefDir "index_$hash.tsv"

    $act = 'Indexing XPO files'
    Write-Progress -Activity $act -Status 'Loading cache...' -PercentComplete 0
    $db = AxRef-SqliteOpen $dbFile
    $statements = $null
    $transactionOpen = $false
    $cancelled = $false
    $old = @{}
    $new = @{}
    $hits = New-Object System.Collections.Generic.List[object]
    $hitTexts = @{}
    $parsed = 0
    $i = 0
    $total = 0
    try {
        AxRef-SqliteSchema $db
        $statements = AxRef-SqliteCreateStatements $db

        if ((AxRef-SqliteMetaValue $db 'legacy_tsv_imported') -ne '1') {
            if (-not $rebuild -and (Test-Path -LiteralPath $legacyFile)) {
                $legacy = AxRef-ReadCache $legacyFile
                $legacyCount = $legacy.Count
                $migrated = 0
                if ($legacyCount -gt 0) {
                    AxRef-SqliteExec $db 'BEGIN'
                    $transactionOpen = $true
                    foreach ($entry in $legacy.Values) {
                        AxRef-SqliteStoreEntry $statements $entry
                        $migrated++
                        if (($migrated % 500) -eq 0) {
                            AxRef-SqliteExec $db 'COMMIT'
                            $transactionOpen = $false
                            $pct = [int](100 * $migrated / $legacyCount)
                            Write-Progress -Activity $act -Status ("Importing existing TSV cache: {0} / {1}" -f $migrated, $legacyCount) -PercentComplete $pct
                            AxRef-SqliteExec $db 'BEGIN'
                            $transactionOpen = $true
                        }
                    }
                    AxRef-SqliteExec $db 'COMMIT'
                    $transactionOpen = $false
                }
            }
            AxRef-SqliteExec $db "INSERT OR REPLACE INTO Meta(key,value) VALUES('legacy_tsv_imported','1')"
        }
        $legacy = $null
        $old = AxRef-SqliteLoadEntries $db $topRoot $cusRoot (-not $indexOnly)

        # List files first so the progress bar can show percentage and ETA.
        Write-Progress -Activity $act -Status 'Counting XPO files...' -PercentComplete 0
        $work = New-Object System.Collections.Generic.List[object]
        $counted = 0
        foreach ($pair in @(@('top', $topRoot), @('cus', $cusRoot))) {
            $di = New-Object System.IO.DirectoryInfo($pair[1])
            foreach ($fi in $di.EnumerateFiles('*.xpo', [System.IO.SearchOption]::AllDirectories)) {
                $work.Add([pscustomobject]@{ Layer = $pair[0]; Root = $pair[1]; File = $fi })
                $counted++
                if (($counted % 1000) -eq 0) {
                    Write-Progress -Activity $act -Status ("Counting XPO files: {0} found..." -f $counted) -PercentComplete 0
                }
            }
        }
        $total = $work.Count
        Write-Host ("Found {0} XPO files." -f $total)

        $hitPaths = AxRef-GetKeywordHitPaths @($topRoot, $cusRoot) $keywords
        if ($null -ne $hitPaths -and $keywords.Count -gt 0) {
            Write-Host ("ripgrep shortlisted {0} keyword-matching XPO files." -f $hitPaths.Count)
        }

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $lastUi = -1000
        $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        AxRef-SqliteExec $db 'BEGIN'
        $transactionOpen = $true
        try {
            foreach ($w in $work) {
                $i++
                $layer = $w.Layer; $root = $w.Root; $fi = $w.File

                if (($sw.ElapsedMilliseconds - $lastUi) -ge 250 -or $i -eq $total) {
                    $lastUi = $sw.ElapsedMilliseconds
                    $pct = [int](100 * $i / [Math]::Max(1, $total))
                    $eta = [System.TimeSpan]::FromMilliseconds($sw.ElapsedMilliseconds / [Math]::Max(1, $i) * ($total - $i))
                    $status = ('{0} / {1} ({2}%)   elapsed {3}   remaining ~{4}   re-parsed {5}   keyword hits {6}' -f `
                        $i, $total, $pct, $sw.Elapsed.ToString('hh\:mm\:ss'), $eta.ToString('hh\:mm\:ss'), $parsed, $hits.Count)
                    Write-Progress -Activity $act -Status $status -PercentComplete $pct -CurrentOperation ($layer + ': ' + $fi.Name)
                }

                $rel = $fi.FullName.Substring($root.Length).TrimStart('\', '/')
                $ck = $layer + '|' + $rel
                [void]$seen.Add($ck)
                $e = $old[$ck]
                $unchanged = ($e -and -not $rebuild -and $e.Ticks -eq $fi.LastWriteTimeUtc.Ticks -and $e.Len -eq $fi.Length)
                $isHit = $false
                $text = $null
                if ($null -eq $hitPaths) {
                    $text = [System.IO.File]::ReadAllText($fi.FullName)
                    foreach ($k in $keywords) {
                        if ($text.IndexOf($k, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                            $isHit = $true
                            break
                        }
                    }
                }
                else {
                    $isHit = $hitPaths.Contains($fi.FullName)
                    if (-not $unchanged -or $isHit) { $text = [System.IO.File]::ReadAllText($fi.FullName) }
                }

                if (-not $unchanged) {
                    $relDir = [System.IO.Path]::GetDirectoryName($rel)
                    if ($null -eq $relDir) { $relDir = '' }
                    $p = AxRef-ParseText $text $relDir $fi.Name
                    $e = [pscustomobject]@{
                        Layer = $layer; Rel = $rel; Path = ''; Ticks = $fi.LastWriteTimeUtc.Ticks; Len = $fi.Length
                        Type = $p.Type; Name = $p.Name; DeclRefs = $p.DeclRefs; Methods = $p.Methods
                    }
                    AxRef-SqliteStoreEntry $statements $e
                    $parsed++
                }
                $e.Path = $fi.FullName
                $new[$ck] = $e
                if ($isHit) {
                    $hits.Add($e)
                    if ($null -ne $text) { $hitTexts[$ck] = $text }
                }

                if (($i % 250) -eq 0 -and $i -lt $total) {
                    AxRef-SqliteExec $db 'COMMIT'
                    $transactionOpen = $false
                    AxRef-SqliteExec $db 'BEGIN'
                    $transactionOpen = $true
                }
            }
        }
        catch {
            if ($_.Exception -is [System.Management.Automation.PipelineStoppedException] -or $_.FullyQualifiedErrorId -match 'PipelineStopped') {
                $cancelled = $true
            }
            else { throw }
        }

        if ($transactionOpen) {
            AxRef-SqliteExec $db 'COMMIT'
            $transactionOpen = $false
        }

        if (-not $cancelled) {
            $deleteMethods = $statements.DeleteMethods
            $deleteEdges = $statements.DeleteEdges
            $deleteFile = $statements.DeleteFile
            $removed = 0
            AxRef-SqliteExec $db 'BEGIN'
            $transactionOpen = $true
            foreach ($key in $old.Keys) {
                if ($seen.Contains($key)) { continue }
                $entry = $old[$key]
                foreach ($stmt in @($deleteMethods, $deleteEdges, $deleteFile)) {
                    AxRef-SqliteBindText $stmt 1 ([string]$entry.Layer)
                    AxRef-SqliteBindText $stmt 2 ([string]$entry.Rel)
                    AxRef-SqliteRun $stmt
                }
                $removed++
                if (($removed % 500) -eq 0) {
                    AxRef-SqliteExec $db 'COMMIT'
                    $transactionOpen = $false
                    AxRef-SqliteExec $db 'BEGIN'
                    $transactionOpen = $true
                }
            }
            if ($transactionOpen) {
                AxRef-SqliteExec $db 'COMMIT'
                $transactionOpen = $false
            }
        }
    }
    catch {
        $isPipelineStop = ($_.Exception -is [System.Management.Automation.PipelineStoppedException] -or $_.FullyQualifiedErrorId -match 'PipelineStopped')
        if ($transactionOpen) {
            if ($isPipelineStop) { AxRef-SqliteExec $db 'COMMIT' }
            else { AxRef-SqliteExec $db 'ROLLBACK' }
            $transactionOpen = $false
        }
        if ($isPipelineStop) { $cancelled = $true }
        else { throw }
    }
    finally {
        if ($transactionOpen) {
            AxRef-SqliteExec $db 'COMMIT'
        }
        if ($statements) { AxRef-SqliteDisposeStatements $statements }
        [void][AxRefSqliteNative]::Close($db)
        Write-Progress -Activity $act -Completed
    }

    return [pscustomobject]@{
        Entries = $new; Hits = $hits; HitTexts = $hitTexts; Parsed = $parsed
        Canceled = $cancelled; IndexPath = $dbFile; IndexedFiles = $new.Count
    }
}

# merges TOP and CUS entries into one model (CUS methods win)
function AxRef-BuildModel($entries, [string]$cusObjectPrefix = '', [string]$indexPath = '') {
    $model = [pscustomobject]@{
        Elems   = (New-Object System.Collections.Generic.List[object])
        ByKey   = @{}
        ByName  = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
        Methods = (New-Object System.Collections.Generic.List[object])
        IndexPath = $indexPath
        CusObjectPrefix = $cusObjectPrefix
    }
    $cusKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if ($cusObjectPrefix) {
        foreach ($e in $entries.Values) {
            if ($e.Layer -eq 'cus') { [void]$cusKeys.Add(($e.Type + '|' + $e.Name)) }
        }
    }
    $topPrefixWithoutCus = New-Object System.Collections.Generic.List[string]
    $act = 'Building object model'
    $done = 0; $all = $entries.Count
    foreach ($layer in @('top', 'cus')) {
        foreach ($e in $entries.Values) {
            if ($e.Layer -ne $layer) { continue }
            $key = ($e.Type + '|' + $e.Name).ToLowerInvariant()
            if ($layer -eq 'top' -and $cusObjectPrefix -and $e.Name.StartsWith($cusObjectPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                if ($cusKeys.Contains(($e.Type + '|' + $e.Name))) { continue }
                $topPrefixWithoutCus.Add($e.Type + '\' + $e.Name)
            }
            $done++
            if (($done % 500) -eq 0) {
                Write-Progress -Activity $act -Status ('{0} / {1}' -f $done, $all) -PercentComplete ([int](100 * $done / [Math]::Max(1, $all)))
            }
            $el = $model.ByKey[$key]
            if (-not $el) {
                $el = [pscustomobject]@{
                    Key = $key; Type = $e.Type; Name = $e.Name; TopPath = $null; CusPath = $null
                    DeclRefs = ','; DeclRefsLayer = $null
                    Methods = @{}; MethodsByName = @{}; Noise = $script:AxRefNoise.Contains($e.Name)
                }
                $model.ByKey[$key] = $el
                $model.Elems.Add($el)
                if ($model.ByName.ContainsKey($e.Name)) { $model.ByName[$e.Name].Add($el) }
                else {
                    $lst = New-Object System.Collections.Generic.List[object]
                    $lst.Add($el)
                    $model.ByName[$e.Name] = $lst
                }
            }
            if ($layer -eq 'top') { $el.TopPath = $e.Path } else { $el.CusPath = $e.Path }
            if ($e.DeclRefs.Length -gt 1) { $el.DeclRefs = $e.DeclRefs; $el.DeclRefsLayer = $layer }
            foreach ($m in $e.Methods) {
                $mk = AxRef-MKey $m.Owner $m.Name
                $mo = $el.Methods[$mk]
                if ($mo) { $mo.Calls = $m.Calls; $mo.Refs = $m.Refs; $mo.SourceLayer = $layer }
                else {
                    $mo = [pscustomobject]@{ Elem = $el; Owner = $m.Owner; Name = $m.Name; Key = $mk; Calls = $m.Calls; Refs = $m.Refs; SourceLayer = $layer }
                    $el.Methods[$mk] = $mo
                    $methodName = $m.Name.ToLowerInvariant()
                    if (-not $el.MethodsByName.ContainsKey($methodName)) {
                        $el.MethodsByName[$methodName] = New-Object System.Collections.Generic.List[object]
                    }
                    $el.MethodsByName[$methodName].Add($mo)
                    $model.Methods.Add($mo)
                }
            }
        }
    }
    Write-Progress -Activity $act -Completed
    if ($topPrefixWithoutCus.Count -gt 0) {
        Write-Warning ("CUS object prefix '{0}' matched TOP object(s) without a CUS counterpart; kept their existing TOP behavior: {1}" -f $cusObjectPrefix, (($topPrefixWithoutCus | Sort-Object -Unique) -join ', '))
    }
    return $model
}

# ===========================================================================
#  keyword hits -> methods / elements
# ===========================================================================
function AxRef-FindHits($model, $entries, [string[]]$keywords, $hitTexts, [string]$cusObjectPrefix = '') {
    $hitMethods = New-Object System.Collections.Generic.List[object]
    $hitElems   = New-Object System.Collections.Generic.List[object]
    $seen = New-Object System.Collections.Generic.HashSet[string]

    $act = 'Locating keyword matches'
    $n = 0; $cnt = $entries.Count
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $lastUi = -1000
    foreach ($e in $entries) {
        $n++
        if (($sw.ElapsedMilliseconds - $lastUi) -ge 300 -or $n -eq $cnt) {
            $lastUi = $sw.ElapsedMilliseconds
            $pct = [Math]::Min(100, [Math]::Max(0, [int](100 * $n / [Math]::Max(1, $cnt))))
            Write-Progress -Activity $act -Status ('{0} / {1}   {2}' -f $n, $cnt, $e.Name) -PercentComplete $pct
        }
        $el = $model.ByKey[($e.Type + '|' + $e.Name).ToLowerInvariant()]
        if (-not $el) { continue }
        if ($e.Layer -eq 'top' -and $el.CusPath -and $cusObjectPrefix -and $e.Name.StartsWith($cusObjectPrefix, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        $ck = $e.Layer + '|' + $e.Rel
        if ($hitTexts.ContainsKey($ck)) { $text = $hitTexts[$ck] }
        else { $text = [System.IO.File]::ReadAllText($e.Path) }
        $withOwner = ($e.Type -eq 'Forms' -or $e.Type -eq 'Reports')

        foreach ($s in (AxRef-GetSources $text $withOwner)) {
            $found = $false
            foreach ($k in $keywords) {
                if ($s.Code.IndexOf($k, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $found = $true; break }
            }
            if (-not $found) { continue }
            $mo = $el.Methods[(AxRef-MKey $s.Owner $s.Name)]
            if ($mo) { if ($seen.Add($el.Key + '|' + $mo.Key)) { $hitMethods.Add($mo) } }
            else     { if ($seen.Add($el.Key + '|')) { $hitElems.Add($el) } }
        }

        # keyword outside any method (properties, fields, declaration of a table...)
        $rest = $script:AxRefSrcRx.Replace($text, '')
        foreach ($k in $keywords) {
            if ($rest.IndexOf($k, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                if ($seen.Add($el.Key + '|')) { $hitElems.Add($el) }
                break
            }
        }
    }
    Write-Progress -Activity $act -Completed
    return [pscustomobject]@{ Methods = $hitMethods; Elems = $hitElems }
}

function AxRef-Include($sel, $roles, $el, [string]$role) {
    if (-not $sel.ContainsKey($el.Key)) {
        $sel[$el.Key] = $el
        $roles[$el.Key] = New-Object System.Collections.Generic.List[string]
    }
    if (-not $roles[$el.Key].Contains($role)) { $roles[$el.Key].Add($role) }
}

# ===========================================================================
#  caller / called expansion
# ===========================================================================

# frontier items: [pscustomobject]@{ Elem; Method }  (Method $null = whole element)
function AxRef-FindCallers($model, $frontier) {
    $res  = New-Object System.Collections.Generic.List[object]
    $seen = New-Object System.Collections.Generic.HashSet[string]
    $byName = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $elemOnly = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($f in $frontier) {
        if ($null -ne $f.Method) {
            $ln = $f.Method.Name
            if (-not $byName.ContainsKey($ln)) { $byName[$ln] = New-Object System.Collections.Generic.List[object] }
            $byName[$ln].Add($f)
        }
        elseif (-not $f.Elem.Noise) { [void]$elemOnly.Add($f.Elem.Name) }
    }

    $db = AxRef-SqliteOpen $model.IndexPath
    $statement = [IntPtr]::Zero
    try {
        $statement = AxRef-SqlitePrepare $db 'SELECT e.source_layer,f.type,f.name,e.source_owner,e.source_method
FROM Edges e JOIN Files f ON f.layer=e.source_layer AND f.rel=e.source_rel
WHERE e.target=?1 COLLATE NOCASE AND e.kind=?2'

        foreach ($methodName in $byName.Keys) {
            $candidates = AxRef-SqliteReadEdgeSources $statement $methodName 'call'
            foreach ($source in $candidates) {
                $elemKey = ($source.Type + '|' + $source.Name).ToLowerInvariant()
                $elem = $model.ByKey[$elemKey]
                if (-not $elem) { continue }
                $methodKey = AxRef-MKey $source.Owner $source.Method
                $x = $elem.Methods[$methodKey]
                if (-not $x -or $x.SourceLayer -ne $source.Layer) { continue }
                foreach ($f in $byName[$methodName]) {
                    if ([object]::ReferenceEquals($x, $f.Method)) { continue }
                    $ok = [object]::ReferenceEquals($x.Elem, $f.Elem)
                    if (-not $ok -and -not $f.Elem.Noise) {
                        $probe = ',' + $f.Elem.Name + ','
                        $ok = ($x.Refs.Contains($probe) -or $x.Elem.DeclRefs.Contains($probe))
                    }
                    if ($ok -and $seen.Add($x.Elem.Key + '|' + $x.Key)) {
                        $res.Add([pscustomobject]@{ Elem = $x.Elem; Method = $x })
                    }
                }
            }
        }

        foreach ($elementName in $elemOnly) {
            $methodCandidates = AxRef-SqliteReadEdgeSources $statement $elementName 'ref'
            foreach ($source in $methodCandidates) {
                $elemKey = ($source.Type + '|' + $source.Name).ToLowerInvariant()
                $elem = $model.ByKey[$elemKey]
                if (-not $elem) { continue }
                $methodKey = AxRef-MKey $source.Owner $source.Method
                $x = $elem.Methods[$methodKey]
                if ($x -and $x.SourceLayer -eq $source.Layer -and $seen.Add($x.Elem.Key + '|' + $x.Key)) {
                    $res.Add([pscustomobject]@{ Elem = $x.Elem; Method = $x })
                }
            }

            $elementCandidates = AxRef-SqliteReadEdgeSources $statement $elementName 'decl'
            foreach ($source in $elementCandidates) {
                $elemKey = ($source.Type + '|' + $source.Name).ToLowerInvariant()
                $elem = $model.ByKey[$elemKey]
                if (-not $elem -or $elem.DeclRefsLayer -ne $source.Layer) { continue }
                if ($seen.Add($elem.Key + '|')) {
                    $res.Add([pscustomobject]@{ Elem = $elem; Method = $null })
                }
            }
        }
    }
    finally {
        if ($statement -ne [IntPtr]::Zero) { [void][AxRefSqliteNative]::Finalize($statement) }
        [void][AxRefSqliteNative]::Close($db)
    }
    return , $res
}

function AxRef-ExpandCalled($model, $seeds, [int]$depth, $sel, $roles, [int]$maxObjects) {
    $seen = New-Object System.Collections.Generic.HashSet[string]
    $frontier = New-Object System.Collections.Generic.List[object]
    foreach ($s in $seeds) {
        if ($null -ne $s.Method -and $seen.Add($s.Elem.Key + '|' + $s.Method.Key)) { $frontier.Add($s) }
    }

    for ($lvl = 1; $lvl -le $depth; $lvl++) {
        if ($frontier.Count -eq 0) { break }
        Write-Host ("Called level {0}: following {1} methods ..." -f $lvl, $frontier.Count)
        $next = New-Object System.Collections.Generic.List[object]

        foreach ($f in $frontier) {
            $m = $f.Method

            # objects referenced by this method
            foreach ($tok in $m.Refs.Split(',')) {
                if ($tok.Length -eq 0 -or -not $model.ByName.ContainsKey($tok)) { continue }
                foreach ($r in $model.ByName[$tok]) {
                    if ($r.Noise -or [object]::ReferenceEquals($r, $f.Elem)) { continue }
                    if ($sel.Count -ge $maxObjects -and -not $sel.ContainsKey($r.Key)) {
                        Write-Warning "Object limit ($maxObjects) reached while following 'called' links."
                        return
                    }
                    AxRef-Include $sel $roles $r ('called L' + $lvl)

                    # specific methods of that object that this method calls
                    $matched = $false
                    foreach ($callName in $m.Calls.Split(',')) {
                        if (-not $callName -or -not $r.MethodsByName.ContainsKey($callName)) { continue }
                        foreach ($rm in $r.MethodsByName[$callName]) {
                            $matched = $true
                            if ($seen.Add($r.Key + '|' + $rm.Key)) { $next.Add([pscustomobject]@{ Elem = $r; Method = $rm }) }
                        }
                    }
                    # a class that is only instantiated: follow its constructor
                    if (-not $matched -and $r.Type -eq 'Classes') {
                        $ctor = $r.Methods['new']
                        if ($ctor -and $seen.Add($r.Key + '|' + $ctor.Key)) { $next.Add([pscustomobject]@{ Elem = $r; Method = $ctor }) }
                    }
                }
            }

            # other methods of the same object that this method calls
            foreach ($callName in $m.Calls.Split(',')) {
                if (-not $callName -or -not $f.Elem.MethodsByName.ContainsKey($callName)) { continue }
                foreach ($sm in $f.Elem.MethodsByName[$callName]) {
                    if ([object]::ReferenceEquals($sm, $m)) { continue }
                    if ($seen.Add($f.Elem.Key + '|' + $sm.Key)) {
                        $next.Add([pscustomobject]@{ Elem = $f.Elem; Method = $sm })
                    }
                }
            }
        }
        $frontier = $next
    }
}

# ===========================================================================
#  sanitizing
# ===========================================================================
function AxRef-Sanitize([string]$text, [string]$type, [string]$name, [string]$label) {
    $withOwner = ($type -eq 'Forms' -or $type -eq 'Reports')
    $items = New-Object System.Collections.Generic.List[object]

    foreach ($s in (AxRef-GetSources $text $withOwner)) {
        $code = AxRef-Tidy (AxRef-Strip $s.Code $false)
        $items.Add([pscustomobject]@{ Index = $s.Index; Kind = 'code'; Owner = $s.Owner; Text = $code })
    }

    if ($type -ne 'Classes') {
        foreach ($pm in $script:AxRefPropRx.Matches($text)) {
            $prev = $pm.Groups['prev'].Value.Trim()
            $kind = ($prev -split '\s+')[0]
            $depth = [Math]::Max(0, [int](($pm.Groups['ind'].Value.Length - 4) / 2))
            $pname = ''
            $kv = New-Object System.Collections.Generic.List[string]
            foreach ($pl in ($pm.Groups['b'].Value -split "\r?\n")) {
                if ($pl -match '^\s*(\w+)\s+#(.*)$') {
                    $k = $Matches[1]; $v = $Matches[2].Trim()
                    if ($k -ieq 'Name') { $pname = $v }
                    elseif ($v.Length -gt 0 -and $script:AxRefKeep.Contains($k.ToLower())) { $kv.Add($k + '=' + $v) }
                }
            }
            if ($pname) {
                $line = ('  ' * $depth) + $kind.ToLower() + ' ' + $pname
                if ($kv.Count -gt 0) { $line += '  [' + ($kv -join ', ') + ']' }
                $items.Add([pscustomobject]@{ Index = $pm.Index; Kind = 'prop'; Owner = ''; Text = $line })
            }
        }
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("// $type : $name  [$label]")
    $prevKind = ''
    $lastOwner = ''
    foreach ($it in ($items | Sort-Object Index)) {
        if ($it.Kind -eq 'prop') {
            if ($prevKind -eq 'code') { [void]$sb.AppendLine('') }
            [void]$sb.AppendLine($it.Text)
        }
        else {
            if ($it.Text.Length -eq 0) { continue }
            [void]$sb.AppendLine('')
            if ($withOwner -and $it.Owner -and $it.Owner -ne $name -and $it.Owner -ne $lastOwner) {
                [void]$sb.AppendLine('// [' + $it.Owner + ']')
                $lastOwner = $it.Owner
            }
            [void]$sb.AppendLine($it.Text)
        }
        $prevKind = $it.Kind
    }
    return $sb.ToString()
}

# ===========================================================================
#  project XPO (format is learned from a sample project exported from AX)
# ===========================================================================
function AxRef-WalkNodes($node, $model, $acc) {
    foreach ($c in $node.Children) {
        $name = ''
        foreach ($h in $c.Header) {
            if ($h -match '^\s*NAME\s+#(.+?)\s*$') { $name = $Matches[1]; break }
        }
        $hdr = New-Object System.Collections.Generic.List[string]
        foreach ($h in $c.Header) {
            $x = $h -replace '^(\s*UTILOBJECTID\s+)\d+', '${1}0'
            if ($name) { $x = $x -replace ('(?i)#' + [regex]::Escape($name) + '\s*$'), '#{{NAME}}' }
            $hdr.Add($x)
        }
        if ($c.Children.Count -gt 0) {
            if (-not $acc.Group) { $acc.Group = $hdr.ToArray() }
            AxRef-WalkNodes $c $model $acc
        }
        elseif ($name -and $model.ByName.ContainsKey($name)) {
            $t = $model.ByName[$name][0].Type
            if (-not $acc.Leaf.ContainsKey($t)) { $acc.Leaf[$t] = $hdr.ToArray() }
        }
        else { $acc.Unmatched.Add($name) }
    }
}

function AxRef-LearnTemplate([string]$samplePath, $model) {
    $bytes = [System.IO.File]::ReadAllBytes($samplePath)
    $enc = 'ansi'
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { $enc = 'utf16le' }
    elseif ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { $enc = 'utf8bom' }

    $lines = [System.IO.File]::ReadAllText($samplePath) -split "\r?\n"
    $first = -1; $last = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $t = $lines[$i].Trim()
        if ($t -eq 'BEGINNODE' -and $first -lt 0) { $first = $i }
        if ($t -eq 'ENDNODE') { $last = $i }
    }
    if ($first -lt 0 -or $last -lt 0) { Write-Warning 'No BEGINNODE/ENDNODE found in the sample: is it a project xpo?'; return $null }

    $prefix = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $first; $i++) { $prefix.Add($lines[$i]) }
    $suffix = New-Object System.Collections.Generic.List[string]
    for ($i = $last + 1; $i -lt $lines.Count; $i++) { $suffix.Add($lines[$i]) }

    $pname = ''
    foreach ($l in $prefix) { if ($l -match '^\s*PROJECT\s+#(\S+)') { $pname = $Matches[1]; break } }
    $guidRx = '\{[0-9A-Fa-f]{8}(-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}\}'
    $norm = {
        param($arr)
        $o = New-Object System.Collections.Generic.List[string]
        foreach ($l in $arr) {
            $x = $l -replace $guidRx, '{{GUID}}'
            if ($pname) { $x = $x -replace ('\b' + [regex]::Escape($pname) + '\b'), '{{PROJECT}}' }
            $o.Add($x)
        }
        return , $o.ToArray()
    }
    $prefixN = & $norm $prefix
    $suffixN = & $norm $suffix

    # node tree
    $root = [pscustomobject]@{ Header = (New-Object System.Collections.Generic.List[string]); Children = (New-Object System.Collections.Generic.List[object]) }
    $stack = New-Object System.Collections.Stack
    $stack.Push($root)
    for ($i = $first; $i -le $last; $i++) {
        $t = $lines[$i].Trim()
        if ($t -eq 'BEGINNODE') {
            $n = [pscustomobject]@{ Header = (New-Object System.Collections.Generic.List[string]); Children = (New-Object System.Collections.Generic.List[object]) }
            $stack.Peek().Children.Add($n)
            $stack.Push($n)
        }
        elseif ($t -eq 'ENDNODE') { [void]$stack.Pop() }
        else {
            $cur = $stack.Peek()
            if ($cur.Children.Count -eq 0) { $cur.Header.Add($lines[$i]) }
        }
    }

    $acc = @{ Group = $null; Leaf = @{}; Unmatched = (New-Object System.Collections.Generic.List[string]) }
    AxRef-WalkNodes $root $model $acc

    if ($acc.Leaf.Count -eq 0) {
        Write-Warning 'No object nodes from your top/cus folders were found in the sample.'
        return $null
    }
    if ($acc.Unmatched.Count -gt 0) {
        Write-Warning ('Sample objects not found in your top/cus folders (skipped): ' + ($acc.Unmatched -join ', '))
    }

    $indent = ($lines[$first] -replace '\S.*$', '')
    $tpl = [pscustomobject]@{
        Encoding = $enc; Indent = $indent; Prefix = $prefixN; Suffix = $suffixN
        Leaf = $acc.Leaf
    }
    New-Item -ItemType Directory -Force -Path $script:AxRefDir | Out-Null
    $tpl | ConvertTo-Json -Depth 6 | Set-Content -Path (Join-Path $script:AxRefDir 'project-template.json') -Encoding UTF8
    Write-Host ('Project template saved. Object types learned: ' + (($acc.Leaf.Keys | Sort-Object) -join ', '))
    return (Get-Content -Raw -Path (Join-Path $script:AxRefDir 'project-template.json') | ConvertFrom-Json)
}

function AxRef-LoadTemplate {
    $p = Join-Path $script:AxRefDir 'project-template.json'
    if (Test-Path $p) { return (Get-Content -Raw -Path $p | ConvertFrom-Json) }
    return $null
}

function AxRef-WriteProject($tpl, [string]$projectName, $elements, [string]$path) {
    $guid = '{' + [guid]::NewGuid().ToString().ToUpper() + '}'
    $ind = ''
    if ($tpl.Indent) { $ind = [string]$tpl.Indent }
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($l in @($tpl.Prefix)) { $lines.Add($l.Replace('{{PROJECT}}', $projectName).Replace('{{GUID}}', $guid)) }

    # one flat list, same types next to each other (classes, then forms, then tables, ...)
    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($el in ($elements | Sort-Object Type, Name)) {
        $lt = $tpl.Leaf.PSObject.Properties[$el.Type]
        if (-not $lt) { $missing.Add($el.Type); continue }
        $lines.Add($ind + 'BEGINNODE')
        foreach ($h in @($lt.Value)) { $lines.Add($h.Replace('{{NAME}}', $el.Name)) }
        $lines.Add($ind + 'ENDNODE')
    }

    foreach ($l in @($tpl.Suffix)) { $lines.Add($l.Replace('{{PROJECT}}', $projectName).Replace('{{GUID}}', $guid)) }

    $enc = switch ($tpl.Encoding) {
        'utf16le' { New-Object System.Text.UnicodeEncoding($false, $true) }
        'utf8bom' { New-Object System.Text.UTF8Encoding($true) }
        default   { [System.Text.Encoding]::Default }
    }
    [System.IO.File]::WriteAllText($path, (($lines -join "`r`n") + "`r`n"), $enc)
    return , $missing
}

# ===========================================================================
#  AGENTS.md / MANIFEST.md
# ===========================================================================
function AxRef-WriteDocs([string]$out, $info, $sel, $roles, $fileMap) {
    $kw = ($info.Keywords | ForEach-Object { '`' + $_ + '`' }) -join ', '
    $nameFilters = ($info.ObjectNameContains | ForEach-Object { '`' + $_ + '`' }) -join ', '
    if (-not $nameFilters) { $nameFilters = '(none)' }
    $cusObjectPrefix = '(none)'
    if ($info.CusObjectPrefix) { $cusObjectPrefix = '`' + $info.CusObjectPrefix + '`' }
    $manifestRows = @{}
    $manifestPath = Join-Path $out 'MANIFEST.md'
    if (Test-Path $manifestPath) {
        foreach ($line in [System.IO.File]::ReadLines($manifestPath)) {
            if (-not $line.StartsWith('|')) { continue }
            $fields = $line.Trim().Trim('|').Split('|')
            if ($fields.Count -lt 4) { continue }
            $type = $fields[0].Trim()
            $name = $fields[1].Trim()
            if (-not $type -or $type -eq 'Type' -or $type -match '^-+$') { continue }
            $manifestRows[($type + '|' + $name).ToLowerInvariant()] = [pscustomobject]@{
                Type = $type; Name = $name; Files = $fields[2].Trim(); Roles = $fields[3].Trim()
            }
        }
    }

    # ---- manifest
    $mf = New-Object System.Text.StringBuilder
    [void]$mf.AppendLine('# Manifest')
    [void]$mf.AppendLine('')
    [void]$mf.AppendLine("Keywords: $kw  ")
    [void]$mf.AppendLine("Object-name fragments: $nameFilters  ")
    [void]$mf.AppendLine("CUS-only object prefix: $cusObjectPrefix  ")
    [void]$mf.AppendLine("Caller depth: $($info.CallerDepth), called depth: $($info.CalledDepth)  ")
    [void]$mf.AppendLine("Generated: $($info.When)")
    [void]$mf.AppendLine('')
    [void]$mf.AppendLine('Roles: `match` = keyword found in the object, `target` = object name contains a requested name fragment, `caller Ln` = n levels up (L1 = the method containing the keyword), `called Ln` = n levels down.')
    [void]$mf.AppendLine('')
    [void]$mf.AppendLine('| Type | Object | Files | Roles |')
    [void]$mf.AppendLine('|---|---|---|---|')
    foreach ($el in ($sel.Values | Sort-Object Type, Name)) {
        $files = ($fileMap[$el.Key] -join '<br>')
        $r = ($roles[$el.Key] -join '; ')
        $manifestRows[$el.Key] = [pscustomobject]@{
            Type = $el.Type; Name = $el.Name; Files = $files; Roles = $r
        }
    }
    foreach ($row in ($manifestRows.Values | Sort-Object Type, Name)) {
        [void]$mf.AppendLine("| $($row.Type) | $($row.Name) | $($row.Files) | $($row.Roles) |")
    }
    [System.IO.File]::WriteAllText((Join-Path $out 'MANIFEST.md'), $mf.ToString(), (New-Object System.Text.UTF8Encoding($false)))

    # ---- agents
    $projText = 'No project xpo was generated for this reference set.'
    if ($info.ProjectFile) {
        $projText = 'The file `' + $info.ProjectFile + '` is an importable AX project that lists every object in this folder, with objects of the same type next to each other. It only references objects by name and type; it contains no code. In the AX project window you can also switch to "Group by type".'
    }

    $doc = @'
# AGENTS.md

This folder is a **sanitized reference set** extracted from a Microsoft Dynamics AX 2012 R3 code base. It exists so an AI assistant (or a developer) can understand a specific area of the system and help fix bugs or build features, without exposing comments, names or other private information.

## What was searched

- Keywords: {{KEYWORDS}}
- Object-name fragments: {{OBJECTNAMEFRAGMENTS}}
- CUS-only object prefix: {{CUSOBJECTPREFIX}}
- Caller depth: {{CALLERDEPTH}}  /  called depth: {{CALLEDDEPTH}}
- Objects included: {{COUNT}} (see `MANIFEST.md` for every object and why it is here)

## Folder layout

The sub folders mirror the AOT:

`Classes`, `Forms`, `Reports`, `Queries`, `Macros`, `Jobs`, `Data Dictionary\Tables`, `Data Dictionary\Maps`, `Data Dictionary\Views`, `Data Dictionary\Extended Data Types`, `Data Dictionary\Base Enums`, `Data Dictionary\Table Collections`, `Data Dictionary\Configuration Keys`, `Data Dictionary\License Codes`, `Menus`, `Menu Items\Display|Output|Action`, `Services`, `Service Groups`, `Security\Roles|Duties|Privileges`.

Each object is one file `<ObjectName>.xpp`.

## `_top` and `_cus` suffixes

Two sources were used:

- **TOP**: the complete object as the AOT shows it. For standard objects with no customization this is the pure standard (SYS) code. For standard objects that were customized, it is the standard code **with the customizations merged in**.
- **CUS**: only the customizations: new objects, plus for modified objects only the methods and properties that were changed or added.

How to read the file names:

- `Name_top.xpp` + `Name_cus.xpp`: the object exists in both. `_top` is the full object, `_cus` shows just what was customized. If a method is in both, the `_cus` version is the customization.
- `Name.xpp` (no suffix) and the object is only in TOP: standard object, not customized.
- `Name.xpp` (no suffix) and the object is only in CUS: a purely custom object that has no standard version.

## What "sanitized" means

- All comments (`//` and `/* */`) were removed, because they can contain names or other private information.
- XPO metadata (ids, GUIDs, versions, `#` prefixes, layer information, UI properties) was removed. The files are not importable XPO.
- Classes contain declarations and methods only. Other objects contain a short outline (fields, indexes, relations, data sources, controls with the key properties) followed by their methods.
- String literals were kept as they are.

## How the objects were chosen

Static text analysis, so treat it as a good guide, not a guarantee:

- A method is a **keyword match** when the keyword appears in its source (comments included, before they were removed).
- **Callers**: methods that call a matched method (same object, or an object that references the method's object), repeated up to the caller depth.
- **Called**: objects, and the specific methods of those objects, that a matched method uses, repeated up to the called depth. Tables, EDTs and enums stop the chain unless one of their methods is called.
- Calls made through reflection, runtime class names, form control events and framework hooks may be missing. Very common framework objects (Global, Args, Common, Query...) are intentionally not followed.

## Project xpo

{{PROJECT}}

## How development in Dynamics AX 2012 R3 works (short primer)

- **AOT**: the Application Object Tree holds everything: classes, tables, forms, queries, reports, EDTs, base enums, menus, security, services. Developers work in the MorphX IDE and code in **X++** (C#-like, case insensitive, declarations at the top of a method, no `var`, containers are 1-based, `strFmt`, `info/warning/error` write to the infolog).
- **Layers**: code is stored in layers: SYS, SYP, GLS, GLP, FPK, FPP, SLN, SLP, ISV, ISP, VAR, VAP, CUS, CUP, USR, USP. A higher layer overrides the lower one. Customizations are stored as differences, per method and property. Changing a standard object in CUS is called **over-layering**: it works, but it makes upgrades harder, so keep changes small and put new logic in new classes where possible.
- **No extension methods or partial classes** in AX 2012. R2/R3 do support pre/post **event handlers** and **delegates**, which can avoid over-layering.
- **Data access**: `select`, `while select`, `firstOnly`, `forUpdate`, `join`, `exists join`, `update_recordset`, `delete_from`, `insert_recordset`. Table buffers derive from `Common`. Database changes happen inside `ttsBegin` / `ttsCommit` (`ttsAbort` rolls back).
- **Tables**: standard methods `find`, `exist`, `initValue`, `modifiedField`, `validateField`, `validateWrite`, `insert`, `update`, `delete`. Table relations, indexes, field groups and delete actions are metadata, not code.
- **Classes**: `classDeclaration` holds member variables; `new` is the constructor; the usual pattern is a `static construct()` plus `parm` methods. Batch/dialog frameworks: `RunBase`, `RunBaseBatch`, `SysOperation`. Pass data between forms and classes with `Args`.
- **Forms**: form methods, data source methods (`init`, `executeQuery`, `active`, field `modified`), and control methods (`clicked`, `modified`). Their methods are listed in the file with the owner in a `// [Owner]` line.
- **Types**: EDTs (extended data types) carry labels, lengths and relations. Base enums (such as `NoYes`) are accessed as `NoYes::Yes`. Labels look like `@SYS12345`.
- **Macros**: lines starting with `#` in code are macros (`#define`, `#localmacro`, included macro libraries). Those are real X++ and were kept.

## Suggestions for whoever uses this folder

- Prefer the smallest possible change and say which layer and object/method it goes in.
- When proposing code, give the complete method, in X++ 2012 style (no features newer than AX 2012 R3).
- If an object is missing from this folder, say so instead of guessing what it contains.

## When asked for help with a feature or bug

- Create or update a root-level `recommendation.md` with the proposed code changes. Show the relevant before and after code, a syntax-highlighted `diff` block, and complete copy/pasteable X++ code.
- Explain the change and recommendation, list possible risks or assumptions, and provide logic checks or test scenarios to validate the behavior.
- Ask whether more context is needed. If an object's behavior or a keyword is unclear, name the specific object or keyword that needs a definition instead of guessing.
- If asked for a technical design document, create a separate root-level `TDD.md`. Keep both `recommendation.md` and `TDD.md` available to Git.
'@

    $doc = $doc.Replace('{{KEYWORDS}}', $kw).Replace('{{OBJECTNAMEFRAGMENTS}}', $nameFilters)
    $doc = $doc.Replace('{{CUSOBJECTPREFIX}}', $cusObjectPrefix)
    $doc = $doc.Replace('{{CALLERDEPTH}}', [string]$info.CallerDepth).Replace('{{CALLEDDEPTH}}', [string]$info.CalledDepth)
    $doc = $doc.Replace('{{COUNT}}', [string]$manifestRows.Count).Replace('{{PROJECT}}', $projText)
    [System.IO.File]::WriteAllText((Join-Path $out 'AGENTS.md'), $doc, (New-Object System.Text.UTF8Encoding($false)))

    $gitignore = @'
# Generated X++ reference files. Git omits their folders when they contain no other files.
*.xpp
# Keep recommendation.md, TDD.md, and other Markdown documents trackable.
'@
    [System.IO.File]::WriteAllText((Join-Path $out '.gitignore'), $gitignore, (New-Object System.Text.UTF8Encoding($false)))
}

# ===========================================================================
#  main
# ===========================================================================
function Export-AxReference {
    [CmdletBinding()]
    param(
        [string]$TopFolder,
        [string]$CusFolder,
        [string[]]$Keyword,
        [string[]]$ObjectNameContains,
        [string]$CusObjectPrefix,
        [switch]$BuildIndex,
        [int]$CallerDepth = -1,
        [int]$CalledDepth = -1,
        [int]$MaxObjects = 400,
        [switch]$Project,
        [string]$ProjectName,
        [string]$ProjectSample,
        [switch]$RebuildIndex
    )

    $cwd = (Get-Location).Path
    $lastFile = Join-Path $script:AxRefDir 'last.json'
    $last = $null
    if (Test-Path $lastFile) { try { $last = Get-Content -Raw $lastFile | ConvertFrom-Json } catch { $last = $null } }
    $defTop = 'top'; $defCus = 'cus'
    if ($last) { $defTop = $last.Top; $defCus = $last.Cus }

    if (-not $TopFolder) { $TopFolder = AxRef-Ask 'Top folder name (top layer, with or without customizations)' $defTop }
    if (-not $CusFolder) { $CusFolder = AxRef-Ask 'CUS folder name (customizations only)' $defCus }

    $topRoot = $TopFolder; $cusRoot = $CusFolder
    if (-not [System.IO.Path]::IsPathRooted($topRoot)) { $topRoot = Join-Path $cwd $topRoot }
    if (-not [System.IO.Path]::IsPathRooted($cusRoot)) { $cusRoot = Join-Path $cwd $cusRoot }
    $topRoot = $topRoot.TrimEnd('\', '/'); $cusRoot = $cusRoot.TrimEnd('\', '/')
    foreach ($p in @($topRoot, $cusRoot)) {
        if (-not (Test-Path $p -PathType Container)) { Write-Error "Folder not found: $p"; return }
    }
    New-Item -ItemType Directory -Force -Path $script:AxRefDir | Out-Null
    ([pscustomobject]@{ Top = $TopFolder; Cus = $CusFolder } | ConvertTo-Json) | Set-Content -Path $lastFile -Encoding UTF8

    if ($BuildIndex) {
        Write-Host 'Building/updating the SQLite reference index. Press Ctrl+C to stop; completed batches are retained.'
        $noKeywords = [string[]]@()
        $ix = AxRef-BuildIndex $topRoot $cusRoot $noKeywords ([bool]$RebuildIndex) $true
        if ($ix.Canceled) { Write-Warning ("Index build stopped. Completed work is saved in '{0}' and will resume on the next run." -f $ix.IndexPath) }
        else {
            Write-Host ("Index ready: {0} files; parsed this run: {1}." -f $ix.IndexedFiles, $ix.Parsed)
            Write-Host ("Stored at {0}" -f $ix.IndexPath)
        }
        return
    }

    # keywords
    $kws = New-Object System.Collections.Generic.List[string]
    if ($Keyword) { foreach ($k in $Keyword) { if ($k.Trim()) { $kws.Add($k.Trim()) } } }
    $nameFilters = New-Object System.Collections.Generic.List[string]
    if ($ObjectNameContains) {
        foreach ($name in $ObjectNameContains) {
            if (-not [string]::IsNullOrWhiteSpace($name)) { $nameFilters.Add($name.Trim()) }
        }
    }
    if ($kws.Count -eq 0 -and $nameFilters.Count -eq 0) {
        Write-Host 'Enter keywords one at a time. Press Enter to skip/finish keyword entry; you can add object-name fragments next.'
        while ($true) {
            $k = Read-Host ('Keyword #' + ($kws.Count + 1))
            if ([string]::IsNullOrWhiteSpace($k)) { break }
            $kws.Add($k.Trim())
        }
    }

    # optional object-name fragments
    if ($nameFilters.Count -eq 0) {
        Write-Host 'Optional: enter object-name fragments to include matching objects (case-insensitive substring). Press Enter on an empty line to finish.'
        while ($true) {
            $name = Read-Host ('Object name fragment #' + ($nameFilters.Count + 1))
            if ([string]::IsNullOrWhiteSpace($name)) { break }
            $nameFilters.Add($name.Trim())
        }
    }
    if (-not $PSBoundParameters.ContainsKey('CusObjectPrefix')) {
        $CusObjectPrefix = Read-Host 'Optional CUS-only object prefix (blank keeps normal TOP/CUS handling)'
        if ($CusObjectPrefix) { $CusObjectPrefix = $CusObjectPrefix.Trim() }
    }
    if ($kws.Count -eq 0 -and $nameFilters.Count -eq 0) {
        Write-Warning 'Enter at least one keyword or object-name fragment.'
        return
    }

    if ($CallerDepth -lt 0) { $CallerDepth = [int](AxRef-Ask 'Caller depth (levels up)' '5') }
    if ($CalledDepth -lt 0) { $CalledDepth = [int](AxRef-Ask 'Called depth (levels down)' '5') }

    # output folder
    $out = Join-Path $cwd 'reference'
    if (Test-Path $out) {
        Write-Host "Keeping existing reference folder '$out'. Existing objects will only be replaced after confirmation."
    }

    # ---- index + hits
    Write-Host 'Reading and updating the SQLite XPO index (unchanged files reuse the cache)...'
    $ix = AxRef-BuildIndex $topRoot $cusRoot $kws.ToArray() ([bool]$RebuildIndex)
    if ($ix.Canceled) {
        Write-Warning ("Index update stopped. Completed work is saved in '{0}'. Rerun axref to resume." -f $ix.IndexPath)
        return
    }
    Write-Host ('SQLite index: ' + $ix.IndexPath)
    Write-Host ('Files: ' + $ix.Entries.Count + ', (re)parsed this run: ' + $ix.Parsed + ', files containing a keyword: ' + $ix.Hits.Count)
    $model = AxRef-BuildModel $ix.Entries $CusObjectPrefix $ix.IndexPath
    Write-Host ('Objects: ' + $model.Elems.Count + ', methods: ' + $model.Methods.Count)

    $hits = AxRef-FindHits $model $ix.Hits $kws.ToArray() $ix.HitTexts $CusObjectPrefix
    $targets = New-Object System.Collections.Generic.List[object]
    foreach ($el in $model.Elems) {
        foreach ($name in $nameFilters) {
            if ($el.Name.IndexOf($name, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $targets.Add($el)
                break
            }
        }
    }
    if ($hits.Methods.Count -eq 0 -and $hits.Elems.Count -eq 0 -and $targets.Count -eq 0) {
        Write-Warning 'Nothing matched the keywords or object-name fragments.'
        return
    }
    Write-Host ('Keyword hits: ' + $hits.Methods.Count + ' methods, ' + $hits.Elems.Count + ' other places; object-name hits: ' + $targets.Count + ' objects')

    $sel = @{}; $roles = @{}
    foreach ($m in $hits.Methods) {
        AxRef-Include $sel $roles $m.Elem 'match'
        AxRef-Include $sel $roles $m.Elem 'caller L1'
    }
    foreach ($el in $hits.Elems) { AxRef-Include $sel $roles $el 'match' }
    foreach ($el in $targets)    { AxRef-Include $sel $roles $el 'target' }

    # ---- callers (up)
    $maxPerLevel = 300
    $cSeen = New-Object System.Collections.Generic.HashSet[string]
    $frontier = New-Object System.Collections.Generic.List[object]
    foreach ($m in $hits.Methods) {
        if ($cSeen.Add($m.Elem.Key + '|' + $m.Key)) { $frontier.Add([pscustomobject]@{ Elem = $m.Elem; Method = $m }) }
    }
    # Add both collections explicitly. This avoids PowerShell 5.1's
    # ArgumentException when concatenating certain generic List instances.
    foreach ($el in $hits.Elems) {
        if ($cSeen.Add($el.Key + '|')) { $frontier.Add([pscustomobject]@{ Elem = $el; Method = $null }) }
    }
    foreach ($el in $targets) {
        if ($cSeen.Add($el.Key + '|')) { $frontier.Add([pscustomobject]@{ Elem = $el; Method = $null }) }
    }
    $seeds = New-Object System.Collections.Generic.List[object]
    foreach ($m in $hits.Methods) { $seeds.Add([pscustomobject]@{ Elem = $m.Elem; Method = $m }) }
    foreach ($el in $targets) {
        foreach ($tm in $el.Methods.Values) { $seeds.Add([pscustomobject]@{ Elem = $el; Method = $tm }) }
    }

    for ($lvl = 2; $lvl -le $CallerDepth; $lvl++) {
        if ($frontier.Count -eq 0) { break }
        if ($sel.Count -ge $MaxObjects) { Write-Warning "Object limit ($MaxObjects) reached while following callers."; break }
        $slice = $frontier
        if ($frontier.Count -gt $maxPerLevel) {
            Write-Warning "Caller level $lvl has $($frontier.Count) items, following only the first $maxPerLevel."
            $slice = $frontier.GetRange(0, $maxPerLevel)
        }
        Write-Host "Caller level $lvl ..."
        $found = AxRef-FindCallers $model $slice
        $frontier = New-Object System.Collections.Generic.List[object]
        foreach ($c in $found) {
            $k = $c.Elem.Key + '|'
            if ($null -ne $c.Method) { $k = $c.Elem.Key + '|' + $c.Method.Key }
            if ($cSeen.Add($k)) {
                $frontier.Add($c)
                if ($sel.Count -lt $MaxObjects -or $sel.ContainsKey($c.Elem.Key)) { AxRef-Include $sel $roles $c.Elem ('caller L' + $lvl) }
            }
        }
    }

    # ---- called (down)
    Write-Host 'Following called objects ...'
    AxRef-ExpandCalled $model $seeds $CalledDepth $sel $roles $MaxObjects

    Write-Host ('Objects selected: ' + $sel.Count)

    # ---- write sanitized files
    New-Item -ItemType Directory -Force -Path $out | Out-Null
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $fileMap = @{}
    $wi = 0
    foreach ($el in ($sel.Values | Sort-Object Type, Name)) {
        $wi++
        Write-Progress -Activity 'Writing sanitized files' -Status ('{0} / {1}   {2}' -f $wi, $sel.Count, $el.Name) -PercentComplete ([int](100 * $wi / [Math]::Max(1, $sel.Count)))
        $dir = Join-Path $out $el.Type
        $existingFiles = New-Object System.Collections.Generic.List[string]
        foreach ($suffix in @('', '_top', '_cus')) {
            $candidate = Join-Path $dir ($el.Name + $suffix + '.xpp')
            if (Test-Path -LiteralPath $candidate -PathType Leaf) { $existingFiles.Add($candidate) }
        }
        if ($existingFiles.Count -gt 0) {
            $answer = AxRef-Ask ("Object '{0}\{1}' already exists. Replace its file(s)? (y/N)" -f $el.Type, $el.Name) 'N'
            if ($answer -notmatch '^(y|yes)$') {
                $names = New-Object System.Collections.Generic.List[string]
                foreach ($path in $existingFiles) { $names.Add(($el.Type + '\' + [System.IO.Path]::GetFileName($path))) }
                $fileMap[$el.Key] = $names
                Write-Host ("Kept existing object: {0}\{1}" -f $el.Type, $el.Name)
                continue
            }
            foreach ($path in $existingFiles) { Remove-Item -LiteralPath $path -Force }
        }
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        $both = ($el.TopPath -and $el.CusPath)
        $names = New-Object System.Collections.Generic.List[string]
        foreach ($src in @(@('top', $el.TopPath, '_top', 'TOP'), @('cus', $el.CusPath, '_cus', 'CUS'))) {
            if (-not $src[1]) { continue }
            $suffix = ''
            if ($both) { $suffix = $src[2] }
            $text = [System.IO.File]::ReadAllText($src[1])
            $clean = AxRef-Sanitize $text $el.Type $el.Name $src[3]
            $fn = $el.Name + $suffix + '.xpp'
            [System.IO.File]::WriteAllText((Join-Path $dir $fn), $clean, $utf8)
            $names.Add(($el.Type + '\' + $fn))
        }
        $fileMap[$el.Key] = $names
    }
    Write-Progress -Activity 'Writing sanitized files' -Completed

    # ---- optional project xpo
    $projectFile = ''
    if (-not $Project) {
        $a = AxRef-Ask 'Generate a project XPO you can import into AX? (y/N)' 'N'
        $Project = [bool]($a -match '^(y|yes)$')
    }
    if ($Project) {
        $defName = 'Reference_' + (($kws[0] -replace '[^A-Za-z0-9_]', ''))
        if (-not $ProjectName) { $ProjectName = AxRef-Ask 'Project name' $defName }
        $tpl = AxRef-LoadTemplate
        if ($ProjectSample) { $tpl = AxRef-LearnTemplate $ProjectSample $model }
        if (-not $tpl) {
            Write-Host ''
            Write-Host 'The project xpo layout is learned once from a small project exported from your own AX.'
            Write-Host 'In AX: create a project, add ONE object of each type you care about to it (class, table, form, ...;'
            Write-Host 'use objects that exist in your top/cus folders), then export the project to .xpo.'
            $sp = AxRef-Ask 'Path to that sample project .xpo (Enter to skip)' ''
            if ($sp) { $tpl = AxRef-LearnTemplate $sp $model }
        }
        if ($tpl) {
            $pf = Join-Path $out ($ProjectName + '.xpo')
            $replaceProject = $true
            if (Test-Path -LiteralPath $pf -PathType Leaf) {
                $answer = AxRef-Ask ("Project '{0}' already exists. Replace it? (y/N)" -f $ProjectName) 'N'
                $replaceProject = ($answer -match '^(y|yes)$')
            }
            if ($replaceProject) {
                $missing = AxRef-WriteProject $tpl $ProjectName ($sel.Values) $pf
                $projectFile = $ProjectName + '.xpo'
                if ($missing.Count -gt 0) {
                    Write-Warning ('No template for these object types, so those objects are NOT in the project: ' + (($missing | Sort-Object -Unique) -join '; '))
                }
            }
            else { $projectFile = $ProjectName + '.xpo'; Write-Host 'Kept existing project XPO.' }
        }
        else { Write-Warning 'Project xpo skipped (no template).' }
    }

    $info = [pscustomobject]@{
        Keywords = $kws.ToArray(); ObjectNameContains = $nameFilters.ToArray(); CusObjectPrefix = $CusObjectPrefix
        CallerDepth = $CallerDepth; CalledDepth = $CalledDepth
        When = (Get-Date).ToString('yyyy-MM-dd HH:mm'); ProjectFile = $projectFile
    }
    AxRef-WriteDocs $out $info $sel $roles $fileMap

    Write-Host ''
    Write-Host "Done. $($sel.Count) objects written to: $out"
}

Set-Alias -Name axref -Value Export-AxReference

# forget the learned project layout (e.g. after an AX upgrade)
function Reset-AxRefProjectTemplate {
    $p = Join-Path $script:AxRefDir 'project-template.json'
    if (Test-Path $p) { Remove-Item $p; Write-Host 'Project template removed.' }
}
