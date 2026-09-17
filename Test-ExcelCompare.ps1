$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ExcelCompare.ps1') -LibraryOnly

function Assert-Equal($Actual, $Expected, [string]$Label) {
    if ($Actual -cne $Expected) { throw "$Label : expected [$Expected], actual [$Actual]" }
}
function Write-ZipParts([string]$Path, [hashtable]$Parts) {
    $file = [IO.File]::Open($Path, [IO.FileMode]::CreateNew)
    $zip = New-Object IO.Compression.ZipArchive($file, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($part in $Parts.GetEnumerator()) {
            $entry = $zip.CreateEntry($part.Key)
            $writer = New-Object IO.StreamWriter($entry.Open(), (New-Object Text.UTF8Encoding $false))
            try { $writer.Write([string]$part.Value) } finally { $writer.Dispose() }
        }
    } finally { $zip.Dispose(); $file.Dispose() }
}
function New-Fixture([string]$Path, [string]$Side) {
    $s = if ($Side -eq 'A') { '0' } else { '1' }
    $richExtra = if ($Side -eq 'B') { '<strike/>' } else { '' }
    $formula = if ($Side -eq 'A') { '1+1' } else { '2+0' }
    $text = if ($Side -eq 'A') { 'ABC' } else { 'abc' }
    $zero = if ($Side -eq 'A') { '<c r="A9"><v>0</v></c>' } else { '' }
    $type = if ($Side -eq 'A') { '<c r="A10" t="inlineStr"><is><t>1</t></is></c>' } else { '<c r="A10"><v>1</v></c>' }
    $bool = if ($Side -eq 'A') { 'b' } else { 'n' }
    $trailing = if ($Side -eq 'A') { 'has space ' } else { 'has space' }
    $newline = if ($Side -eq 'A') { 'a&#10;b' } else { 'ab' }
    $cache = if ($Side -eq 'A') { '10' } else { '11' }
    $sharedFormula = if ($Side -eq 'A') { '<c r="B1"><f t="shared" si="0" ref="B1:B2">A1</f><v/></c>' } else { '<c r="B1"><f>A1</f><v/></c>' }
    $sharedFollower = if ($Side -eq 'A') { '<c r="B2"><f t="shared" si="0"/><v/></c>' } else { '<c r="B2"><f>A2</f><v/></c>' }
    $parts = @{
        '[Content_Types].xml' = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>'
        '_rels/.rels' = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        'xl/workbook.xml' = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="s1"/><sheet name="' + $Side + '_only" sheetId="2" r:id="s2"/></sheets></workbook>'
        'xl/_rels/workbook.xml.rels' = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="s1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="s2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="style" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="strings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'
        'xl/styles.xml' = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><name val="Yu Gothic"/></font><font><name val="Yu Gothic"/><strike/></font><font><name val="Yu Gothic"/><color rgb="FFFF0000"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf fontId="0"/></cellStyleXfs><cellXfs count="3"><xf fontId="0"/><xf fontId="1"/><xf fontId="2"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="1"><dxf><font><strike/></font></dxf></dxfs></styleSheet>'
        'xl/sharedStrings.xml' = '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><r><t>a</t></r><r><rPr><strike/></rPr><t>bc</t></r><r><t>d</t></r><r><rPr>' + $richExtra + '</rPr><t>e</t></r><r><t>f</t></r></si><si><t>_x005F_x0041_ &amp; &lt; &gt; _x000D_</t></si></sst>'
        'xl/worksheets/sheet2.xml' = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'
        'xl/worksheets/sheet1.xml' = @"
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>same</t></is></c>$sharedFormula</row>
<row r="2"><c r="A2" s="$s" t="inlineStr"><is><t>strike</t></is></c>$sharedFollower</row>
<row r="3"><c r="A3" t="s"><v>0</v></c></row>
<row r="4"><c r="A4"><f>$formula</f><v/></c></row>
<row r="5"><c r="A5" t="inlineStr"><is><t>$text</t></is></c></row>
<row r="6"><c r="A6" s="$s"/></row>
<row r="7"><c r="A7" t="inlineStr"><is><t>=literal</t></is></c></row>
<row r="8"><c r="A8" s="2" t="inlineStr"><is><t>same</t></is></c></row>
<row r="9">$zero</row><row r="10">$type</row>
<row r="11"><c r="A11" t="$bool"><v>1</v></c></row>
<row r="12"><c r="A12" t="inlineStr"><is><t xml:space="preserve">$trailing</t></is></c></row>
<row r="13"><c r="A13" t="inlineStr"><is><t>$newline</t></is></c></row>
<row r="14"><c r="A14" t="s"><v>1</v></c></row>
<row r="15"><c r="A15"><f>9+1</f><v>$cache</v></c></row>
<row r="16"><c r="A16" t="inlineStr"><is><r><rPr><strike/></rPr><t>😀</t></r><r><t>日本語</t></r></is></c></row>
<row r="17"><c r="A17"><v>1.000</v></c></row>
<row r="1000000"><c r="Z1000000" s="2"/></row>
</sheetData><conditionalFormatting sqref="A1"><cfRule type="cellIs" operator="equal" priority="1" dxfId="0"><formula>1</formula></cfRule></conditionalFormatting></worksheet>
"@
    }
    Write-ZipParts $Path $parts
}

$testRoot = Join-Path $PSScriptRoot ('work/test_' + [Guid]::NewGuid().ToString('N').Substring(0, 10))
$a = Join-Path $testRoot 'A'; $b = Join-Path $testRoot 'B'
[void](New-Item -ItemType Directory -Force -Path $a,$b)
New-Fixture (Join-Path $a 'match.xlsx') 'A'
New-Fixture (Join-Path $b 'match.xlsx') 'B'
Copy-Item -LiteralPath (Join-Path $a 'match.xlsx') -Destination (Join-Path $a 'same.xlsx')
Copy-Item -LiteralPath (Join-Path $a 'match.xlsx') -Destination (Join-Path $b 'same.xlsx')
Copy-Item -LiteralPath (Join-Path $a 'match.xlsx') -Destination (Join-Path $a 'only_a.xlsx')
Copy-Item -LiteralPath (Join-Path $b 'match.xlsx') -Destination (Join-Path $b 'only_b.xlsx')
foreach ($folder in @($a,$b)) {
    [IO.File]::WriteAllText((Join-Path $folder 'broken.xlsx'), 'broken')
    [IO.File]::WriteAllText((Join-Path $folder 'legacy.xls'), 'legacy')
    [IO.File]::WriteAllText((Join-Path $folder '~$ignored.xlsx'), 'temporary')
}
$before = @(Get-ChildItem $a,$b -File | Get-FileHash | Sort-Object Path | ForEach-Object { $_.Hash }) -join ','
$output = Join-Path $testRoot 'result.xlsx'
$result = [ExcelFolderCompare.Engine]::Compare($a,$b,$output)
Assert-Equal $result.Files 6 'file count'
Assert-Equal $result.Differences 13 'difference count'
Assert-Equal $result.Errors 1 'error count'
Assert-Equal $result.Unsupported 1 'unsupported count'
$report = [ExcelFolderCompare.Engine]::ReadBook($output)
Assert-Equal $report.Sheets.Count 3 'report sheets'
$detail = $report.Sheets['差分詳細']
$found = @{}
for ($row=2; $row -le 14; $row++) {
    $address = $detail['C'+$row].Value
    if ($address) { $found[$address] = $row }
}
Assert-Equal (($found.Keys | Sort-Object) -join ',') 'A10,A11,A12,A13,A15,A2,A3,A4,A5,A6,A9' 'changed addresses'
$row = $found['A3']
Assert-Equal $detail['I'+$row].Value '一部（文字位置: 2-3）' 'partial strike A'
Assert-Equal $detail['J'+$row].Value '一部（文字位置: 2-3, 5）' 'partial strike B'
$row = $found['A4']
Assert-Equal $detail['G'+$row].Value '=1+1' 'literal formula'
Assert-Equal $detail['G'+$row].Type '文字列' 'literal formula type'
$source = [ExcelFolderCompare.Engine]::ReadBook((Join-Path $a 'match.xlsx'))
Assert-Equal $source.Sheets['Data']['A16'].Strike '一部（文字位置: 1）' 'supplementary character'
Assert-Equal $source.Sheets['Data']['A14'].Value "_x0041_ & < > `r" 'OOXML text escapes'
Assert-Equal $source.Sheets['Data']['B2'].Formula '=A2' 'shared formula'
Assert-Equal $source.Sheets['Data']['A17'].Value '1' 'numeric normalization'
Assert-Equal $source.Notes.Count 1 'conditional strike warning'
Assert-Equal ([ExcelFolderCompare.Engine]::Translate('SUM(A1,$B2,C$3,$D$4,A:A,1:1,"A1",Table1[A1],''A1 sheet''!A1,LOG10(A1),ABC123!A1)',1,1)) 'SUM(B2,$B3,D$3,$D$4,B:B,2:2,"A1",Table1[A1],''A1 sheet''!B2,LOG10(B2),ABC123!B2)' 'shared reference translation'
Assert-Equal ([ExcelFolderCompare.Engine]::Translate('A1',-1,0)) '#REF!' 'reference bounds'
Assert-Equal ([ExcelFolderCompare.Engine]::Translate('SUM(A1:B2!C3)',1,1)) 'SUM(A1:B2!D4)' 'three-dimensional sheet reference'
$blocked = $false
try { [void][ExcelFolderCompare.Engine]::Compare($a,$b,$output) } catch { $blocked = $true }
Assert-Equal $blocked $true 'overwrite protection'
$after = @(Get-ChildItem $a,$b -File | Get-FileHash | Sort-Object Path | ForEach-Object { $_.Hash }) -join ','
Assert-Equal $after $before 'source files unchanged'
Write-Host 'PASS: 比較・取消線・共有数式・Excel出力・エラー記録・上書き防止・入力非変更'
Write-Host "テスト結果: $output"
