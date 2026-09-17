param(
    [string]$FolderA,
    [string]$FolderB,
    [string]$OutputPath,
    [switch]$NoDialog,
    [switch]$LibraryOnly
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# This engine uses only the .NET Framework assemblies supplied with Windows.
if (-not ('ExcelFolderCompare.Engine' -as [type])) {
Add-Type -ReferencedAssemblies System.Xml,System.Core,System.IO.Compression,System.IO.Compression.FileSystem -TypeDefinition @'
using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelFolderCompare {
    public sealed class CellData {
        public string Type = "空白", Value = "", Formula = "", Strike = "なし", CacheType = "空白", Cache = "";
    }
    public sealed class BookData {
        public Dictionary<string, Dictionary<string, CellData>> Sheets = new Dictionary<string, Dictionary<string, CellData>>(StringComparer.Ordinal);
        public List<string> Notes = new List<string>();
        public bool Date1904;
    }
    public sealed class Result {
        public int Files, Differences, Errors, Unsupported;
        public string Output;
    }
    internal sealed class RichValue {
        public string Text;
        public List<KeyValuePair<string, bool?>> Runs = new List<KeyValuePair<string, bool?>>();
    }
    public static class Engine {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        static readonly Regex Address = new Regex(@"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$", RegexOptions.Compiled);
        static readonly Regex SheetPrefix = new Regex(@"\G[A-Za-z0-9_\\][A-Za-z0-9_.\\]*(?::[A-Za-z0-9_\\][A-Za-z0-9_.\\]*)?!", RegexOptions.Compiled);
        static readonly Regex RefToken = new Regex(@"\G(?<![A-Za-z0-9_.])(?:(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)(?![A-Za-z0-9_.(])|(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_.])|(\$?)([1-9][0-9]*):(\$?)([1-9][0-9]*)(?![A-Za-z0-9_.]))", RegexOptions.Compiled);
        static IEnumerable<XmlElement> Children(XmlNode node, string name) {
            if (node != null) foreach (XmlNode child in node.ChildNodes)
                if (child is XmlElement && child.LocalName == name) yield return (XmlElement)child;
        }
        static XmlElement Child(XmlNode node, string name) { return Children(node, name).FirstOrDefault(); }
        static bool On(XmlElement node) {
            if (node == null) return false;
            string value = node.GetAttribute("val");
            return value == "" || value == "1" || value.Equals("true", StringComparison.OrdinalIgnoreCase) || value == "on";
        }
        static XmlDocument ReadXml(ZipArchive zip, string part, bool required) {
            ZipArchiveEntry entry = zip.GetEntry(part);
            if (entry == null) { if (required) throw new InvalidDataException("ファイル内の部品がありません: " + part); return null; }
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 256L * 1024 * 1024 };
            var doc = new XmlDocument { XmlResolver = null };
            using (Stream stream = entry.Open()) using (XmlReader reader = XmlReader.Create(stream, settings)) doc.Load(reader);
            return doc;
        }
        static string ResolvePart(string parent, string target) {
            var uri = new Uri(new Uri("http://package/" + parent), target.Replace('\\', '/'));
            if (uri.Host != "package") throw new InvalidDataException("外部の部品は読み込みません。");
            return Uri.UnescapeDataString(uri.AbsolutePath.TrimStart('/'));
        }
        static Dictionary<string, string[]> Relationships(ZipArchive zip, string part) {
            var result = new Dictionary<string, string[]>();
            XmlDocument doc = ReadXml(zip, part, true);
            foreach (XmlElement rel in Children(doc.DocumentElement, "Relationship")) {
                if (rel.GetAttribute("TargetMode") != "External") result[rel.GetAttribute("Id")] = new [] { rel.GetAttribute("Target"), rel.GetAttribute("Type") };
            }
            return result;
        }
        static string DecodeText(string text) {
            // Single pass preserves literal escapes such as _x005F_x0041_.
            return Regex.Replace(text ?? "", @"_x([0-9A-Fa-f]{4})_", m => ((char)Int32.Parse(m.Groups[1].Value, NumberStyles.HexNumber)).ToString());
        }
        static RichValue Rich(XmlElement element) {
            var result = new RichValue { Text = "" };
            if (element == null) return result;
            foreach (XmlNode node in element.ChildNodes) {
                var el = node as XmlElement; if (el == null) continue;
                if (el.LocalName == "t") {
                    string text = DecodeText(el.InnerText); result.Text += text;
                    result.Runs.Add(new KeyValuePair<string, bool?>(text, null));
                } else if (el.LocalName == "r") {
                    XmlElement t = Child(el, "t");
                    string text = DecodeText(t == null ? "" : t.InnerText); result.Text += text;
                    XmlElement props = Child(el, "rPr");
                    // A run with explicit properties has its own font; no strike means off.
                    bool? strike = props == null ? (bool?)null : On(Child(props, "strike"));
                    result.Runs.Add(new KeyValuePair<string, bool?>(text, strike));
                }
            }
            return result;
        }
        static string Strike(RichValue value, bool basis) {
            if (value == null || value.Text.Length == 0) return basis ? "全体" : "なし";
            var positions = new List<bool>();
            foreach (var run in value.Runs) {
                // Count Unicode code points, so supplementary characters occupy one position.
                for (int i = 0; i < run.Key.Length; i++) {
                    positions.Add(run.Value ?? basis);
                    if (Char.IsHighSurrogate(run.Key[i]) && i + 1 < run.Key.Length && Char.IsLowSurrogate(run.Key[i+1])) i++;
                }
            }
            if (positions.All(x => x)) return "全体";
            if (!positions.Any(x => x)) return "なし";
            var ranges = new List<string>(); int start = -1;
            for (int i = 0; i <= positions.Count; i++) {
                bool on = i < positions.Count && positions[i];
                if (on && start < 0) start = i + 1;
                if (!on && start > 0) { ranges.Add(start == i ? start.ToString() : start + "-" + i); start = -1; }
            }
            return "一部（文字位置: " + String.Join(", ", ranges) + "）";
        }
        static int Col(string letters) { int n = 0; foreach (char c in letters.ToUpperInvariant()) n = n * 26 + c - 'A' + 1; return n; }
        public static string ColName(int n) { string s = ""; while (n > 0) { n--; s = (char)('A' + n % 26) + s; n /= 26; } return s; }
        static int[] ParseAddress(string text) {
            Match m = Address.Match(text);
            if (!m.Success) throw new InvalidDataException("不正なセル位置: " + text);
            int col = Col(m.Groups[1].Value), row = Int32.Parse(m.Groups[2].Value, Inv);
            if (col > 16384 || row > 1048576) throw new InvalidDataException("セル位置が範囲外です: " + text);
            return new [] { row, col };
        }
        static string ShiftCol(string dollar, string col, int offset) {
            int n = Col(col) + (dollar == "$" ? 0 : offset);
            return n < 1 || n > 16384 ? "#REF!" : dollar + ColName(n);
        }
        static string ShiftRow(string dollar, string row, int offset) {
            int n = Int32.Parse(row, Inv) + (dollar == "$" ? 0 : offset);
            return n < 1 || n > 1048576 ? "#REF!" : dollar + n.ToString(Inv);
        }
        public static string Translate(string formula, int dr, int dc) {
            var result = new StringBuilder();
            for (int i = 0; i < formula.Length;) {
                char ch = formula[i];
                // Strings, quoted worksheet names, and structured references must remain literal.
                if (ch == '"' || ch == '\'' || ch == '[') {
                    char close = ch == '[' ? ']' : ch; int depth = 1;
                    result.Append(ch); i++;
                    while (i < formula.Length && depth > 0) {
                        char c = formula[i++]; result.Append(c);
                        if (ch == '[' && c == '[') depth++;
                        if (c == close) {
                            if (ch != '[' && i < formula.Length && formula[i] == close) result.Append(formula[i++]);
                            else depth--;
                        }
                    }
                    continue;
                }
                Match sheetPrefix = SheetPrefix.Match(formula, i);
                if (sheetPrefix.Success && sheetPrefix.Index == i) { result.Append(sheetPrefix.Value); i += sheetPrefix.Length; continue; }
                Match m = RefToken.Match(formula, i);
                // A sheet name such as ABC123! is not a cell reference.
                if (m.Success && m.Index == i && (i + m.Length == formula.Length || formula[i + m.Length] != '!')) {
                    string shifted;
                    if (m.Groups[2].Success) {
                        if (Col(m.Groups[2].Value) > 16384) { result.Append(ch); i++; continue; }
                        string c = ShiftCol(m.Groups[1].Value, m.Groups[2].Value, dc), r = ShiftRow(m.Groups[3].Value, m.Groups[4].Value, dr);
                        shifted = c == "#REF!" || r == "#REF!" ? "#REF!" : c + r;
                    } else if (m.Groups[6].Success) shifted = ShiftCol(m.Groups[5].Value, m.Groups[6].Value, dc) + ":" + ShiftCol(m.Groups[7].Value, m.Groups[8].Value, dc);
                    else shifted = ShiftRow(m.Groups[9].Value, m.Groups[10].Value, dr) + ":" + ShiftRow(m.Groups[11].Value, m.Groups[12].Value, dr);
                    result.Append(shifted); i += m.Length;
                } else { result.Append(ch); i++; }
            }
            return result.ToString();
        }
        static string[] Value(string type, string value, bool present) {
            if (!present) return new [] { "空白", "" };
            if (type == "b") return new [] { "真偽値", value == "1" || value == "true" ? "TRUE" : "FALSE" };
            if (type == "e") return new [] { "Excelエラー", value };
            if (type == "d") return new [] { "日時", value };
            if (type == "str") return new [] { "文字列", DecodeText(value) };
            double number;
            if ((type == "" || type == "n") && Double.TryParse(value, NumberStyles.Float, Inv, out number)) return new [] { "数値", number.ToString("R", Inv) };
            if ((type == "" || type == "n") && value == "") return new [] { "空白", "" };
            throw new InvalidDataException("不明なセル値の形式: " + type);
        }
        static int IntAttr(XmlElement el, string name, int fallback) {
            string value = el.GetAttribute(name); return value == "" ? fallback : Int32.Parse(value, Inv);
        }
        public static BookData ReadBook(string path) {
            var book = new BookData();
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Read)) {
                var roots = Relationships(zip, "_rels/.rels");
                var rootRel = roots.Values.FirstOrDefault(x => x[1].EndsWith("/officeDocument", StringComparison.Ordinal));
                if (rootRel == null) throw new InvalidDataException("Excelブックの定義がありません。");
                string mainPart = ResolvePart("", rootRel[0]);
                XmlDocument wb = ReadXml(zip, mainPart, true);
                var props = Child(wb.DocumentElement, "workbookPr");
                book.Date1904 = props != null && (props.GetAttribute("date1904") == "1" || props.GetAttribute("date1904") == "true");
                string dir = mainPart.Substring(0, mainPart.LastIndexOf('/') + 1);
                var rels = Relationships(zip, dir + "_rels/" + mainPart.Substring(dir.Length) + ".rels");
                var shared = new List<RichValue>();
                var sharedRel = rels.Values.FirstOrDefault(x => x[1].EndsWith("/sharedStrings", StringComparison.Ordinal));
                if (sharedRel != null) {
                    var ss = ReadXml(zip, ResolvePart(mainPart, sharedRel[0]), true);
                    foreach (XmlElement si in Children(ss.DocumentElement, "si")) shared.Add(Rich(si));
                }
                var styleStrike = new List<bool>(); var differentialStrike = new HashSet<int>();
                var stylesRel = rels.Values.FirstOrDefault(x => x[1].EndsWith("/styles", StringComparison.Ordinal));
                if (stylesRel != null) {
                    var styles = ReadXml(zip, ResolvePart(mainPart, stylesRel[0]), true);
                    var fonts = Children(Child(styles.DocumentElement, "fonts"), "font").Select(x => On(Child(x, "strike"))).ToList();
                    foreach (XmlElement xf in Children(Child(styles.DocumentElement, "cellXfs"), "xf")) {
                        int fontId = IntAttr(xf, "fontId", 0);
                        if (fontId < 0 || fontId >= fonts.Count) throw new InvalidDataException("不正なフォント参照です。");
                        styleStrike.Add(fonts[fontId]);
                    }
                    int dxf = 0;
                    foreach (XmlElement el in Children(Child(styles.DocumentElement, "dxfs"), "dxf")) {
                        if (Child(Child(el, "font"), "strike") != null) differentialStrike.Add(dxf); dxf++;
                    }
                }
                if (styleStrike.Count == 0) styleStrike.Add(false);
                foreach (XmlElement sheet in Children(Child(wb.DocumentElement, "sheets"), "sheet")) {
                    string name = sheet.GetAttribute("name"), id = "";
                    foreach (XmlAttribute attr in sheet.Attributes) if (attr.LocalName == "id") id = attr.Value;
                    if (!rels.ContainsKey(id)) throw new InvalidDataException("シートの参照がありません: " + name);
                    if (!rels[id][1].EndsWith("/worksheet", StringComparison.Ordinal)) { book.Notes.Add(name + ": ワークシート以外のシートは対象外です。"); continue; }
                    var xml = ReadXml(zip, ResolvePart(mainPart, rels[id][0]), true);
                    var allCells = Children(Child(xml.DocumentElement, "sheetData"), "row").SelectMany(x => Children(x, "c")).ToList();
                    var anchors = new Dictionary<string, string[]>();
                    foreach (var cell in allCells) {
                        var f = Child(cell, "f");
                        if (f != null && f.GetAttribute("t") == "shared" && f.InnerText != "") anchors[f.GetAttribute("si")] = new [] { cell.GetAttribute("r"), f.InnerText };
                    }
                    var cells = new Dictionary<string, CellData>(StringComparer.Ordinal);
                    foreach (var cell in allCells) {
                        string address = cell.GetAttribute("r").ToUpperInvariant(); ParseAddress(address);
                        int s = IntAttr(cell, "s", 0);
                        if (s < 0 || s >= styleStrike.Count) throw new InvalidDataException("不正な書式参照です: " + address);
                        string type = cell.GetAttribute("t"); var v = Child(cell, "v"); var f = Child(cell, "f");
                        var data = new CellData(); RichValue rich = null;
                        if (f != null) {
                            string ft = f.GetAttribute("t"), formula = f.InnerText;
                            if (ft == "dataTable") throw new InvalidDataException(name + "!" + address + ": What-Ifデータテーブルは対象外です。");
                            if (ft == "shared" && formula == "") {
                                string key = f.GetAttribute("si");
                                if (!anchors.ContainsKey(key)) throw new InvalidDataException("共有数式の基準セルがありません: " + address);
                                int[] origin = ParseAddress(anchors[key][0]), target = ParseAddress(address);
                                formula = Translate(anchors[key][1], target[0] - origin[0], target[1] - origin[1]);
                            }
                            data.Type = "数式"; data.Formula = "=" + formula;
                            if (ft == "array") data.Formula += " [配列範囲:" + f.GetAttribute("ref") + "]";
                            if (v == null || (v.InnerText == "" && type != "str")) data.CacheType = "保存値なし";
                            else { string[] val = Value(type, v.InnerText, true); data.CacheType = val[0]; data.Cache = val[1]; }
                        } else if (type == "s") {
                            int index;
                            if (v == null || !Int32.TryParse(v.InnerText, out index) || index < 0 || index >= shared.Count) throw new InvalidDataException("不正な文字列参照です: " + address);
                            rich = shared[index]; data.Type = "文字列"; data.Value = rich.Text;
                        } else if (type == "inlineStr") {
                            rich = Rich(Child(cell, "is")); data.Type = "文字列"; data.Value = rich.Text;
                        } else {
                            string[] val = Value(type, v == null ? "" : v.InnerText, v != null); data.Type = val[0]; data.Value = val[1];
                        }
                        data.Strike = Strike(rich, styleStrike[s]);
                        if (data.Type != "空白" || data.Formula != "" || data.Strike != "なし") cells.Add(address, data);
                    }
                    foreach (XmlElement cf in xml.GetElementsByTagName("cfRule", xml.DocumentElement.NamespaceURI)) {
                        if (differentialStrike.Contains(IntAttr(cf, "dxfId", -1))) { book.Notes.Add(name + ": 条件付き書式による取消線は対象外です。"); break; }
                    }
                    foreach (XmlElement el in xml.GetElementsByTagName("row", xml.DocumentElement.NamespaceURI).Cast<XmlElement>().Concat(xml.GetElementsByTagName("col", xml.DocumentElement.NamespaceURI).Cast<XmlElement>())) {
                        int s = IntAttr(el, el.LocalName == "row" ? "s" : "style", -1);
                        if (s >= 0 && s < styleStrike.Count && styleStrike[s]) { book.Notes.Add(name + ": 行・列の取消線設定あり。セルに保存された書式のみ比較します。"); break; }
                    }
                    if (xml.GetElementsByTagName("extLst", xml.DocumentElement.NamespaceURI).Count > 0) book.Notes.Add(name + ": 拡張設定あり。拡張形式の条件付き書式などは評価しません。");
                    book.Sheets.Add(name, cells);
                }
            }
            return book;
        }
        static string EscapeText(string text) {
            text = Regex.Replace(text ?? "", @"_x[0-9A-Fa-f]{4}_", m => "_x005F_" + m.Value.Substring(1));
            var b = new StringBuilder();
            for (int i = 0; i < text.Length; i++) {
                char c = text[i];
                if (Char.IsHighSurrogate(c) && i+1 < text.Length && Char.IsLowSurrogate(text[i+1])) { b.Append(c); b.Append(text[++i]); }
                else if (Char.IsSurrogate(c) || (c < 32 && c != '\t' && c != '\n') || c == '\ufffe' || c == '\uffff') b.Append("_x" + ((int)c).ToString("X4") + "_");
                else b.Append(c);
            }
            return b.ToString();
        }
        static void XmlPart(ZipArchive zip, string name, Action<XmlWriter> write) {
            var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
            using (var stream = entry.Open()) using (var x = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), CheckCharacters = true })) { x.WriteStartDocument(true); write(x); x.WriteEndDocument(); }
        }
        static void Element(XmlWriter x, string name, params string[] attrs) {
            x.WriteStartElement(name); for (int i=0; i<attrs.Length; i+=2) x.WriteAttributeString(attrs[i], attrs[i+1]); x.WriteEndElement();
        }
        static void Sheet(ZipArchive zip, string part, string[] headers, List<object[]> rows, double[] widths) {
            if (rows.Count >= 1048576) throw new InvalidOperationException("差分がExcelの行数上限を超えました。入力フォルダを分けてください。");
            XmlPart(zip, part, x => {
                x.WriteStartElement("worksheet", Main);
                Element(x, "dimension", "ref", "A1:" + ColName(headers.Length) + (rows.Count+1));
                x.WriteStartElement("sheetViews"); x.WriteStartElement("sheetView"); x.WriteAttributeString("workbookViewId", "0"); x.WriteAttributeString("showGridLines", "0");
                Element(x,"pane","ySplit","1","topLeftCell","A2","activePane","bottomLeft","state","frozen");
                Element(x,"selection","pane","bottomLeft","activeCell","A2","sqref","A2"); x.WriteEndElement(); x.WriteEndElement();
                x.WriteStartElement("cols"); for(int i=0;i<widths.Length;i++) Element(x,"col","min",(i+1).ToString(),"max",(i+1).ToString(),"width",widths[i].ToString(Inv),"customWidth","1"); x.WriteEndElement();
                x.WriteStartElement("sheetData");
                for(int r=0;r<=rows.Count;r++) {
                    object[] record = r==0 ? headers.Cast<object>().ToArray() : rows[r-1];
                    int lines=1;
                    for(int c=0;c<record.Length;c++) {
                        string text=Convert.ToString(record[c],Inv) ?? "";
                        if(text.Length>32767) throw new InvalidOperationException("結果の1セルがExcelの文字数上限を超えました。");
                        int count=0; foreach(string line in text.Split('\n')) count+=Math.Max(1,(int)Math.Ceiling(line.Length/(widths[c]/2)));
                        lines=Math.Max(lines,count);
                    }
                    x.WriteStartElement("row"); x.WriteAttributeString("r",(r+1).ToString()); x.WriteAttributeString("ht",Math.Min(409,Math.Max(30,16*lines+8)).ToString()); x.WriteAttributeString("customHeight","1");
                    for(int c=0;c<record.Length;c++) {
                        x.WriteStartElement("c"); x.WriteAttributeString("r",ColName(c+1)+(r+1));
                        x.WriteAttributeString("s",r==0 ? "1" : (headers.Length==14 && (c==8 || c==9)) ? "3" : r%2==1 ? "2" : "0");
                        if(record[c] is int) { x.WriteElementString("v",record[c].ToString()); }
                        else { x.WriteAttributeString("t","inlineStr"); x.WriteStartElement("is"); x.WriteStartElement("t"); x.WriteAttributeString("xml","space",null,"preserve"); x.WriteString(EscapeText(Convert.ToString(record[c],Inv))); x.WriteEndElement(); x.WriteEndElement(); }
                        x.WriteEndElement();
                    }
                    x.WriteEndElement();
                }
                x.WriteEndElement(); Element(x,"autoFilter","ref","A1:"+ColName(headers.Length)+(rows.Count+1)); x.WriteEndElement();
            });
        }
        static void WriteReport(string output, List<object[]> summary, List<object[]> details, List<object[]> conditions) {
            bool created=false;
            try {
                using(var stream=new FileStream(output,FileMode.CreateNew,FileAccess.Write,FileShare.None)) {
                    created=true;
                    using(var zip=new ZipArchive(stream,ZipArchiveMode.Create)) {
                        XmlPart(zip,"[Content_Types].xml",x=>{
                            x.WriteStartElement("Types","http://schemas.openxmlformats.org/package/2006/content-types");
                            Element(x,"Default","Extension","rels","ContentType","application/vnd.openxmlformats-package.relationships+xml");
                            Element(x,"Default","Extension","xml","ContentType","application/xml");
                            Element(x,"Override","PartName","/xl/workbook.xml","ContentType","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
                            Element(x,"Override","PartName","/xl/styles.xml","ContentType","application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
                            for(int i=1;i<=3;i++) Element(x,"Override","PartName","/xl/worksheets/sheet"+i+".xml","ContentType","application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"); x.WriteEndElement();
                        });
                        XmlPart(zip,"_rels/.rels",x=>{x.WriteStartElement("Relationships","http://schemas.openxmlformats.org/package/2006/relationships");Element(x,"Relationship","Id","rId1","Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument","Target","xl/workbook.xml");x.WriteEndElement();});
                        XmlPart(zip,"xl/workbook.xml",x=>{
                            x.WriteStartElement("workbook",Main); x.WriteAttributeString("xmlns","r",null,"http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                            x.WriteStartElement("sheets"); int i=1; foreach(string name in new[]{"比較一覧","差分詳細","比較条件"}) {x.WriteStartElement("sheet");x.WriteAttributeString("name",name);x.WriteAttributeString("sheetId",i.ToString());x.WriteAttributeString("r","id",null,"rId"+i);x.WriteEndElement();i++;} x.WriteEndElement(); x.WriteEndElement();
                        });
                        XmlPart(zip,"xl/_rels/workbook.xml.rels",x=>{
                            x.WriteStartElement("Relationships","http://schemas.openxmlformats.org/package/2006/relationships");
                            for(int i=1;i<=3;i++) Element(x,"Relationship","Id","rId"+i,"Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet","Target","worksheets/sheet"+i+".xml");
                            Element(x,"Relationship","Id","rId4","Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles","Target","styles.xml");x.WriteEndElement();
                        });
                        XmlPart(zip,"xl/styles.xml",x=>{
                            x.WriteStartElement("styleSheet",Main);
                            x.WriteStartElement("fonts");x.WriteAttributeString("count","2");for(int i=0;i<2;i++){x.WriteStartElement("font");Element(x,"sz","val","10");Element(x,"name","val","Yu Gothic");Element(x,"color","rgb",i==0?"FF203047":"FFFFFFFF");if(i==1)Element(x,"b");x.WriteEndElement();}x.WriteEndElement();
                            x.WriteStartElement("fills");x.WriteAttributeString("count","5");foreach(string color in new[]{"none","gray125","FF234B68","FFF0F5F8","FFFFF0CC"}){x.WriteStartElement("fill");x.WriteStartElement("patternFill");x.WriteAttributeString("patternType",color.Length==8?"solid":color);if(color.Length==8){Element(x,"fgColor","rgb",color);Element(x,"bgColor","indexed","64");}x.WriteEndElement();x.WriteEndElement();}x.WriteEndElement();
                            x.WriteStartElement("borders");x.WriteAttributeString("count","1");x.WriteStartElement("border");foreach(string edge in new[]{"left","right","top","bottom","diagonal"})Element(x,edge);x.WriteEndElement();x.WriteEndElement();
                            x.WriteStartElement("cellStyleXfs");x.WriteAttributeString("count","1");Element(x,"xf","numFmtId","0","fontId","0","fillId","0","borderId","0");x.WriteEndElement();
                            x.WriteStartElement("cellXfs");x.WriteAttributeString("count","4");for(int i=0;i<4;i++){x.WriteStartElement("xf");x.WriteAttributeString("numFmtId","0");x.WriteAttributeString("fontId",i==1?"1":"0");x.WriteAttributeString("fillId",i==0?"0":(i+1).ToString());x.WriteAttributeString("borderId","0");x.WriteAttributeString("xfId","0");x.WriteAttributeString("applyAlignment","1");Element(x,"alignment","vertical","top","wrapText","1");x.WriteEndElement();}x.WriteEndElement();
                            x.WriteStartElement("cellStyles");x.WriteAttributeString("count","1");Element(x,"cellStyle","name","Normal","xfId","0","builtinId","0");x.WriteEndElement();x.WriteEndElement();
                        });
                        Sheet(zip,"xl/worksheets/sheet1.xml",new[]{"ファイル名","結果","差分行数","備考"},summary,new double[]{35,25,15,90});
                        Sheet(zip,"xl/worksheets/sheet2.xml",new[]{"ファイル名","シート名","セル","差分の種類","Aの値","Bの値","Aの数式","Bの数式","Aの取消線","Bの取消線","Aの型","Bの型","Aの保存済み計算結果","Bの保存済み計算結果"},details,new double[]{28,20,12,26,36,36,36,36,30,30,15,15,30,30});
                        Sheet(zip,"xl/worksheets/sheet3.xml",new[]{"項目","内容"},conditions,new double[]{25,110});
                    }
                }
            } catch { if(created && File.Exists(output)) File.Delete(output); throw; }
        }
        static Dictionary<string,string> Files(string folder,string output) {
            var files=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
            foreach(string path in Directory.GetFiles(folder)) {
                string name=Path.GetFileName(path), ext=Path.GetExtension(path).ToLowerInvariant();
                if(name.StartsWith("~$",StringComparison.Ordinal) || String.Equals(path,output,StringComparison.OrdinalIgnoreCase)) continue;
                if(ext==".xlsx" || ext==".xlsm" || ext==".xls" || ext==".xlsb") files.Add(name,path);
            }
            return files;
        }
        public static Result Compare(string folderA,string folderB,string output) {
            folderA=Path.GetFullPath(folderA); folderB=Path.GetFullPath(folderB); output=Path.GetFullPath(output);
            if(!Directory.Exists(folderA) || !Directory.Exists(folderB)) throw new ArgumentException("A・Bには存在するフォルダを指定してください。");
            if(String.Equals(folderA.TrimEnd('\\','/'),folderB.TrimEnd('\\','/'),StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("A・Bには別々のフォルダを指定してください。");
            if(!Path.GetExtension(output).Equals(".xlsx",StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("出力ファイルは.xlsxを指定してください。");
            if(File.Exists(output)) throw new IOException("出力先は既に存在します。新しいファイル名を指定してください。");
            if(!Directory.Exists(Path.GetDirectoryName(output))) throw new IOException("保存先フォルダが存在しません。");
            var a=Files(folderA,output);var b=Files(folderB,output);
            var names=a.Keys.Union(b.Keys,StringComparer.OrdinalIgnoreCase).OrderBy(x=>x,StringComparer.OrdinalIgnoreCase).ToList();
            if(names.Count==0)throw new IOException("指定フォルダの直下にExcelファイルがありません。");
            var summary=new List<object[]>();var details=new List<object[]>();var result=new Result{Files=names.Count,Output=output};
            foreach(string name in names){
                Console.WriteLine("比較中: "+name);
                string ext=Path.GetExtension(name).ToLowerInvariant();
                if(ext==".xls" || ext==".xlsb"){summary.Add(new object[]{name,"対象外","",".xlsxまたは.xlsmに変換してください。"});result.Unsupported++;continue;}
                if(!a.ContainsKey(name)||!b.ContainsKey(name)){summary.Add(new object[]{name,a.ContainsKey(name)?"Aのみ":"Bのみ","","同名ファイルがありません。"});continue;}
                int start=details.Count;
                try {
                    var ba=ReadBook(a[name]);var bb=ReadBook(b[name]);
                    foreach(string sn in ba.Sheets.Keys.Union(bb.Sheets.Keys,StringComparer.Ordinal).OrderBy(x=>x,StringComparer.Ordinal)){
                        if(!ba.Sheets.ContainsKey(sn)||!bb.Sheets.ContainsKey(sn)){
                            details.Add(new object[]{name,sn,"",ba.Sheets.ContainsKey(sn)?"Aのみのシート":"Bのみのシート","","","","","","","","","",""});continue;
                        }
                        foreach(string address in ba.Sheets[sn].Keys.Union(bb.Sheets[sn].Keys).OrderBy(x=>(long)ParseAddress(x)[0]*16385+ParseAddress(x)[1])){
                            CellData ca,cb;if(!ba.Sheets[sn].TryGetValue(address,out ca))ca=new CellData();if(!bb.Sheets[sn].TryGetValue(address,out cb))cb=new CellData();
                            var changes=new List<string>();
                            if(ca.Type!=cb.Type||ca.Value!=cb.Value)changes.Add("値・型");
                            if(ca.Formula!=cb.Formula)changes.Add("数式");
                            if(ca.Strike!=cb.Strike)changes.Add("取消線");
                            if(ca.CacheType!=cb.CacheType||ca.Cache!=cb.Cache)changes.Add("保存済み計算結果");
                            if(changes.Count>0)details.Add(new object[]{name,sn,address,String.Join("・",changes),ca.Value,cb.Value,ca.Formula,cb.Formula,ca.Strike,cb.Strike,ca.Type,cb.Type,ca.CacheType=="保存値なし"?"（保存値なし）":ca.Cache,cb.CacheType=="保存値なし"?"（保存値なし）":cb.Cache});
                        }
                    }
                    var notes=ba.Notes.Concat(bb.Notes).Distinct().ToList();
                    if(ba.Date1904!=bb.Date1904)notes.Add("A・Bの日付システムが異なります。日付は内部シリアル値で比較するため確認が必要です。");
                    int count=details.Count-start;string status=count>0?"差分あり":"差分なし";
                    if(notes.Count>0)status+="（注意あり）";
                    summary.Add(new object[]{name,status,count,String.Join("\n",notes)});
                } catch(Exception ex){details.RemoveRange(start,details.Count-start);summary.Add(new object[]{name,"エラー","",ex.Message});result.Errors++;}
            }
            var conditions=new List<object[]>{
                new object[]{"Aフォルダ",folderA},new object[]{"Bフォルダ",folderB},new object[]{"実行日時",DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")},
                new object[]{"比較方法","直下の同名ファイル（拡張子込み、英字大小を区別しない）→同名シート→同じセル位置。サブフォルダは含みません。"},
                new object[]{"対象","値・型、数式、保存済み計算結果、直接設定された取消線（文字の一部を含む）。文字位置は先頭を1とします。"},
                new object[]{"対象外","色・罫線・その他書式、条件付き書式の表示結果、行・列の書式継承、画像・グラフ・コメント・マクロ。"},
                new object[]{"数式","再計算せず保存済み計算結果を比較します。結果未保存は「保存値なし」。共有数式はセル位置に合わせて展開します。"},
                new object[]{"日時・数値","日付・時刻は原則Excel内部のシリアル値で比較・表示します。数値と文字列、空白とゼロを区別します。数値の許容誤差は設けません。"},
                new object[]{"差分行数","変更セル1件を1行、片方だけのシート1件を1行で記録します。行挿入や並べ替えを追跡しません。"},
                new object[]{"エラー","暗号化・破損ファイルやWhat-Ifデータテーブルなどは比較一覧にエラーとして記録します。"}
            };
            WriteReport(output,summary,details,conditions);result.Differences=details.Count;return result;
        }
    }
}
'@
}
if ($LibraryOnly) { return }

function Select-CompareFolder([string]$Title) {
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Title
    try { if ($dialog.ShowDialog() -eq 'OK') { return $dialog.SelectedPath }; return '' }
    finally { $dialog.Dispose() }
}

try {
    if (-not $NoDialog) { Add-Type -AssemblyName System.Windows.Forms }
    if (-not $FolderA) {
        if ($NoDialog) { throw 'FolderAを指定してください。' }
        $FolderA = Select-CompareFolder 'Aフォルダ（比較元）を選択してください'
        if (-not $FolderA) { return }
    }
    if (-not $FolderB) {
        if ($NoDialog) { throw 'FolderBを指定してください。' }
        $FolderB = Select-CompareFolder 'Bフォルダ（比較先）を選択してください'
        if (-not $FolderB) { return }
    }
    if (-not $OutputPath) {
        if ($NoDialog) { throw 'OutputPathを指定してください。' }
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Title = '結果の保存先（新しいファイル名）'
        $dialog.Filter = 'Excel ブック (*.xlsx)|*.xlsx'
        $dialog.FileName = 'Excel差分_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.xlsx'
        try { if ($dialog.ShowDialog() -ne 'OK') { return }; $OutputPath = $dialog.FileName }
        finally { $dialog.Dispose() }
    }
    Write-Host '比較を開始します。ファイルが大きい場合は時間がかかります。'
    $result = [ExcelFolderCompare.Engine]::Compare($FolderA, $FolderB, $OutputPath)
    $message = "保存しました。`n$($result.Output)`nファイル: $($result.Files)件 / 差分: $($result.Differences)行 / エラー: $($result.Errors)件 / 対象外: $($result.Unsupported)件"
    Write-Host $message
    if (-not $NoDialog) { [void][System.Windows.Forms.MessageBox]::Show($message, 'Excel比較') }
    if ($NoDialog) { $result }
} catch {
    $message = $_.Exception.GetBaseException().Message
    Write-Host $message -ForegroundColor Red
    if (-not $NoDialog) { [void][System.Windows.Forms.MessageBox]::Show($message, 'Excel比較 エラー') }
    exit 1
}
