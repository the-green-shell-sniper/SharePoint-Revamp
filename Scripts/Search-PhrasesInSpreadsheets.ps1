<#
.SYNOPSIS
    Searches the contents of every CSV / Excel file under a folder tree for
    each phrase listed in a search-phrases CSV, falling back to fuzzy
    matching for phrases that have no exact match.

.DESCRIPTION
    1) Loads the search phrases (one per row, first column; no header row).
       Handles CSV quoting and ANSI (Windows-1252) or UTF-8 encoded files.
    2) Recursively finds *.csv, *.xlsx, *.xlsm, *.xls and *.xlsb files under
       SearchRoot (Excel lock files "~$*" and anything under OutputFolder are
       skipped).
    3) Pass 1 - exact: every cell is normalized (URL-decoded, lower-cased,
       backslashes -> forward slashes, whitespace collapsed) and checked for
       each phrase as a substring. So "My File.docx" matches a cell holding
       ".../Shared%20Documents/My%20File.docx".
    4) Pass 2 - fuzzy: only for phrases with zero exact matches, every cell is
       searched for the closest approximate substring (edit distance). A hit
       is kept when Score = 1 - (edits / phrase length) >= FuzzyThreshold.
       The best MaxFuzzyHitsPerPhrase hits per phrase are reported.

    .xlsx/.xlsm are read directly (no Excel needed). .xls/.xlsb require
    Microsoft Excel to be installed; otherwise they're logged as skipped.

.PARAMETER SearchRoot
    Folder to search (recursively).

.PARAMETER PhrasesCsvPath
    CSV with one search phrase per row (first column, no header).

.PARAMETER OutputFolder
    Where the Results / Errors CSVs are written.

.PARAMETER FuzzyThreshold
    Minimum fuzzy score (0-1) to report. 0.80 = up to 20% of the phrase's
    characters may differ. Default 0.80.

.PARAMETER MaxFuzzyHitsPerPhrase
    Maximum number of fuzzy hits reported per phrase (highest scores first).
    Default 25.

.OUTPUTS
    PhraseSearch-Results-<timestamp>.csv  - SearchPhrase, MatchType (Exact/Fuzzy),
                                            FilePath, Sheet (Excel only), Row.
                                            One line per phrase/file/row.
    PhraseSearch-Errors-<timestamp>.csv   - only if some files couldn't be read.

.EXAMPLE
    .\Search-PhrasesInSpreadsheets.ps1

.EXAMPLE
    .\Search-PhrasesInSpreadsheets.ps1 -FuzzyThreshold 0.7 -MaxFuzzyHitsPerPhrase 50
#>
[CmdletBinding()]
param(
    [string]$SearchRoot = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\DocumentUrlLookup-OutputFiles",

    [string]$PhrasesCsvPath = "P:\IT\Private\zzBPosey\whaitapp09\searchphrases.csv",

    [string]$OutputFolder = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\PhraseSearch-OutputFiles",

    [ValidateRange(0.0, 1.0)]
    [double]$FuzzyThreshold = 0.80,

    [ValidateRange(1, 100000)]
    [int]$MaxFuzzyHitsPerPhrase = 25
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Compiled helpers (CSV/XLSX readers + matcher)
$helperSource = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace WhaPhraseSearch
{
    public sealed class PhraseHit
    {
        public int PhraseIndex;
        public string MatchType;
        public double Score;
        public string MatchedText;
        public string FilePath;
        public string Sheet;
        public int Row;
        public string Column;
        public string ColumnHeader;
        public string CellValue;
    }

    public static class TextUtil
    {
        public static string Normalize(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            string s = value;
            if (s.IndexOf('%') >= 0)
            {
                try { s = Uri.UnescapeDataString(s); } catch { }
            }
            StringBuilder sb = new StringBuilder(s.Length);
            bool pendingSpace = false;
            foreach (char ch in s)
            {
                if (char.IsWhiteSpace(ch) || ch == '\u00A0')
                {
                    pendingSpace = sb.Length > 0;
                    continue;
                }
                if (pendingSpace) { sb.Append(' '); pendingSpace = false; }
                sb.Append(ch == '\\' ? '/' : char.ToLowerInvariant(ch));
            }
            return sb.ToString();
        }

        public static HashSet<long> Trigrams(string s)
        {
            HashSet<long> set = new HashSet<long>();
            for (int i = 0; i + 2 < s.Length; i++)
            {
                set.Add(((long)s[i] << 32) | ((long)s[i + 1] << 16) | (long)s[i + 2]);
            }
            return set;
        }

        // Sellers' algorithm: smallest edit distance between pattern and any substring of text.
        public static int BestSubstring(string pattern, string text, out int start, out int end)
        {
            int bestEnd;
            int best = SellersForward(pattern, text, text.Length, out bestEnd);
            char[] rp = pattern.ToCharArray(); Array.Reverse(rp);
            char[] rt = text.Substring(0, bestEnd).ToCharArray(); Array.Reverse(rt);
            int revEnd;
            SellersForward(new string(rp), new string(rt), rt.Length, out revEnd);
            end = bestEnd;
            start = bestEnd - revEnd;
            if (start < 0) start = 0;
            return best;
        }

        private static int SellersForward(string p, string t, int n, out int bestEnd)
        {
            int m = p.Length;
            int[] col = new int[m + 1];
            for (int i = 0; i <= m; i++) col[i] = i;
            int best = m;
            bestEnd = 0;
            for (int j = 0; j < n; j++)
            {
                int prevDiag = col[0];
                col[0] = 0;
                char tc = t[j];
                for (int i = 1; i <= m; i++)
                {
                    int above = col[i];
                    int v = prevDiag + (p[i - 1] == tc ? 0 : 1);
                    if (above + 1 < v) v = above + 1;
                    if (col[i - 1] + 1 < v) v = col[i - 1] + 1;
                    col[i] = v;
                    prevDiag = above;
                }
                if (col[m] < best) { best = col[m]; bestEnd = j + 1; if (best == 0) break; }
            }
            return best;
        }

        public static Encoding DetectEncoding(string path)
        {
            byte[] buf = new byte[65536];
            int read;
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                read = fs.Read(buf, 0, buf.Length);
            }
            if (read >= 2 && ((buf[0] == 0xFF && buf[1] == 0xFE) || (buf[0] == 0xFE && buf[1] == 0xFF))) return Encoding.Unicode;
            if (read >= 3 && buf[0] == 0xEF && buf[1] == 0xBB && buf[2] == 0xBF) return new UTF8Encoding(false);
            // Trim a possibly split multi-byte char at the end of the sample.
            int len = read;
            int back = 0;
            while (len > 0 && back < 3 && (buf[len - 1] & 0xC0) == 0x80) { len--; back++; }
            if (len > 0 && buf[len - 1] >= 0xC0) len--;
            try
            {
                new UTF8Encoding(false, true).GetString(buf, 0, len);
                return new UTF8Encoding(false);
            }
            catch (DecoderFallbackException)
            {
                try { return Encoding.GetEncoding(1252); }
                catch { return Encoding.GetEncoding(28591); }
            }
        }
    }

    public static class CsvReader
    {
        public static IEnumerable<List<string>> ReadRecords(TextReader reader)
        {
            List<string> record = new List<string>();
            StringBuilder field = new StringBuilder();
            bool inQuotes = false;
            bool any = false;
            int c;
            while ((c = reader.Read()) != -1)
            {
                char ch = (char)c;
                any = true;
                if (inQuotes)
                {
                    if (ch == '"')
                    {
                        if (reader.Peek() == '"') { reader.Read(); field.Append('"'); }
                        else inQuotes = false;
                    }
                    else field.Append(ch);
                }
                else if (ch == '"') inQuotes = true;
                else if (ch == ',') { record.Add(field.ToString()); field.Length = 0; }
                else if (ch == '\r' || ch == '\n')
                {
                    if (ch == '\r' && reader.Peek() == '\n') reader.Read();
                    record.Add(field.ToString()); field.Length = 0;
                    yield return record;
                    record = new List<string>();
                    any = false;
                }
                else field.Append(ch);
            }
            if (any)
            {
                record.Add(field.ToString());
                yield return record;
            }
        }

        public static List<List<string>> ReadFile(string path)
        {
            List<List<string>> rows = new List<List<string>>();
            Encoding enc = TextUtil.DetectEncoding(path);
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (StreamReader sr = new StreamReader(fs, enc, true))
            {
                foreach (List<string> r in ReadRecords(sr)) rows.Add(r);
            }
            return rows;
        }
    }

    public sealed class PhraseMatcher
    {
        public readonly string[] Phrases;
        public readonly int[] ExactCounts;
        public readonly int[] FuzzyCounts;
        public readonly List<PhraseHit> ExactHits = new List<PhraseHit>();
        public double Threshold;
        public int MaxFuzzyHitsPerPhrase;
        public int MaxCellValueLength = 2000;
        public bool FuzzyPass;
        public long CellsProcessed;

        private readonly string[] _norm;
        private readonly HashSet<long>[] _grams;
        private readonly List<PhraseHit>[] _fuzzy;
        private readonly bool[] _fuzzyActive;

        public PhraseMatcher(string[] phrases, double threshold, int maxFuzzyHitsPerPhrase)
        {
            Phrases = phrases;
            Threshold = threshold;
            MaxFuzzyHitsPerPhrase = maxFuzzyHitsPerPhrase;
            int n = phrases.Length;
            ExactCounts = new int[n];
            FuzzyCounts = new int[n];
            _norm = new string[n];
            _grams = new HashSet<long>[n];
            _fuzzy = new List<PhraseHit>[n];
            _fuzzyActive = new bool[n];
            for (int i = 0; i < n; i++)
            {
                _norm[i] = TextUtil.Normalize(phrases[i]);
                _grams[i] = TextUtil.Trigrams(_norm[i]);
                _fuzzy[i] = new List<PhraseHit>();
            }
        }

        public int BeginFuzzyPass()
        {
            FuzzyPass = true;
            int active = 0;
            for (int i = 0; i < Phrases.Length; i++)
            {
                _fuzzyActive[i] = ExactCounts[i] == 0 && _norm[i].Length > 0;
                if (_fuzzyActive[i]) active++;
            }
            return active;
        }

        public void ProcessCell(string file, string sheet, int row, string column, string header, string value)
        {
            if (string.IsNullOrEmpty(value)) return;
            string norm = TextUtil.Normalize(value);
            if (norm.Length == 0) return;
            CellsProcessed++;

            if (!FuzzyPass)
            {
                for (int i = 0; i < _norm.Length; i++)
                {
                    if (_norm[i].Length == 0) continue;
                    if (norm.IndexOf(_norm[i], StringComparison.Ordinal) >= 0)
                    {
                        ExactCounts[i]++;
                        ExactHits.Add(MakeHit(i, "Exact", 1.0, Phrases[i], file, sheet, row, column, header, value));
                    }
                }
                return;
            }

            HashSet<long> cellGrams = null;
            for (int i = 0; i < _norm.Length; i++)
            {
                if (!_fuzzyActive[i]) continue;
                string p = _norm[i];
                int m = p.Length;
                int k = (int)Math.Floor((1.0 - Threshold) * m + 1e-9);
                if (norm.Length < m - k) continue;

                // q-gram filter: each edit can destroy at most 3 of the phrase's trigrams.
                int needed = _grams[i].Count - 3 * k;
                if (needed > 0)
                {
                    if (cellGrams == null) cellGrams = TextUtil.Trigrams(norm);
                    int shared = 0;
                    foreach (long g in _grams[i]) { if (cellGrams.Contains(g)) shared++; }
                    if (shared < needed) continue;
                }

                int start, end;
                int dist = TextUtil.BestSubstring(p, norm, out start, out end);
                if (dist > k) continue;
                double score = Math.Round(1.0 - (double)dist / m, 4);
                FuzzyCounts[i]++;
                List<PhraseHit> list = _fuzzy[i];
                list.Add(MakeHit(i, "Fuzzy", score, norm.Substring(start, end - start), file, sheet, row, column, header, value));
                if (list.Count > MaxFuzzyHitsPerPhrase * 2 + 50) Trim(list);
            }
        }

        public List<PhraseHit> GetFuzzyHits(int phraseIndex)
        {
            List<PhraseHit> list = _fuzzy[phraseIndex];
            Trim(list);
            return list;
        }

        private void Trim(List<PhraseHit> list)
        {
            list.Sort(delegate (PhraseHit a, PhraseHit b) { return b.Score.CompareTo(a.Score); });
            if (list.Count > MaxFuzzyHitsPerPhrase) list.RemoveRange(MaxFuzzyHitsPerPhrase, list.Count - MaxFuzzyHitsPerPhrase);
        }

        private PhraseHit MakeHit(int i, string type, double score, string matched, string file, string sheet, int row, string column, string header, string value)
        {
            PhraseHit h = new PhraseHit();
            h.PhraseIndex = i;
            h.MatchType = type;
            h.Score = score;
            h.MatchedText = matched;
            h.FilePath = file;
            h.Sheet = sheet;
            h.Row = row;
            h.Column = column;
            h.ColumnHeader = header;
            h.CellValue = value.Length > MaxCellValueLength ? value.Substring(0, MaxCellValueLength) + "..." : value;
            return h;
        }
    }

    public static class FileScanner
    {
        public static string ColumnLetter(int col)
        {
            StringBuilder sb = new StringBuilder();
            while (col > 0) { int r = (col - 1) % 26; sb.Insert(0, (char)('A' + r)); col = (col - 1) / 26; }
            return sb.ToString();
        }

        private static int ColumnFromRef(string cellRef)
        {
            int col = 0;
            foreach (char ch in cellRef)
            {
                char u = char.ToUpperInvariant(ch);
                if (u < 'A' || u > 'Z') break;
                col = col * 26 + (u - 'A' + 1);
            }
            return col;
        }

        public static void ScanCsv(string path, PhraseMatcher matcher)
        {
            Encoding enc = TextUtil.DetectEncoding(path);
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (StreamReader sr = new StreamReader(fs, enc, true))
            {
                List<string> headers = null;
                int rowNum = 0;
                foreach (List<string> rec in CsvReader.ReadRecords(sr))
                {
                    rowNum++;
                    if (headers == null) headers = rec;
                    for (int c = 0; c < rec.Count; c++)
                    {
                        string header = c < headers.Count ? headers[c] : string.Empty;
                        matcher.ProcessCell(path, string.Empty, rowNum, ColumnLetter(c + 1), header, rec[c]);
                    }
                }
            }
        }

        private static XmlReaderSettings XmlSettings()
        {
            XmlReaderSettings s = new XmlReaderSettings();
            s.DtdProcessing = DtdProcessing.Prohibit;
            s.XmlResolver = null;
            s.IgnoreComments = true;
            s.IgnoreProcessingInstructions = true;
            return s;
        }

        private static ZipArchiveEntry FindEntry(ZipArchive zip, string name)
        {
            ZipArchiveEntry e = zip.GetEntry(name);
            if (e != null) return e;
            foreach (ZipArchiveEntry x in zip.Entries)
            {
                if (string.Equals(x.FullName, name, StringComparison.OrdinalIgnoreCase)) return x;
            }
            return null;
        }

        private static List<string> ReadSharedStrings(ZipArchive zip)
        {
            List<string> list = new List<string>();
            ZipArchiveEntry e = FindEntry(zip, "xl/sharedStrings.xml");
            if (e == null) return list;
            using (Stream s = e.Open())
            using (XmlReader r = XmlReader.Create(s, XmlSettings()))
            {
                StringBuilder sb = new StringBuilder();
                bool inT = false;
                int rphDepth = 0;
                while (r.Read())
                {
                    if (r.NodeType == XmlNodeType.Element)
                    {
                        if (r.LocalName == "si") { sb.Length = 0; if (r.IsEmptyElement) list.Add(string.Empty); }
                        else if (r.LocalName == "rPh" && !r.IsEmptyElement) rphDepth++;
                        else if (r.LocalName == "t" && !r.IsEmptyElement) inT = true;
                    }
                    else if (r.NodeType == XmlNodeType.EndElement)
                    {
                        if (r.LocalName == "si") list.Add(sb.ToString());
                        else if (r.LocalName == "rPh") rphDepth--;
                        else if (r.LocalName == "t") inT = false;
                    }
                    else if (inT && rphDepth == 0 && (r.NodeType == XmlNodeType.Text || r.NodeType == XmlNodeType.CDATA || r.NodeType == XmlNodeType.Whitespace || r.NodeType == XmlNodeType.SignificantWhitespace))
                    {
                        sb.Append(r.Value);
                    }
                }
            }
            return list;
        }

        private static List<KeyValuePair<string, string>> ReadSheetList(ZipArchive zip)
        {
            List<KeyValuePair<string, string>> sheets = new List<KeyValuePair<string, string>>();
            Dictionary<string, string> rels = new Dictionary<string, string>();
            ZipArchiveEntry relEntry = FindEntry(zip, "xl/_rels/workbook.xml.rels");
            if (relEntry != null)
            {
                using (Stream s = relEntry.Open())
                using (XmlReader r = XmlReader.Create(s, XmlSettings()))
                {
                    while (r.Read())
                    {
                        if (r.NodeType == XmlNodeType.Element && r.LocalName == "Relationship")
                        {
                            string id = r.GetAttribute("Id");
                            string target = r.GetAttribute("Target");
                            if (id == null || target == null) continue;
                            target = target.Replace('\\', '/');
                            target = target.StartsWith("/") ? target.TrimStart('/') : "xl/" + target;
                            rels[id] = target;
                        }
                    }
                }
            }
            ZipArchiveEntry wb = FindEntry(zip, "xl/workbook.xml");
            if (wb != null)
            {
                using (Stream s = wb.Open())
                using (XmlReader r = XmlReader.Create(s, XmlSettings()))
                {
                    while (r.Read())
                    {
                        if (r.NodeType != XmlNodeType.Element || r.LocalName != "sheet") continue;
                        string name = r.GetAttribute("name");
                        string rid = null;
                        if (r.MoveToFirstAttribute())
                        {
                            do { if (r.LocalName == "id" && r.Prefix.Length > 0) rid = r.Value; } while (r.MoveToNextAttribute());
                            r.MoveToElement();
                        }
                        string target;
                        if (rid != null && rels.TryGetValue(rid, out target)) sheets.Add(new KeyValuePair<string, string>(name ?? target, target));
                    }
                }
            }
            if (sheets.Count == 0)
            {
                foreach (ZipArchiveEntry x in zip.Entries)
                {
                    if (x.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase) && x.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) && x.FullName.IndexOf('/', 14) < 0)
                        sheets.Add(new KeyValuePair<string, string>(Path.GetFileNameWithoutExtension(x.FullName), x.FullName));
                }
            }
            return sheets;
        }

        public static void ScanXlsx(string path, PhraseMatcher matcher)
        {
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                List<string> shared = ReadSharedStrings(zip);
                foreach (KeyValuePair<string, string> sheet in ReadSheetList(zip))
                {
                    ZipArchiveEntry e = FindEntry(zip, sheet.Value);
                    if (e == null) continue;
                    Dictionary<int, string> headers = new Dictionary<int, string>();
                    using (Stream s = e.Open())
                    using (XmlReader r = XmlReader.Create(s, XmlSettings()))
                    {
                        int rowNum = 0, colIdx = 0;
                        string cellType = null;
                        bool inCell = false, inV = false, inT = false;
                        StringBuilder vBuf = new StringBuilder(), tBuf = new StringBuilder();
                        while (r.Read())
                        {
                            if (r.NodeType == XmlNodeType.Element)
                            {
                                string ln = r.LocalName;
                                if (ln == "row")
                                {
                                    int rn;
                                    string ra = r.GetAttribute("r");
                                    rowNum = (ra != null && int.TryParse(ra, out rn)) ? rn : rowNum + 1;
                                    colIdx = 0;
                                }
                                else if (ln == "c")
                                {
                                    string cr = r.GetAttribute("r");
                                    int parsed = cr != null ? ColumnFromRef(cr) : 0;
                                    colIdx = parsed > 0 ? parsed : colIdx + 1;
                                    cellType = r.GetAttribute("t");
                                    vBuf.Length = 0; tBuf.Length = 0;
                                    inCell = !r.IsEmptyElement;
                                }
                                else if (inCell && ln == "v" && !r.IsEmptyElement) inV = true;
                                else if (inCell && ln == "t" && !r.IsEmptyElement) inT = true;
                            }
                            else if (r.NodeType == XmlNodeType.EndElement)
                            {
                                string ln = r.LocalName;
                                if (ln == "v") inV = false;
                                else if (ln == "t") inT = false;
                                else if (ln == "c" && inCell)
                                {
                                    inCell = false;
                                    string val;
                                    if (cellType == "s")
                                    {
                                        int idx;
                                        val = (int.TryParse(vBuf.ToString(), out idx) && idx >= 0 && idx < shared.Count) ? shared[idx] : vBuf.ToString();
                                    }
                                    else if (cellType == "inlineStr") val = tBuf.ToString();
                                    else val = vBuf.ToString();
                                    if (rowNum == 1) headers[colIdx] = val;
                                    string header;
                                    if (!headers.TryGetValue(colIdx, out header)) header = string.Empty;
                                    matcher.ProcessCell(path, sheet.Key, rowNum, ColumnLetter(colIdx), header, val);
                                }
                            }
                            else if (r.NodeType == XmlNodeType.Text || r.NodeType == XmlNodeType.CDATA || r.NodeType == XmlNodeType.Whitespace || r.NodeType == XmlNodeType.SignificantWhitespace)
                            {
                                if (inV) vBuf.Append(r.Value);
                                else if (inT) tBuf.Append(r.Value);
                            }
                        }
                    }
                }
            }
        }
    }
}
'@

if (-not ('WhaPhraseSearch.PhraseMatcher' -as [type])) {
    if ($PSVersionTable.PSEdition -eq 'Desktop') {
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.Xml
        $refs = @(
            [System.IO.Compression.ZipArchive].Assembly.Location,
            [System.Xml.XmlReader].Assembly.Location
        )
        Add-Type -TypeDefinition $helperSource -Language CSharp -ReferencedAssemblies $refs
    }
    else {
        Add-Type -TypeDefinition $helperSource -Language CSharp
    }
}
#endregion

#region Validate inputs / load phrases
if (-not (Test-Path -LiteralPath $SearchRoot -PathType Container)) {
    throw "SearchRoot folder not found: $SearchRoot"
}
if (-not (Test-Path -LiteralPath $PhrasesCsvPath -PathType Leaf)) {
    throw "Search phrases CSV not found: $PhrasesCsvPath"
}
$SearchRoot = (Resolve-Path -LiteralPath $SearchRoot).ProviderPath.TrimEnd('\')
$PhrasesCsvPath = (Resolve-Path -LiteralPath $PhrasesCsvPath).ProviderPath

$seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$phraseList = New-Object 'System.Collections.Generic.List[string]'
foreach ($record in [WhaPhraseSearch.CsvReader]::ReadFile($PhrasesCsvPath)) {
    if ($record.Count -eq 0) { continue }
    $phrase = ($record[0] -replace '\u00A0', ' ').Trim()
    if ($phrase -and $seen.Add($phrase)) { $phraseList.Add($phrase) }
}
if ($phraseList.Count -eq 0) { throw "No search phrases found in $PhrasesCsvPath" }
Write-Host "Loaded $($phraseList.Count) unique search phrase(s) from $PhrasesCsvPath"

if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}
$OutputFolder = (Resolve-Path -LiteralPath $OutputFolder).ProviderPath.TrimEnd('\')
#endregion

#region Find files
$extensions = @('.csv', '.xlsx', '.xlsm', '.xls', '.xlsb')
$files = @(Get-ChildItem -LiteralPath $SearchRoot -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object {
        ($extensions -contains $_.Extension.ToLowerInvariant()) -and
        (-not $_.Name.StartsWith('~$')) -and
        (-not $_.FullName.StartsWith($OutputFolder + '\', [StringComparison]::OrdinalIgnoreCase))
    } | Sort-Object FullName)

if ($files.Count -eq 0) {
    Write-Warning "No CSV/Excel files found under $SearchRoot"
    return
}
Write-Host "Found $($files.Count) CSV/Excel file(s) under $SearchRoot"
#endregion

#region Scan
$matcher = [WhaPhraseSearch.PhraseMatcher]::new([string[]]$phraseList.ToArray(), $FuzzyThreshold, $MaxFuzzyHitsPerPhrase)
$errors = New-Object 'System.Collections.Generic.List[object]'
$failedFiles = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$script:excelApp = $null

function Invoke-ExcelComScan {
    param([string]$Path, [WhaPhraseSearch.PhraseMatcher]$Matcher)

    if (-not $script:excelApp) {
        try {
            $script:excelApp = New-Object -ComObject Excel.Application
            $script:excelApp.Visible = $false
            $script:excelApp.DisplayAlerts = $false
        }
        catch {
            throw "Microsoft Excel is required to read .xls/.xlsb files but could not be started: $($_.Exception.Message)"
        }
    }

    $wb = $script:excelApp.Workbooks.Open($Path, 0, $true)
    try {
        foreach ($ws in $wb.Worksheets) {
            $used = $ws.UsedRange
            $values = $used.Value2
            $startRow = $used.Row
            $startCol = $used.Column
            if ($null -eq $values) { continue }
            if ($values -isnot [array]) {
                $Matcher.ProcessCell($Path, $ws.Name, $startRow, [WhaPhraseSearch.FileScanner]::ColumnLetter($startCol), '', [string]$values)
                continue
            }
            $rows = $values.GetLength(0)
            $cols = $values.GetLength(1)
            $headers = @{}
            if ($startRow -eq 1) {
                for ($c = 1; $c -le $cols; $c++) { $headers[$c] = [string]$values[1, $c] }
            }
            for ($r = 1; $r -le $rows; $r++) {
                for ($c = 1; $c -le $cols; $c++) {
                    $v = $values[$r, $c]
                    if ($null -eq $v) { continue }
                    $hdr = if ($headers.ContainsKey($c)) { $headers[$c] } else { '' }
                    $Matcher.ProcessCell($Path, $ws.Name, $startRow + $r - 1, [WhaPhraseSearch.FileScanner]::ColumnLetter($startCol + $c - 1), $hdr, [string]$v)
                }
            }
        }
    }
    finally {
        $wb.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    }
}

function Invoke-FileScan {
    param([System.IO.FileInfo]$File)

    switch ($File.Extension.ToLowerInvariant()) {
        '.csv' { [WhaPhraseSearch.FileScanner]::ScanCsv($File.FullName, $matcher) }
        { $_ -in '.xlsx', '.xlsm' } {
            try { [WhaPhraseSearch.FileScanner]::ScanXlsx($File.FullName, $matcher) }
            catch [System.IO.InvalidDataException] {
                # Mis-named file (e.g. an old binary workbook saved as .xlsx) - let Excel try.
                Invoke-ExcelComScan -Path $File.FullName -Matcher $matcher
            }
        }
        default { Invoke-ExcelComScan -Path $File.FullName -Matcher $matcher }
    }
}

function Invoke-ScanPass {
    param([string]$PassName, [int]$ProgressId)

    $i = 0
    foreach ($file in $files) {
        $i++
        if ($failedFiles.Contains($file.FullName)) { continue }
        Write-Progress -Id $ProgressId -Activity "$PassName pass" -Status "[$i/$($files.Count)] $($file.Name)" -PercentComplete (($i - 1) * 100 / $files.Count)
        try {
            Invoke-FileScan -File $file
        }
        catch {
            [void]$failedFiles.Add($file.FullName)
            $errors.Add([pscustomobject]@{
                FilePath = $file.FullName
                Pass     = $PassName
                Error    = $_.Exception.Message
            })
            Write-Warning "Could not read '$($file.FullName)': $($_.Exception.Message)"
        }
    }
    Write-Progress -Id $ProgressId -Activity "$PassName pass" -Completed
}

$sw = [System.Diagnostics.Stopwatch]::StartNew()
try {
    Write-Host "Pass 1: exact matching..."
    Invoke-ScanPass -PassName 'Exact' -ProgressId 1
    Write-Host ("  Scanned {0:N0} non-empty cells in {1:N1}s" -f $matcher.CellsProcessed, $sw.Elapsed.TotalSeconds)

    $fuzzyCount = $matcher.BeginFuzzyPass()
    if ($fuzzyCount -gt 0) {
        Write-Host "Pass 2: fuzzy matching for $fuzzyCount phrase(s) with no exact match (threshold $FuzzyThreshold)..."
        Invoke-ScanPass -PassName 'Fuzzy' -ProgressId 1
    }
    else {
        Write-Host "All phrases had exact matches; skipping fuzzy pass."
    }
}
finally {
    if ($script:excelApp) {
        $script:excelApp.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:excelApp)
        $script:excelApp = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
#endregion

#region Output
# One row per phrase / file / sheet / row (several matching cells in the same row collapse to one).
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$resultsPath = Join-Path $OutputFolder "PhraseSearch-Results-$stamp.csv"
$errorsPath = Join-Path $OutputFolder "PhraseSearch-Errors-$stamp.csv"

$allHits = New-Object 'System.Collections.Generic.List[object]'
foreach ($hit in $matcher.ExactHits) { $allHits.Add($hit) }
for ($p = 0; $p -lt $phraseList.Count; $p++) {
    if ($matcher.ExactCounts[$p] -eq 0) {
        foreach ($hit in $matcher.GetFuzzyHits($p)) { $allHits.Add($hit) }
    }
}

$rowKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$resultRows = New-Object 'System.Collections.Generic.List[object]'
foreach ($hit in ($allHits | Sort-Object PhraseIndex, @{ Expression = 'Score'; Descending = $true }, FilePath, Sheet, Row)) {
    $key = '{0}|{1}|{2}|{3}' -f $hit.PhraseIndex, $hit.FilePath, $hit.Sheet, $hit.Row
    if (-not $rowKeys.Add($key)) { continue }
    $resultRows.Add([pscustomobject][ordered]@{
        SearchPhrase = $phraseList[$hit.PhraseIndex]
        MatchType    = $hit.MatchType
        FilePath     = $hit.FilePath
        Sheet        = $hit.Sheet
        Row          = $hit.Row
    })
}

if ($resultRows.Count -gt 0) {
    $resultRows | Export-Csv -LiteralPath $resultsPath -NoTypeInformation -Encoding UTF8
}
if ($errors.Count -gt 0) {
    $errors | Export-Csv -LiteralPath $errorsPath -NoTypeInformation -Encoding UTF8
}

$sw.Stop()
$exactN = @($matcher.ExactCounts | Where-Object { $_ -gt 0 }).Count
$fuzzyN = 0
for ($p = 0; $p -lt $phraseList.Count; $p++) {
    if ($matcher.ExactCounts[$p] -eq 0 -and $matcher.FuzzyCounts[$p] -gt 0) { $fuzzyN++ }
}
Write-Host ""
Write-Host ("Done in {0:N1}s. Phrases: {1} exact, {2} fuzzy-only, {3} not found." -f $sw.Elapsed.TotalSeconds, $exactN, $fuzzyN, ($phraseList.Count - $exactN - $fuzzyN))
if ($resultRows.Count -gt 0) { Write-Host "Results: $resultsPath" } else { Write-Host "No matches found; results file not written." }
if ($errors.Count -gt 0) { Write-Host "Errors:  $errorsPath ($($errors.Count) file(s) could not be read)" -ForegroundColor Yellow }
#endregion
