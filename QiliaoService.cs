using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace QiLiaoReply.Services;

#region 结果类型
public sealed class SupplierExport
{
    public string Supplier = "";
    public int RowCount;
    public string? FilePath;
    public string? Buyer; // 该供应商涉及的主要采购员（跨采购员时以 / 连接）
}

public sealed class ExportResult
{
    public bool Success;
    public string? Error;
    public string? PersonName;
    public int PersonCount;
    public int DeficitRows;
    public int TotalRows;
    public List<SupplierExport> Suppliers = new();
    public int UnmatchedCount;
    public string? UnmatchedFilePath;
    public string? OutputDir;
    public string? DateStr;
    public string? JudgeColLetter;
    public List<string> ExtraCols = new();
    public Dictionary<string, string> SupplierOwners = new(); // 供应商 -> 推断的负责人
}

public sealed class BackfillResult
{
    public bool Success;
    public string? Error;
    public int MatchedCount;
    public int ReplyTotal;
    public string? OutputPath;
    public string? OutputDir;
    public int ReplyFilesLoaded;
    public List<string> ReplyErrors = new();
}
#endregion

public static class QiliaoService
{
    private const int T_COL = 19; // T 列（0基）

    /// <summary>用于 UI 列表展示的列头信息。</summary>
    public sealed class ColumnHeaderInfo
    {
        public int Index;       // 0基
        public string Letter = ""; // "A", "M", ...
        public string Name = "";   // 表头文本（可能为空）
        public override string ToString() => string.IsNullOrEmpty(Name) ? Letter : $"{Letter}  {Name}";
    }

    /// <summary>
    /// 仅读取 xlsx 第一个工作表的第一行（表头），返回列字母+名称列表。
    /// 供 UI 在选择文件后自动填充「判定列」下拉框，避免用户手输列字母。
    /// </summary>
    public static List<ColumnHeaderInfo> ReadHeaders(string filePath, out string? error)
    {
        error = null;
        var list = new List<ColumnHeaderInfo>();
        try
        {
            if (!File.Exists(filePath))
            {
                error = $"文件不存在：{filePath}";
                return list;
            }
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".xlsx" && ext != ".xlsm")
            {
                error = $"不支持的文件格式「{ext}」。本工具仅支持 .xlsx / .xlsm。" +
                        $"若是旧版 .xls 或从网页/ERP 导出的 CSV、HTML，请在 Excel 中「另存为」真正的 .xlsx 后再选。";
                return list;
            }
            using var wb = new XLWorkbook(filePath);
            var ws = wb.Worksheet(1);
            int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
            for (int c = 1; c <= lastCol; c++)
            {
                var cell = ws.Cell(1, c);
                string name = "";
                if (!cell.IsEmpty())
                {
                    if (cell.Value.IsText) name = cell.Value.GetText();
                    else if (cell.Value.IsNumber) name = cell.Value.GetNumber().ToString(CultureInfo.InvariantCulture);
                    else if (cell.Value.IsDateTime) name = cell.Value.GetDateTime().ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                    else if (cell.Value.IsBoolean) name = cell.Value.GetBoolean() ? "TRUE" : "FALSE";
                }
                list.Add(new ColumnHeaderInfo { Index = c - 1, Letter = Qiliao.IndexToColLetter(c - 1), Name = name });
            }
        }
        catch (Exception ex)
        {
            error = DiagnoseReadError(ex);
        }
        return list;
    }

    /// <summary>把底层异常翻译成用户可理解的中文提示。</summary>
    private static string DiagnoseReadError(Exception ex)
    {
        string msg = ex.Message;
        string lower = (ex.GetType().Name + " " + msg).ToLowerInvariant();
        if (ex is IOException || lower.Contains("being used") || lower.Contains("shared"))
            return "文件可能被 Excel 或其他程序占用（独占锁定）。请关闭该文件后点「刷新」重试。";
        if (lower.Contains("corrupt") || lower.Contains("central directory") || lower.Contains("zip") || lower.Contains("invalid signature"))
            return "文件已损坏或不是真正的 Excel（很可能是网页/ERP 导出的伪 xlsx 或 CSV）。请在 Excel 中打开并「另存为」.xlsx 后重试。";
        if (lower.Contains("password") || lower.Contains("encrypt") || lower.Contains("保护"))
            return "文件带有打开密码（加密工作簿），本工具无法读取。请去掉密码后重试。";
        if (lower.Contains("empty") || lower.Contains("worksheet"))
            return "工作簿为空或找不到工作表，请确认文件内容。";
        return $"{ex.GetType().Name}：{msg}";
    }


    #region 导出（复刻 process_欠料表）
    public static ExportResult ProcessExport(string filePath, string personName, string judgeColLetter, string? outDir, string? manualCategory = null)
    {
        var result = new ExportResult
        {
            PersonName = personName,
            JudgeColLetter = (judgeColLetter ?? "M").ToUpperInvariant()
        };

        var grid = CellGrid.Read(filePath, out var err);
        if (grid == null) { result.Success = false; result.Error = err ?? "读取失败。"; return result; }
        result.TotalRows = grid.RowCount - 1;

        int judgeColIndex, catColIndex, uCol, zCol, supplierCol;
        try
        {
            judgeColIndex = Qiliao.ColLetterToIndex(judgeColLetter!);
            catColIndex = Qiliao.ColLetterToIndex("W"); // 外协/外购固定读取 W 列（Y=外协，N=外购）
            uCol = Qiliao.ColLetterToIndex("U");
            zCol = Qiliao.ColLetterToIndex("Z");
            int sIdx = Qiliao.FindColumnByKeywords(grid.HeaderRow, "供应商未清采购订单数量明细", "未清采购订单数量明细", "未清采购订单明细");
            supplierCol = sIdx >= 0 ? sIdx : 19; // 回退到 T 列
        }
        catch (Exception ex)
        {
            result.Success = false; result.Error = "列字母解析失败：" + ex.Message; return result;
        }
        if (judgeColIndex < 0 || judgeColIndex >= grid.ColumnCount)
        { result.Success = false; result.Error = "判定列索引超出范围，请检查列字母。"; return result; }
        if (catColIndex < 0 || catColIndex >= grid.ColumnCount)
        { result.Success = false; result.Error = "外协/外购列索引超出范围，请检查列字母。"; return result; }

        var headers = grid.HeaderRow;
        var map = BuildExportMap(headers, judgeColIndex);

        // 分桶：key = 分类(外购/外协/未分类) + 供应商；同一供应商跨采购员合并到一个表
        var buckets = new Dictionary<string, List<RowInfo>>();
        var unmatched = new List<RowInfo>();
        var supplierBuyers = new Dictionary<string, HashSet<string>>();

        foreach (var row in grid.DataRows)
        {
            double jn = GetNumeric(row, judgeColIndex);
            if (jn >= 0) continue; // 仅欠料（判定列 < 0）

            string cat = ClassifyCategory(CellText(row, catColIndex));
            // 手动类别筛选：指定了类别时仅保留该类别；W 为空/未分类的明细也归入所选类别，避免漏数据
            if (manualCategory != null && cat != manualCategory && cat != "未分类")
                continue;

            string buyer = ExtractBuyer(row, uCol, zCol);

            // 采购员筛选（可选，留空 = 全部）：原始 v5.0 为「全表搜索采购员姓名」，
            // 对整行做包含判定，避免仅按 U 列精确比对时漏掉采购员名出现在 Z 等列的欠料行。
            if (!string.IsNullOrWhiteSpace(personName))
            {
                string rowText = string.Join(" ", row.Select(c => c.Display));
                if (!rowText.Contains(personName, StringComparison.Ordinal))
                    continue;
            }

            // U→T→Z 交叉核对：T 列供应商 + 用采购员/T供应商在 Z 列反查出来的「未生效PO」供应商
            var suppliers = ExtractSuppliersCrossChecked(CellText(row, supplierCol), CellText(row, zCol), buyer);
            var info = new RowInfo { Row = row, Buyer = buyer, Category = manualCategory ?? cat, UCol = uCol, ZCol = zCol };
            if (suppliers.Count == 0)
            {
                unmatched.Add(info);
            }
            else
            {
                foreach (var sup in suppliers)
                {
                    // 手动模式仅按供应商分桶：同名供应商（跨采购员 / 跨 PO）自动合并到同一张表
                    string key = manualCategory == null ? (cat + "\u0001" + sup) : sup;
                    if (!buckets.ContainsKey(key)) buckets[key] = new();
                    buckets[key].Add(info);
                    if (!supplierBuyers.ContainsKey(sup)) supplierBuyers[sup] = new();
                    supplierBuyers[sup].Add(buyer);
                }
            }
        }

        if (buckets.Count == 0 && unmatched.Count == 0)
        {
            result.Success = false;
            result.Error = "没有符合「判定列<0（欠料）」条件的明细。";
            return result;
        }

        string dateStr = DateTime.Now.ToString("yyyyMMdd");
        string srcDir = Path.GetDirectoryName(filePath) ?? ".";
        string outputDir = string.IsNullOrWhiteSpace(outDir) ? Path.Combine(srcDir, "欠料表_导出") : outDir;
        Directory.CreateDirectory(outputDir);

        // 全局供应商名称归一：同一公司可能在 T 用短名、Z 用全称（或不同行写法不一），
        // 导致被拆成多张欠料表。若某名称是另一名称的子串，统一到更完整（更长）的那个。
        var canonicalMap = BuildCanonicalSupplierMap(buckets.Keys, manualCategory);
        var normBuckets = new Dictionary<string, List<RowInfo>>();
        var normBuyers = new Dictionary<string, HashSet<string>>();
        foreach (var kv in buckets)
        {
            string catPart = manualCategory == null ? kv.Key.Split('\u0001')[0] : manualCategory!;
            string sup = manualCategory == null ? kv.Key.Split('\u0001')[1] : kv.Key;
            string canon = canonicalMap.TryGetValue(sup, out var c) ? c : sup;
            string newKey = manualCategory == null ? (catPart + "\u0001" + canon) : canon;
            if (!normBuckets.ContainsKey(newKey)) normBuckets[newKey] = new();
            normBuckets[newKey].AddRange(kv.Value);
            if (!normBuyers.ContainsKey(canon)) normBuyers[canon] = new();
            if (supplierBuyers.TryGetValue(sup, out var bs))
                foreach (var b in bs) normBuyers[canon].Add(b);
        }
        buckets = normBuckets;
        supplierBuyers = normBuyers;

        int deficit = 0;
        foreach (var kv in buckets)
        {
            string sup;
            string catForFolder;
            if (manualCategory == null)
            {
                var parts = kv.Key.Split('\u0001');
                catForFolder = parts[0];
                sup = parts[1];
            }
            else
            {
                catForFolder = manualCategory;
                sup = kv.Key;
            }
            // 手动指定类别时直接输出到 outputDir；选「全部」时按 外购/外协/未分类 分目录
            string subDir = (manualCategory == null) ? Path.Combine(outputDir, catForFolder) : outputDir;
            Directory.CreateDirectory(subDir);
            string safe = SafeFileName(sup);
            string outPath = Qiliao.GetUniqueFilePath(Path.Combine(subDir, $"{safe}欠料表{dateStr}.xlsx"));
            CreateWorkbookForRows(kv.Value, map, outPath, sup);
            var buyers = supplierBuyers.TryGetValue(sup, out var bs) ? bs.ToList() : new();
            string buyerDisp = buyers.Count == 0 ? "" : (buyers.Count == 1 ? buyers[0] : string.Join("/", buyers.Distinct()));
            result.Suppliers.Add(new SupplierExport
            {
                Supplier = sup,
                RowCount = kv.Value.Count,
                FilePath = outPath,
                Buyer = buyerDisp
            });
            result.SupplierOwners[sup] = buyerDisp;
            deficit += kv.Value.Count;
        }
        if (unmatched.Count > 0)
        {
            string outPath = Qiliao.GetUniqueFilePath(Path.Combine(outputDir, $"未匹配供应商欠料表{dateStr}.xlsx"));
            CreateWorkbookForRows(unmatched, map, outPath, "");
            result.UnmatchedCount = unmatched.Count;
            result.UnmatchedFilePath = outPath;
            deficit += unmatched.Count;
        }

        result.Success = true;
        result.DeficitRows = deficit;
        result.OutputDir = outputDir;
        result.DateStr = dateStr;
        result.ExtraCols = new List<string> { "交期回复", "供应商回复交货日期", "供应商未清采购订单数量明细", "SRM未生效PO", "采购员", "外协/外购" };
        return result;
    }

    #region 导出辅助类型与方法
    private sealed class RowInfo
    {
        public List<Cell> Row = new();
        public string Buyer = "";
        public string Category = "";
        public int UCol;
        public int ZCol;
    }
    private sealed class OutCol
    {
        public string Header = "";
        public int SrcIdx = -1;
        public bool IsEmpty;
        public Func<RowInfo, string?>? Computed;
    }

    private static double GetNumeric(List<Cell> row, int idx)
    {
        var c = idx >= 0 && idx < row.Count ? row[idx] : null;
        if (c == null) return 0;
        if (c.Type == 2 && c.Number.HasValue) return c.Number.Value;
        if (!c.IsEmpty && double.TryParse(c.Display.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var d)) return d;
        return 0;
    }
    private static string CellText(List<Cell> row, int idx)
    {
        if (idx < 0 || idx >= row.Count) return "";
        var c = row[idx];
        return c.IsEmpty ? "" : c.Display;
    }
    private static string ClassifyCategory(string s)
    {
        s = (s ?? "").Trim();
        if (s == "Y" || s.Equals("外协", StringComparison.OrdinalIgnoreCase)) return "外协";
        if (s == "N" || s.Equals("外购", StringComparison.OrdinalIgnoreCase)) return "外购";
        return "未分类";
    }
    /// <summary>
    /// 采购员：以 U 列（采购员）为准（用户明确：通过 U 列确认本次处理的采购员名字）。
    /// Z 列是「SRM未生效PO」明细，格式为「订单号/供应商/…」，整体文本不是采购员名，
    /// 故绝不能把它当作采购员来源，否则会把 PO 串误当采购员，导致人员筛选与 Z 反查全部失效。
    /// 仅当 U 为空时，才从 Z 中挑一个含字母（中文/拼音）的片段作为回退。
    /// </summary>
    private static string ExtractBuyer(List<Cell> row, int uCol, int zCol)
    {
        // 采购员以 U 列（供应商未清采购订单创建人明细）为准。
        string u = CellText(row, uCol).Trim();
        if (!string.IsNullOrEmpty(u)) return u;
        // U 为空才回退到 Z 列（SRM未生效PO）。其格式为「订单号/供应商/数量/编码/采购员/00010」，
        // 采购员名在「/」第 5 段（index 4）。绝不能取「含字母」的分段——PO 单号（PO2026...）也含字母 P/O，会误当人名。
        string z = CellText(row, zCol).Trim();
        if (!string.IsNullOrEmpty(z))
        {
            foreach (var entry in z.Split(','))
            {
                var segs = entry.Trim().Split('/');
                if (segs.Length >= 5)
                {
                    var b = segs[4].Trim();
                    if (b.Length >= 2) return b;
                }
            }
        }
        return "";
    }
    /// <summary>
    /// 供应商解析（U→T→Z 交叉核对）：
    ///   1) U 列确认采购员；
    ///   2) T 列得到该采购员实际下单的供应商集合（有效单，永远保留）；
    ///   3) Z 列（SRM未生效PO）按「,」拆成单条 PO，每条独立处理：
    ///      - 先用「名称归一」：该条 PO 的供应商（第 2 段）与任一 T 供应商互为包含（相等 / 互为子串）
    ///        → 归入该 T 供应商（同一公司 T 用短名、Z 用全称也不会被拆成两桶）；
    ///      - 否则若该条 PO 含「采购员名」（U 列） → 以 Z 自身供应商名并入（纯未生效、新供应商）；
    ///   4) T 为空时，Z 全部作为该行供应商（含未生效 PO），避免漏。
    /// 这样导出的欠料表是「T 有效单 + Z 中关联/同名的未生效 PO」完整合并，且不会把无关的 Z 供应商混进来。
    /// </summary>
    private static List<string> ExtractSuppliersCrossChecked(string tText, string zText, string buyer)
    {
        var tSuppliers = Qiliao.ExtractSuppliersFromTCell(tText);

        // T 为空：直接以 Z 全部解析结果作为该行供应商（含未生效 PO）
        if (tSuppliers.Count == 0)
            return Qiliao.ExtractSuppliersFromTCell(zText).ToList();

        // T 非空：保留 T 全量（有效单不丢）
        var merged = new HashSet<string>(tSuppliers, StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrEmpty(zText)) return merged.ToList();

        foreach (var raw in zText.Split(','))
        {
            var entry = raw.Trim();
            if (string.IsNullOrEmpty(entry)) continue;

            // 解析该条未生效 PO 的供应商（格式：订单号/供应商/数量/编码/采购员/00010，第 2 段）
            var segs = entry.Split('/');
            if (segs.Length < 2) continue;
            var zSup = segs[1].Trim();
            if (string.IsNullOrEmpty(zSup)) continue;

            // 1) 名称归一：Z 供应商与某个 T 供应商互为包含 → 归入该 T 供应商（避免同一公司分两桶）
            string? canonical = null;
            foreach (var ts in tSuppliers)
            {
                if (string.Equals(zSup, ts, StringComparison.OrdinalIgnoreCase)
                    || zSup.Contains(ts, StringComparison.OrdinalIgnoreCase)
                    || ts.Contains(zSup, StringComparison.OrdinalIgnoreCase))
                {
                    canonical = ts;
                    break;
                }
            }
            if (canonical != null) { merged.Add(canonical); continue; }

            // 2) 否则按采购员名命中 → 以 Z 自身供应商名并入（纯未生效、且为本采购员的 PO）
            if (!string.IsNullOrEmpty(buyer) && entry.Contains(buyer, StringComparison.OrdinalIgnoreCase))
                merged.Add(zSup);
        }
        return merged.ToList();
    }

    /// <summary>
    /// 全局供应商名称归一：若某名称是另一名称的子串（互为包含且不等），统一到更完整（更长）的那个。
    /// 用于把「T 用短名 / Z 用全称 / 各行写法不一」的同一公司合并到一张欠料表，且不会把互不包含的异名公司误并。
    /// </summary>
    private static Dictionary<string, string> BuildCanonicalSupplierMap(IEnumerable<string> keys, string? manualCategory)
    {
        var names = new List<string>();
        foreach (var k in keys)
        {
            string s = manualCategory == null ? k.Split('\u0001')[1] : k;
            if (!names.Contains(s)) names.Add(s);
        }
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var a in names)
        {
            string best = a;
            foreach (var b in names)
            {
                if (a == b) continue;
                if (b.Contains(a, StringComparison.OrdinalIgnoreCase) && b.Length > best.Length)
                    best = b;
            }
            map[a] = best;
        }
        return map;
    }
    private static OutCol[] BuildExportMap(IReadOnlyList<Cell> headers, int judgeColIndex)
    {
        string judgeHeader = (judgeColIndex >= 0 && judgeColIndex < headers.Count)
            ? (headers[judgeColIndex].Text ?? "判定列数值") : "判定列数值";
        return new OutCol[]
        {
            // 原始表头是拆分的：「物料」(col0)=物料编码、「名称」(col1)=物料名称、「型号/规格/品牌」独立成列。
            // 故物料号需能命中「物料」，物料名称需能命中「名称」，否则导出该列为空、回填将因找不到物料编码而整体失败。
            new() { Header = "物料名称", SrcIdx = Qiliao.FindColumnByKeywords(headers, "物料名称", "名称") },
            // 严格按照表中「规格」列(原表 D 列)导出，不另作兜底填充。
            new() { Header = "规格", SrcIdx = Qiliao.FindColumnByKeywords(headers, "规格") },
            new() { Header = "型号", SrcIdx = Qiliao.FindColumnByKeywords(headers, "型号") },
            new() { Header = "品牌", SrcIdx = Qiliao.FindColumnByKeywords(headers, "品牌") },
            new() { Header = "物料号", SrcIdx = Qiliao.FindColumnByKeywords(headers, "物料编码", "物料号", "料号", "物料") },
            new() { Header = judgeHeader, SrcIdx = judgeColIndex },
            new() { Header = "交期回复", IsEmpty = true },
            new() { Header = "供应商回复交货日期", SrcIdx = Qiliao.FindColumnByKeywords(headers, "供应商回复交货日期", "回复交货日期", "交货日期", "到货日期") },
            new() { Header = "供应商未清采购订单数量明细", SrcIdx = Qiliao.FindColumnByKeywords(headers, "供应商未清采购订单数量明细", "未清采购订单数量明细", "未清采购订单明细") },
            new() { Header = "SRM未生效PO", SrcIdx = Qiliao.FindColumnByKeywords(headers, "SRM未生效PO", "未生效PO", "SRM未生效") },
            new() { Header = "采购员", Computed = r => string.IsNullOrEmpty(r.Buyer) ? null : r.Buyer },
            new() { Header = "外协/外购", Computed = r => r.Category },
        };
    }
    #endregion

    /// <summary>
    /// 将一组欠料明细写出为工作簿。supplierName 非空时追加「供应商名称」列，
    /// 该列是回填双键匹配（物料编码+供应商）的关键依据，确保回填能正确归位到对应供应商。
    /// </summary>
    private static void CreateWorkbookForRows(List<RowInfo> rows, OutCol[] map, string outPath, string? supplierName = null)
    {
        // 仅在指定供应商时追加「供应商名称」列（同一张表里该列值固定，便于回填按供应商定位）
        OutCol[] useMap = map;
        if (!string.IsNullOrEmpty(supplierName))
        {
            useMap = map.Append(new OutCol { Header = "供应商名称", IsEmpty = false, Computed = _ => supplierName! }).ToArray();
        }

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        // 表头（蓝底 + 白字粗体 + 居中 + 细边框）
        for (int col = 0; col < useMap.Length; col++)
        {
            var cell = ws.Cell(1, col + 1);
            cell.Value = useMap[col].Header;
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontColor = XLColor.White;
            cell.Style.Fill.BackgroundColor = XLColor.FromArgb(0x44, 0x72, 0xC4);
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        // 数据
        for (int r = 0; r < rows.Count; r++)
        {
            var ri = rows[r];
            var row = ri.Row;
            for (int col = 0; col < useMap.Length; col++)
            {
                var m = useMap[col];
                object? value = null;
                if (m.Computed != null)
                {
                    var cv = m.Computed(ri);
                    if (!string.IsNullOrEmpty(cv)) value = cv;
                }
                else if (!m.IsEmpty && m.SrcIdx >= 0 && m.SrcIdx < row.Count)
                {
                    var src = row[m.SrcIdx];
                    if (src.Type == 2 && src.Number.HasValue) value = src.Number.Value;
                    else if (src.Type == 3 && src.Date.HasValue) value = src.Date.Value;
                    else if (!src.IsEmpty) value = src.Display;
                }
                var dest = ws.Cell(r + 2, col + 1);
                if (value != null)
                {
                    if (value is DateTime dt)
                    {
                        dest.Value = dt;
                        dest.Style.NumberFormat.Format = "yyyy-mm-dd h:mm:ss";
                    }
                    else if (value is double d)
                    {
                        dest.Value = d;
                    }
                    else
                    {
                        dest.Value = value.ToString()!;
                    }
                }
                dest.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                if (col == 5 && value is double dv && dv < 0)
                {
                    dest.Style.Font.FontColor = XLColor.Red;
                    dest.Style.Font.Bold = true;
                }
            }
        }

        // 列宽：max(10, min(len(str(v))+2, 60))
        for (int col = 0; col < useMap.Length; col++)
        {
            double maxW = 10;
            for (int r = 0; r <= rows.Count; r++)
            {
                var cell = ws.Cell(r + 1, col + 1);
                string s = cell.IsEmpty() ? "" : (cell.GetString() ?? "");
                if (!string.IsNullOrEmpty(s))
                {
                    int w = s.Length + 2;
                    maxW = Math.Max(maxW, Math.Min(w, 60));
                }
            }
            ws.Column(col + 1).Width = maxW;
        }

        ws.SheetView.FreezeRows(1);
        if (rows.Count > 0) ws.Range(1, 1, rows.Count + 1, useMap.Length).SetAutoFilter();
        wb.SaveAs(outPath);
    }
    #endregion

    #region 回填（复刻 backfill_reply）
    public static BackfillResult ProcessBackfill(string originalPath, IList<string> replyPaths, string personName, string targetColLetter)
    {
        var result = new BackfillResult();
        // 回复索引：(物料编码, 供应商) -> (值, 格式)
        var replies = new List<(string Code, string Supplier, XLCellValue Value, string Fmt)>();
        int replyTotal = 0;

        foreach (var rp in replyPaths)
        {
            try
            {
                using var rwb = new XLWorkbook(rp);
                var ws = rwb.Worksheet(1);
                var headers = ReadHeaderRow(ws);
                int rColCode = Qiliao.FindColumnByKeywords(headers, "物料编码", "物料号", "料号", "物料", "编码");
                if (rColCode < 0) { result.ReplyErrors.Add($"[{Path.GetFileName(rp)}] 无法找到物料编码列，已跳过"); continue; }
                // 优先取干净的「供应商名称」列；旧版回复单没有该列时回退到含「供应商」的列
                int rColSupplier = Qiliao.FindColumnByKeywords(headers, "供应商名称", "供应商");
                int rColReply = Qiliao.FindColumnByKeywords(headers, "交期回复", "供应商回复", "回复", "交货", "交期", "回复内容");
                if (rColReply < 0) { result.ReplyErrors.Add($"[{Path.GetFileName(rp)}] 无法找到供应商回复列，已跳过"); continue; }

                int maxRow = ws.LastRowUsed()?.RowNumber() ?? 1;
                for (int r = 2; r <= maxRow; r++)
                {
                    string codeVal = NormalizeCode(GetCellString(ws, r, rColCode + 1));
                    if (string.IsNullOrEmpty(codeVal)) continue;
                    var replyCell = ws.Cell(r, rColReply + 1);
                    if (replyCell.IsEmpty()) continue;
                    string supplierVal = rColSupplier >= 0 ? GetCellString(ws, r, rColSupplier + 1) : "";
                    string fmt = replyCell.Style.NumberFormat.Format ?? "";
                    replies.Add((codeVal, supplierVal, replyCell.Value, fmt));
                    replyTotal++;
                }
                result.ReplyFilesLoaded++;
            }
            catch (Exception ex) { result.ReplyErrors.Add($"[{Path.GetFileName(rp)}] 读取失败：{ex.Message}"); }
        }

        if (replies.Count == 0) { result.Success = false; result.Error = "所有回复文件均未提取到任何有效数据。"; return result; }

        using var owb = new XLWorkbook(originalPath);
        var ows = owb.Worksheet(1);
        var oHeaders = ReadHeaderRow(ows);
        int oColCode = Qiliao.FindColumnByKeywords(oHeaders, "物料编码", "物料号", "料号", "物料", "编码");
        if (oColCode < 0) { result.Success = false; result.Error = "原始文件无法找到物料编码列。"; return result; }
        int oColT = Qiliao.FindColumnByKeywords(oHeaders, "供应商未清采购订单数量明细", "未清采购订单数量明细", "未清采购订单明细");
        if (oColT < 0) oColT = 19; // 回退到 T 列
        // Z 列应为「SRM未生效PO」(col25)；注意「供应商未清采购订单创建人明细」是 U 列（采购员名），
        // 不能放在最前，否则会误命中 U 列导致 Z 解析成采购员名。
        int oColZ = Qiliao.FindColumnByKeywords(oHeaders, "SRM未生效PO", "未生效PO", "SRM未生效");
        if (oColZ < 0) oColZ = 25; // 回退到 Z 列
        int oColU = Qiliao.FindColumnByKeywords(oHeaders, "采购员", "采购员姓名", "采购人");
        if (oColU < 0) oColU = 20; // 回退到 U 列

        int targetColIndex;
        try { targetColIndex = Qiliao.ColLetterToIndex(targetColLetter); }
        catch { result.Success = false; result.Error = $"目标列字母无效：{targetColLetter}"; return result; }

        string srcDir = Path.GetDirectoryName(originalPath) ?? ".";
        string baseName = Path.GetFileNameWithoutExtension(originalPath);
        string outName = baseName + "_回填" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
        string outputPath = Qiliao.GetUniqueFilePath(Path.Combine(srcDir, outName)); // 直接输出到源目录，不嵌套子目录

        int matched = 0;
        int oMax = ows.LastRowUsed()?.RowNumber() ?? 1;
        for (int r = 2; r <= oMax; r++)
        {
            string codeVal = NormalizeCode(GetCellString(ows, r, oColCode + 1));
            if (string.IsNullOrEmpty(codeVal)) continue;
            // 原始侧供应商解析：与导出严格一致，U→T→Z 交叉核对
            string tText = GetCellString(ows, r, oColT + 1);
            string zText = GetCellString(ows, r, oColZ + 1);
            string uText = GetCellString(ows, r, oColU + 1);
            var suppliers = ExtractSuppliersCrossChecked(tText, zText, uText);

            var hit = FindReply(replies, codeVal, suppliers);
            if (hit == null) continue;

            var dest = ows.Cell(r, targetColIndex + 1);
            // 回填时尽量还原回复单中的显示效果：
            // 回复单里若是 Excel 日期（DateTime，或落在合理区间内的日期序列号数字，如 46084=2026-02-15），
            // 则转换成真正的日期并套用日期格式，避免出现「回填后变成数字序列号」的情况；
            // 其余（文本、普通数字）原样写回，保持回复原貌。
            var rv = hit.Value.Value;
            string replyFmt = hit.Value.Fmt;
            bool wroteAsDate = false;
            if (rv.IsDateTime)
            {
                dest.Value = rv.GetDateTime();
                wroteAsDate = true;
            }
            else if (rv.IsNumber)
            {
                double num = rv.GetNumber();
                // 合理 Excel 日期序列号范围（约 1927~2089 年）且为整数 → 视为日期
                if (num >= 10000 && num <= 80000 && Math.Abs(num - Math.Round(num)) < 1e-6)
                {
                    try { dest.Value = DateTime.FromOADate(num); wroteAsDate = true; }
                    catch { dest.Value = rv; }
                }
                else
                {
                    dest.Value = rv;
                }
            }
            else
            {
                dest.Value = rv;
            }

            if (wroteAsDate)
            {
                // 回复单元格本身是明确的日期格式则沿用，否则用标准日期格式，确保显示为日期
                if (!string.IsNullOrEmpty(replyFmt) && replyFmt != "General")
                    dest.Style.NumberFormat.Format = replyFmt;
                else
                    dest.Style.NumberFormat.Format = "yyyy-mm-dd";
            }
            else if (!string.IsNullOrEmpty(replyFmt) && replyFmt != "General")
            {
                dest.Style.NumberFormat.Format = replyFmt;
            }
            matched++;
        }

        owb.SaveAs(outputPath);
        result.Success = true;
        result.MatchedCount = matched;
        result.ReplyTotal = replyTotal;
        result.OutputPath = outputPath;
        result.OutputDir = srcDir;
        return result;
    }
    #endregion

    #region 辅助
    private static List<Cell> ReadHeaderRow(IXLWorksheet ws)
    {
        var list = new List<Cell>();
        int maxCol = ws.Row(1).LastCellUsed()?.Address.ColumnNumber ?? 1;
        for (int c = 1; c <= maxCol; c++)
        {
            var cell = ws.Cell(1, c);
            list.Add(new Cell { Type = 1, Text = cell.IsEmpty() ? null : cell.GetString() });
        }
        return list;
    }

    private static string GetCellString(IXLWorksheet ws, int r, int c)
    {
        var cell = ws.Cell(r, c);
        if (cell.IsEmpty()) return "";
        return (cell.GetString() ?? "").Trim();
    }

    private static string SafeFileName(string s)
    {
        foreach (var c in Path.GetInvalidFileNameChars()) s = s.Replace(c, '_');
        s = s.Trim().Trim('.');
        return string.IsNullOrEmpty(s) ? "未命名" : s;
    }

    /// <summary>物料编码归一化：去掉所有空白，便于跨文件精确比对。</summary>
    private static string NormalizeCode(string s)
    {
        if (string.IsNullOrEmpty(s)) return "";
        return new string(s.Where(c => !char.IsWhiteSpace(c)).ToArray());
    }
    private static bool CodeEquals(string a, string b)
    {
        if (a == b) return true;
        if (string.Equals(a, b, StringComparison.OrdinalIgnoreCase)) return true;
        return long.TryParse(a, out var la) && long.TryParse(b, out var lb) && la == lb;
    }
    /// <summary>在回复索引中查找匹配：先按 (编码, 供应商) 精确匹配，再回退到仅按编码匹配。</summary>
    private static (XLCellValue Value, string Fmt)? FindReply(
        List<(string Code, string Supplier, XLCellValue Value, string Fmt)> replies,
        string code, List<string> suppliers)
    {
        foreach (var r in replies)
            if (CodeEquals(r.Code, code) && (suppliers.Contains(r.Supplier) || string.IsNullOrEmpty(r.Supplier)))
                return (r.Value, r.Fmt);
        foreach (var r in replies)
            if (CodeEquals(r.Code, code))
                return (r.Value, r.Fmt);
        return null;
    }
    #endregion
}
