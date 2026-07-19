using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ClosedXML.Excel;

namespace QiLiaoReply.Services;

/// <summary>
/// 单个单元格。忠实对应 openpyxl 的 cell.value 类型：
///   0=空  1=字符串/布尔文本  2=数字  3=日期  4=布尔
/// </summary>
public sealed class Cell
{
    public int Type;
    public string? Text;
    public double? Number;
    public DateTime? Date;
    public string? Format;

    public bool IsEmpty => Type == 0 || (Text == null && !Number.HasValue && !Date.HasValue);

    /// <summary>用于「整表搜索采购员姓名」及判定列解析的字符串表示（等价于 openpyxl 的 str(cell.value)）。</summary>
    public string Display =>
        Date.HasValue ? Date.Value.ToString("yyyy-MM-dd HH:mm:ss")
        : Number.HasValue ? Number.Value.ToString(CultureInfo.InvariantCulture)
        : (Text ?? "");
}

/// <summary>
/// 忠实读取第一个工作表（由 C 核心 qiliao_core.dll 加速解析）。
/// Rows[0] 为表头行；其余为数据行。
/// </summary>
public sealed class CellGrid
{
    public List<List<Cell>> Rows { get; } = new();
    public int RowCount => Rows.Count;
    public int ColumnCount => Rows.Count > 0 ? Rows[0].Count : 0;

    public Cell this[int r, int c]
    {
        get
        {
            if (r < 0 || r >= Rows.Count) return _empty;
            var row = Rows[r];
            if (c < 0 || c >= row.Count) return _empty;
            return row[c];
        }
    }

    private static readonly Cell _empty = new();

    public IReadOnlyList<Cell> HeaderRow => Rows.Count > 0 ? Rows[0] : Array.Empty<Cell>();
    public IEnumerable<List<Cell>> DataRows
    {
        get { for (int i = 1; i < Rows.Count; i++) yield return Rows[i]; }
    }

    /// <summary>
    /// 读取 xlsx 第一个工作表。
    /// 使用 ClosedXML（等价于 v5.0 的 openpyxl data_only 读取）以保证对各种 xlsx
    /// （含命名空间前缀、共享字符串、公式缓存值等）的可靠解析。
    /// 注：native/qiliao_core.dll 为原计划的高性能 C 解析内核，但其自研 XML/DEFLATE
    /// 解析器对带命名空间前缀的真实 xlsx 不兼容，暂以 ClosedXML 读取替代；一旦用
    /// 正规工具链重新编译 qiliao_core.c 并修复解析器，可在此处切换回 C 内核。
    /// </summary>
    public static CellGrid? Read(string path, out string? error)
    {
        error = null;
        try
        {
            using var wb = new XLWorkbook(path);
            var ws = wb.Worksheet(1); // 第一个工作表（对应 v5.0 的 wb.active / 首表）
            int lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
            int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
            var grid = new CellGrid();
            for (int r = 1; r <= lastRow; r++)
            {
                var row = new List<Cell>(lastCol);
                for (int c = 1; c <= lastCol; c++)
                    row.Add(MakeCellFromXL(ws.Cell(r, c)));
                grid.Rows.Add(row);
            }
            return grid;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return null;
        }
    }

    private static Cell MakeCellFromXL(IXLCell cell)
    {
        string? fmt = cell.IsEmpty() ? null : cell.Style.NumberFormat.Format;
        if (cell.IsEmpty()) return new Cell { Type = 0, Format = fmt };
        var v = cell.Value;
        if (v.IsDateTime) return new Cell { Type = 3, Date = v.GetDateTime(), Format = fmt };
        if (v.IsNumber)
        {
            double n = v.GetNumber();
            // 与 C 内核一致：带日期格式的数字视为日期
            if (fmt != null && Qiliao.IsDateFormat(fmt))
                return new Cell { Type = 3, Date = DateTime.FromOADate(n), Format = fmt };
            return new Cell { Type = 2, Number = n, Format = fmt };
        }
        if (v.IsBoolean) return new Cell { Type = 4, Text = v.GetBoolean() ? "TRUE" : "FALSE", Format = fmt };
        if (v.IsText) return new Cell { Type = 1, Text = v.GetText(), Format = fmt };
        return new Cell { Type = 1, Text = cell.GetValue<string>(), Format = fmt };
    }

    private static Cell MakeCell(int type, string? text, string? fmt)
    {
        var cell = new Cell { Type = type, Format = fmt };
        switch (type)
        {
            case 1:
            case 4:
                cell.Text = text;
                break;
            case 2:
                if (text != null && double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
                    cell.Number = d;
                else
                    cell.Text = text;
                break;
            case 3:
                if (text != null &&
                    DateTime.TryParseExact(text, new[] { "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd" },
                        CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                    cell.Date = dt;
                else
                    cell.Text = text;
                break;
            default:
                break;
        }
        return cell;
    }

    #region P/Invoke —— qiliao_core.dll
    // 注意：qiliao_core.dll 当前由 C++ 编译器产出，导出名为 Itanium 修饰名。
    // 若将来用 C 编译器（或加 extern "C"）重新编译，请同步改回干净名 QL_ReadSheet / QL_FreeSheet。
    [DllImport("qiliao_core.dll", CharSet = CharSet.Unicode, EntryPoint = "_Z12QL_ReadSheetPKwPiS1_PPPcPS1_S4_S2_i")]
    private static extern int QL_ReadSheet(
        [MarshalAs(UnmanagedType.LPWStr)] string path,
        out int outRows, out int outCols,
        out IntPtr outValues, out IntPtr outTypes, out IntPtr outFormats,
        byte[] errBuf, int errCap);

    [DllImport("qiliao_core.dll", EntryPoint = "_Z12QL_FreeSheetiiPPcPiS0_")]
    private static extern void QL_FreeSheet(int rows, int cols, IntPtr values, IntPtr types, IntPtr formats);
    #endregion
}

/// <summary>复刻 v5.0 的纯函数工具。</summary>
public static class Qiliao
{
    /// <summary>列字母 -> 0基索引，如 "M"->12, "A"->0。</summary>
    public static int ColLetterToIndex(string letter)
    {
        int index = 0;
        foreach (char ch in letter.ToUpperInvariant())
        {
            if (ch < 'A' || ch > 'Z') throw new ArgumentException($"无效的列字母：{letter}");
            index = index * 26 + (ch - 'A' + 1);
        }
        return index - 1;
    }

    /// <summary>0基索引 -> 列字母，如 12->"M", 0->"A", 26->"AA"。</summary>
    public static string IndexToColLetter(int index)
    {
        if (index < 0) throw new ArgumentOutOfRangeException(nameof(index));
        string s = "";
        int n = index + 1;
        while (n > 0)
        {
            int rem = (n - 1) % 26;
            s = (char)('A' + rem) + s;
            n = (n - 1) / 26;
        }
        return s;
    }

    /// <summary>复刻 find_column_by_keywords：按关键字（子串，忽略大小写）在表头中定位列，返回 0基索引，找不到返回 -1。</summary>
    public static int FindColumnByKeywords(IReadOnlyList<Cell> headers, params string[] keywords)
    {
        foreach (var kw in keywords)
        {
            string k = kw.ToLowerInvariant();
            for (int i = 0; i < headers.Count; i++)
            {
                var h = headers[i].Display;
                if (string.IsNullOrEmpty(h)) continue;
                if (h.ToLowerInvariant().Contains(k)) return i;
            }
        }
        return -1;
    }

    /// <summary>复刻 extract_suppliers_from_t_cell：按 "," 拆分，每项按 "/" 取第 2 段作为供应商名。</summary>
    public static HashSet<string> ExtractSuppliersFromTCell(string? tValue)
    {
        var set = new HashSet<string>();
        if (string.IsNullOrWhiteSpace(tValue) || tValue.Trim() == "None") return set;
        string s = tValue.Trim();
        foreach (var order in s.Split(','))
        {
            var parts = order.Trim().Split('/');
            if (parts.Length >= 2)
            {
                var name = parts[1].Trim();
                if (!string.IsNullOrEmpty(name)) set.Add(name);
            }
        }
        return set;
    }

    /// <summary>复刻 _is_date_format。</summary>
    public static bool IsDateFormat(string? fmt)
    {
        if (string.IsNullOrEmpty(fmt)) return false;
        string f = fmt.ToLowerInvariant();
        if (f == "general") return false;
        if (f.Contains('y') || f.Contains('d') || f.Contains('h')) return true;
        if (f.Contains('m') && (f.Contains('/') || f.Contains('-') || f.Contains(':') || f.Contains('\\'))) return true;
        return false;
    }

    /// <summary>复刻 _coerce_date。返回 (值, 格式) 或 null。</summary>
    public static (object value, string fmt)? CoerceDate(object? val, string? fmt)
    {
        if (val is XLCellValue xv)
        {
            if (xv.IsDateTime)
                return (xv.GetDateTime(), IsDateFormat(fmt) ? fmt! : "yyyy/mm/dd");
            if (xv.IsNumber)
            {
                double num = xv.GetNumber();
                if (num >= 20000 && num <= 80000 && (fmt == null || fmt.Equals("General", StringComparison.OrdinalIgnoreCase)))
                    return (new DateTime(1899, 12, 30).AddDays(num), IsDateFormat(fmt) ? fmt! : "yyyy/mm/dd");
            }
            return null;
        }
        if (val is DateTime dt)
            return (dt, IsDateFormat(fmt) ? fmt! : "yyyy/mm/dd");
        if (val is DateTimeOffset dto)
            return (dto.DateTime, IsDateFormat(fmt) ? fmt! : "yyyy/mm/dd");

        if (val is double or int or float or long or decimal)
        {
            double num = Convert.ToDouble(val);
            if (num >= 20000 && num <= 80000 && (fmt == null || fmt.Equals("General", StringComparison.OrdinalIgnoreCase)))
                return (new DateTime(1899, 12, 30).AddDays(num), IsDateFormat(fmt) ? fmt! : "yyyy/mm/dd");
        }
        return null;
    }

    /// <summary>复刻 _get_unique_filepath：存在则追加 _1/_2…</summary>
    public static string GetUniqueFilePath(string path)
    {
        if (!System.IO.File.Exists(path)) return path;
        var dir = System.IO.Path.GetDirectoryName(path) ?? "";
        var name = System.IO.Path.GetFileNameWithoutExtension(path);
        var ext = System.IO.Path.GetExtension(path);
        int i = 1;
        while (System.IO.File.Exists(System.IO.Path.Combine(dir, $"{name}_{i}{ext}"))) i++;
        return System.IO.Path.Combine(dir, $"{name}_{i}{ext}");
    }
}
