using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using QiLiaoReply.Services;

namespace QiLiaoReply;

public partial class MainWindow : Window
{
    private string? _expFilePath;
    private string? _bfOrigPath;
    private List<string> _bfReplyPaths = new();
    private string? _lastExportDir;
    private string? _lastBackfillPath;

    /// <summary>UI 列表展示用的判定列候选项（包装自 QiliaoService.ColumnHeaderInfo）。</summary>
    public sealed class JudgeColumnItem
    {
        public string Letter { get; set; } = "";
        public string Name { get; set; } = "";
        public string Display => string.IsNullOrEmpty(Name) ? Letter : $"{Letter}  {Name}";
        public override string ToString() => Display;
    }

    public MainWindow()
    {
        InitializeComponent();
        try
        {
            Icon = new System.Windows.Media.Imaging.BitmapImage(
                new Uri("pack://application:,,,/Assets/app.ico", UriKind.Absolute));
        }
        catch { /* 图标缺失不阻断启动 */ }
        InitCatCombo();
        Log("就绪。请选择「导出」或「回填」页签操作。");
    }

    /// <summary>初始化「类别」下拉框：手动选择本次导出为外购还是外协。</summary>
    private void InitCatCombo()
    {
        ExpCatCombo.ItemsSource = new[] { "外购", "外协", "全部（按 W 自动区分）" };
        ExpCatCombo.SelectedIndex = 0; // 默认外购
    }

    private void Log(string s) => LogBox.Items.Insert(0, $"[{DateTime.Now:HH:mm:ss}] {s}");

    #region 导出
    private void ExpBrowse_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFileDialog { Filter = "Excel 文件 (*.xlsx)|*.xlsx", Title = "选择欠料表" };
        if (dlg.ShowDialog() == true)
        {
            _expFilePath = dlg.FileName;
            ExpFileBox.Text = _expFilePath;
            LoadJudgeColumns(_expFilePath);
        }
    }

    /// <summary>读取源文件表头，填充判定列下拉框，默认尝试 M 列。</summary>
    private void LoadJudgeColumns(string path)
    {
        ExpJudgeCombo.ItemsSource = null;
        try
        {
            var headers = QiliaoService.ReadHeaders(path, out var err);
            if (!string.IsNullOrEmpty(err))
            {
                Log("读取表头失败：" + err);
                System.Windows.MessageBox.Show("读取表头失败：" + err, "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            var items = headers.Select(h => new JudgeColumnItem { Letter = h.Letter, Name = h.Name }).ToList();
            ExpJudgeCombo.ItemsSource = items;
            // 优先默认 M（v5.0 常见判定列），否则第一列
            var def = items.FirstOrDefault(i => i.Letter == "M") ?? items.FirstOrDefault();
            if (def != null) ExpJudgeCombo.SelectedItem = def;
            // 优先默认 W（外协/外购列），否则第一列
            var defCat = items.FirstOrDefault(i => i.Letter == "W") ?? items.FirstOrDefault();
            if (defCat != null) ExpCatCombo.SelectedItem = defCat;
            Log($"已加载 {items.Count} 列表头，默认判定列：{(def != null ? def.Display : "（无）")}，外协/外购列：{(defCat != null ? defCat.Display : "（无）")}");
        }
        catch (Exception ex)
        {
            Log("读取表头异常：" + ex.Message);
        }
    }

    private void ExpRefreshCols_Click(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrEmpty(_expFilePath) && File.Exists(_expFilePath))
        {
            LoadJudgeColumns(_expFilePath);
        }
        else
        {
            System.Windows.MessageBox.Show("请先选择欠料表文件。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    /// <summary>读取原欠料表表头，填充回填目标列下拉框，默认尝试 O 列。</summary>
    private void LoadBackfillTargetColumns(string path)
    {
        BfTargetCombo.ItemsSource = null;
        try
        {
            var headers = QiliaoService.ReadHeaders(path, out var err);
            if (!string.IsNullOrEmpty(err))
            {
                Log("读取表头失败：" + err);
                System.Windows.MessageBox.Show("读取表头失败：" + err, "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            var items = headers.Select(h => new JudgeColumnItem { Letter = h.Letter, Name = h.Name }).ToList();
            BfTargetCombo.ItemsSource = items;
            // 默认优先匹配「采购交期回复」列（按列名），其次回退 O 列，再退第一列
            var def = items.FirstOrDefault(i => i.Name != null && i.Name.Contains("采购交期回复"))
                      ?? items.FirstOrDefault(i => i.Name != null && i.Name.Contains("交期回复"))
                      ?? items.FirstOrDefault(i => i.Letter == "O")
                      ?? items.FirstOrDefault();
            if (def != null) BfTargetCombo.SelectedItem = def;
            Log($"[回填] 已加载 {items.Count} 列表头，默认目标列：{(def != null ? def.Display : "（无）")}");
        }
        catch (Exception ex)
        {
            Log("读取表头异常：" + ex.Message);
        }
    }

    private void BfRefreshCols_Click(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrEmpty(_bfOrigPath) && File.Exists(_bfOrigPath))
        {
            LoadBackfillTargetColumns(_bfOrigPath);
        }
        else
        {
            System.Windows.MessageBox.Show("请先选择原欠料表文件。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void ExpRun_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_expFilePath) || !File.Exists(_expFilePath))
        { System.Windows.MessageBox.Show("请先选择欠料表文件。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
        string person = ExpPersonBox.Text.Trim(); // 留空 = 不按采购员筛选（导出全部外购/外协）
        // 判定列：从下拉框获取（若为空则回退到 M）
        string? judge = (ExpJudgeCombo.SelectedItem as JudgeColumnItem)?.Letter;
        if (string.IsNullOrEmpty(judge)) judge = "M";
        // 类别：手动选择（外购 / 外协 / 全部）。非「全部」时按 W 列筛选对应类别。
        string? catSel = ExpCatCombo.SelectedItem as string;
        string? manualCategory = (catSel == "全部（按 W 自动区分）" || string.IsNullOrEmpty(catSel)) ? null : catSel;

        try
        {
            Log($"导出：文件={Path.GetFileName(_expFilePath)} 采购员={(string.IsNullOrEmpty(person) ? "（全部）" : person)} 判定列={judge} 类别={(manualCategory ?? "全部")}");
            var sw = Stopwatch.StartNew();
            var res = QiliaoService.ProcessExport(_expFilePath, person, judge, null, manualCategory);
            sw.Stop();
            if (!res.Success)
            {
                Log("导出失败：" + res.Error);
                ExpOpenBtn.Visibility = Visibility.Collapsed;
                return;
            }
            _lastExportDir = res.OutputDir;
            ExpOpenBtn.Visibility = Visibility.Visible;
            // 导出完成后自动弹出输出文件夹页面
            try
            {
                if (!string.IsNullOrEmpty(_lastExportDir) && Directory.Exists(_lastExportDir))
                    Process.Start(new ProcessStartInfo(_lastExportDir) { UseShellExecute = true });
            }
            catch (Exception ex) { Log("自动打开输出文件夹失败：" + ex.Message); }
            string catLabel = manualCategory ?? "全部";
            string status = $"完成（{catLabel}）：欠料明细 {res.DeficitRows} 条，拆分至 {res.Suppliers.Count} 个供应商";
            if (!string.IsNullOrEmpty(person)) status = $"完成（{catLabel} · 采购员「{person}」）：欠料明细 {res.DeficitRows} 条，拆分至 {res.Suppliers.Count} 个供应商";
            if (res.UnmatchedCount > 0) status += $"（另有 {res.UnmatchedCount} 条未匹配供应商，已单独导出）";
            status += $"。用时 {sw.ElapsedMilliseconds} ms。输出目录：{res.OutputDir}";
            Log(status);
        }
        catch (Exception ex)
        {
            Log("导出异常：" + ex.Message);
        }
    }

    private void ExpOpen_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_lastExportDir) || !Directory.Exists(_lastExportDir))
        {
            System.Windows.MessageBox.Show("请先完成一次导出，生成输出文件夹后再打开。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        try
        {
            Process.Start(new ProcessStartInfo(_lastExportDir) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            Log("打开输出文件夹失败：" + ex.Message);
            System.Windows.MessageBox.Show("无法打开输出文件夹：" + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
    #endregion

    #region 回填
    private void BfOrigBrowse_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFileDialog { Filter = "Excel 文件 (*.xlsx)|*.xlsx", Title = "选择原欠料表" };
        if (dlg.ShowDialog() == true)
        {
            _bfOrigPath = dlg.FileName;
            BfOrigBox.Text = _bfOrigPath;
            LoadBackfillTargetColumns(_bfOrigPath);
        }
    }

    private void BfReplyBrowse_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Excel 文件 (*.xlsx)|*.xlsx",
            Title = "选择供应商回复单（可多选）",
            Multiselect = true
        };
        if (dlg.ShowDialog() == true)
        {
            _bfReplyPaths = new List<string>(dlg.FileNames);
            BfReplyBox.Text = string.Join("; ", _bfReplyPaths.ConvertAll(Path.GetFileName));
        }
    }

    private void BfRun_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_bfOrigPath) || !File.Exists(_bfOrigPath))
        { System.Windows.MessageBox.Show("请先选择原欠料表文件。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
        if (_bfReplyPaths.Count == 0)
        { System.Windows.MessageBox.Show("请选择至少一个供应商回复单。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
        // 目标列：从下拉框获取（若为空则回退到 O）
        string? target = (BfTargetCombo.SelectedItem as JudgeColumnItem)?.Letter;
        if (string.IsNullOrEmpty(target)) target = "O";
        string person = BfPersonBox.Text.Trim();

        try
        {
            Log($"回填：原表={Path.GetFileName(_bfOrigPath)} 回复单={_bfReplyPaths.Count} 目标列={target}");
            var sw = Stopwatch.StartNew();
            var res = QiliaoService.ProcessBackfill(_bfOrigPath, _bfReplyPaths, person, target);
            sw.Stop();
            if (!res.Success)
            {
                Log("回填失败：" + res.Error);
                return;
            }
            _lastBackfillPath = res.OutputPath;
            BfOpenBtn.Visibility = Visibility.Visible;
            string detail = $"回填完成：匹配并写回 {res.MatchedCount} 行（共提取回复 {res.ReplyTotal} 条，加载回复文件 {res.ReplyFilesLoaded} 个）。用时 {sw.ElapsedMilliseconds} ms。结果文件：{res.OutputPath}";
            if (res.ReplyErrors.Count > 0) detail += "\n提示：\n" + string.Join("\n", res.ReplyErrors);
            Log(detail);
        }
        catch (Exception ex)
        {
            Log("回填异常：" + ex.Message);
        }
    }

    private void BfOpen_Click(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrEmpty(_lastBackfillPath) && File.Exists(_lastBackfillPath))
            Process.Start(new ProcessStartInfo(_lastBackfillPath) { UseShellExecute = true });
    }
    #endregion
}
