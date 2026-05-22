using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Microsoft.Win32;
using Word = Microsoft.Office.Interop.Word;
using Manuscript_guide.Services;
using Manuscript_guide.Models;
using Manuscript_guide.Scanners;

namespace Manuscript_guide.UI
{
    /// <summary>
    /// Interaction logic for TaskPaneWPFControl.xaml
    /// </summary>
    public partial class TaskPaneWPFControl : UserControl
    {
        private string currentModuleType = "";
        private List<IssueItem> currentIssues = new List<IssueItem>();
        private RadioButton radUseNativeSubscript;
        private RadioButton radUseUnicodeSubscript;

        public TaskPaneWPFControl()
        {
            InitializeComponent();

            // Register global event handlers from Ribbon or settings changes
            EventBus.ScanTriggered += OnScanTriggered;
            EventBus.FullScanRequested += OnFullScanRequested;
            EventBus.FullFixRequested += OnFullFixRequested;
            EventBus.ClearHighlightsRequested += OnClearHighlightsRequested;
            EventBus.OpenSettingsRequested += OnOpenSettingsRequested;

            // Load settings into UI on startup
            LoadSettingsToUI();
        }

        private void OnScanTriggered(string moduleType)
        {
            this.Dispatcher.Invoke(() =>
            {
                this.Visibility = Visibility.Visible;
                if (moduleType == "keyword")
                {
                    ShowPanel("keyword");
                }
                else
                {
                    ShowPanel("active");
                    txtTitleText.Text = "专项检测 - " + DiagnosticModuleRegistry.GetDisplayName(moduleType);
                    txtTitleIcon.Text = DiagnosticModuleRegistry.GetIcon(moduleType);

                    // Execute actual scan
                    RunActiveScan(moduleType);
                }
            });
        }

        private void OnClearHighlightsRequested(string moduleType)
        {
            this.Dispatcher.Invoke(() =>
            {
                try
                {
                    Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                    RunWithWordUiPaused(() => ShadingManager.ClearPluginTraces(doc, moduleType));

                    if (moduleType == "all" || moduleType == currentModuleType)
                    {
                        currentIssues.Clear();
                        itemsIssueList.ItemsSource = null;
                        txtErrCount.Text = "0";
                        txtFooterStatus.Text = "插件标记、留痕、书签与批注已全部清除。";
                        ShowPanel("overview");
                    }

                    ShowToast("插件生成的所有审阅痕迹已清除！");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("清除标记时发生错误：" + ex.Message, "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            });
        }

        private void OnFullScanRequested()
        {
            this.Dispatcher.Invoke(() =>
            {
                this.Visibility = Visibility.Visible;
                ShowPanel("active");
                txtTitleText.Text = "一键检测 - 全部专项";
                txtTitleIcon.Text = "⚙️";
                RunFullScan();
            });
        }

        private void OnFullFixRequested()
        {
            this.Dispatcher.Invoke(() =>
            {
                this.Visibility = Visibility.Visible;
                ShowPanel("active");
                txtTitleText.Text = "一键修复 - 全部专项";
                txtTitleIcon.Text = "⚙️";

                if (RunFullScan())
                {
                    FixCurrentIssues(true);
                }
            });
        }

        private void OnOpenSettingsRequested()
        {
            this.Dispatcher.Invoke(() =>
            {
                this.Visibility = Visibility.Visible;
                ShowPanel("rules");
            });
        }

        private void RunActiveScan(string moduleType)
        {
            currentModuleType = moduleType;
            currentIssues.Clear();
            itemsIssueList.ItemsSource = null;
            txtErrCount.Text = "0";
            emptyScanState.Visibility = Visibility.Collapsed;

            txtFooterStatus.Text = "正在扫描正文...";
            progressAudit.Value = 20;

            Word.Document doc = null;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                txtFooterStatus.Text = "获取文档错误";
                MessageBox.Show("未找到活动 Word 文档，请确保已打开 Word 并在正文中工作！", "错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (doc == null) return;

            try
            {
                RunWithWordUiPaused(() =>
                {
                    using (DocumentScanContext.Begin(doc))
                    {
                        ProtectedRangeService.RefreshProtectedMarkers(doc);

                        // Clear existing highlights of this module first
                        ShadingManager.ClearModuleShading(doc, moduleType);
                        progressAudit.Value = 40;

                        // Retrieve dedicated scanner
                        ISpecializedScanner scanner = DiagnosticModuleRegistry.CreateScanner(moduleType);
                        if (scanner == null)
                        {
                            txtFooterStatus.Text = "未知扫描类型";
                            return;
                        }

                        // Run scan in Word document
                        List<IssueItem> issues = IssueMetadataService.EnrichAll(scanner.Scan(doc));
                        currentIssues = ProtectedRangeService.FilterIssues(doc, issues);
                    }
                });

                itemsIssueList.ItemsSource = currentIssues;
                txtErrCount.Text = currentIssues.Count.ToString();
                emptyScanState.Visibility = currentIssues.Count == 0 ? Visibility.Visible : Visibility.Collapsed;

                progressAudit.Value = 100;
                txtFooterStatus.Text = $"扫描完成，发现 {currentIssues.Count} 处排版建议。";

                PopulateModuleConfigPanel(moduleType);
                ShowToast($"“{DiagnosticModuleRegistry.GetDisplayName(moduleType)}”扫描完成！");
            }
            catch (Exception ex)
            {
                txtFooterStatus.Text = "扫描失败";
                MessageBox.Show("扫描发生错误：" + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool RunFullScan()
        {
            currentModuleType = "all";
            currentIssues.Clear();
            itemsIssueList.ItemsSource = null;
            txtErrCount.Text = "0";
            emptyScanState.Visibility = Visibility.Collapsed;

            txtFooterStatus.Text = "正在执行全部专项检测...";
            progressAudit.Value = 5;

            Word.Document doc = null;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                txtFooterStatus.Text = "获取文档错误";
                MessageBox.Show("未找到活动 Word 文档，请确保已打开 Word 并在正文中工作！", "错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (doc == null)
            {
                return false;
            }

            try
            {
                RunWithWordUiPaused(() =>
                {
                    using (DocumentScanContext.Begin(doc))
                    {
                        ProtectedRangeService.RefreshProtectedMarkers(doc);

                        List<IssueItem> allIssues = new List<IssueItem>();
                        int completed = 0;

                        int moduleCount = DiagnosticModuleRegistry.FullScanModuleCount;
                        foreach (DiagnosticModule module in DiagnosticModuleRegistry.FullScanModules)
                        {
                            txtFooterStatus.Text = "正在扫描：" + module.DisplayName;
                            progressAudit.Value = 5 + (completed * 90 / moduleCount);

                            ShadingManager.ClearModuleShading(doc, module.ModuleType);
                            ISpecializedScanner scanner = module.CreateScanner();
                            if (scanner != null)
                            {
                                List<IssueItem> issues = IssueMetadataService.EnrichAll(scanner.Scan(doc));
                                allIssues.AddRange(ProtectedRangeService.FilterIssues(doc, issues));
                            }

                            completed++;
                        }

                        currentIssues = allIssues;
                    }
                });

                itemsIssueList.ItemsSource = currentIssues;
                txtErrCount.Text = currentIssues.Count.ToString();
                emptyScanState.Visibility = currentIssues.Count == 0 ? Visibility.Visible : Visibility.Collapsed;

                progressAudit.Value = 100;
                txtFooterStatus.Text = $"全部专项检测完成，发现 {currentIssues.Count} 处排版建议。";
                PopulateModuleConfigPanel("all");
                ShowToast($"一键检测完成，共发现 {currentIssues.Count} 处建议。");
                return true;
            }
            catch (Exception ex)
            {
                txtFooterStatus.Text = "一键检测失败";
                MessageBox.Show("一键检测发生错误：" + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        private void PopulateModuleConfigPanel(string moduleType)
        {
            panelModuleRules.Children.Clear();
            var settings = SettingsManager.Current;

            if (moduleType == "sub")
            {
                var tb = new TextBlock
                {
                    Text = settings.UseNativeSubscript ? "优先使用 Word 原生角标排版 (保留 ASCII)" : "当前使用 Unicode 物理角标替换",
                    FontSize = 10,
                    Foreground = new SolidColorBrush(settings.UseNativeSubscript ? Color.FromRgb(16, 124, 65) : Color.FromRgb(216, 124, 0)),
                    Margin = new Thickness(0, 0, 0, 4),
                    TextWrapping = TextWrapping.Wrap
                };
                panelModuleRules.Children.Add(tb);
            }
            else if (moduleType == "cap")
            {
                var tb = new TextBlock
                {
                    Text = $"多数派比例阈值 {settings.CasingRatioThreshold}% | 缩写最大滞后 {settings.MaxAcronymLagCharacters} 字符",
                    FontSize = 10,
                    Foreground = new SolidColorBrush(Color.FromRgb(96, 94, 92)),
                    Margin = new Thickness(0, 0, 0, 4)
                };
                panelModuleRules.Children.Add(tb);
            }
            else
            {
                var tb = new TextBlock
                {
                    Text = moduleType == "all" ? DiagnosticModuleRegistry.GetFullScanDescription() : "规则运行状态：正常激活",
                    FontSize = 10,
                    Foreground = new SolidColorBrush(Color.FromRgb(96, 94, 92))
                };
                panelModuleRules.Children.Add(tb);
            }
        }

        private void ShowPanel(string panelName)
        {
            panelOverview.Visibility = Visibility.Collapsed;
            panelActiveScan.Visibility = Visibility.Collapsed;
            panelKeyword.Visibility = Visibility.Collapsed;
            panelRules.Visibility = Visibility.Collapsed;

            switch (panelName)
            {
                case "overview":
                    panelOverview.Visibility = Visibility.Visible;
                    txtTitleText.Text = "排版助手 - 专项检查通道";
                    txtTitleIcon.Text = "⚙️";
                    break;
                case "active":
                    panelActiveScan.Visibility = Visibility.Visible;
                    break;
                case "keyword":
                    panelKeyword.Visibility = Visibility.Visible;
                    txtTitleText.Text = "快捷检索 - 关键词标记";
                    txtTitleIcon.Text = "🔍";
                    break;
                case "rules":
                    panelRules.Visibility = Visibility.Visible;
                    txtTitleText.Text = "排版助手 - 高级设置";
                    txtTitleIcon.Text = "⚙️";
                    break;
            }
        }

        private void ClosePane_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Globals.ThisAddIn.CustomPane != null)
                {
                    Globals.ThisAddIn.CustomPane.Visible = false;
                    return;
                }
            }
            catch (Exception)
            {
            }

            this.Visibility = Visibility.Collapsed;
        }

        private void GotoSettings_Click(object sender, RoutedEventArgs e)
        {
            ShowPanel("rules");
        }

        private void AcceptIssue_Click(object sender, RoutedEventArgs e)
        {
            Button btn = sender as Button;
            if (btn == null) return;
            string issueId = btn.Tag as string;
            if (string.IsNullOrEmpty(issueId)) return;

            IssueItem issue = currentIssues.Find(x => x.IssueId == issueId);
            if (issue == null) return;

            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            try
            {
                Action<Word.Range> customFormatAction = CorrectionActionResolver.Resolve(issue, doc);
                bool success = CorrectionTracker.Instance.ExecuteCorrection(doc, issueId, issue.Type, issue.RecommendFix, customFormatAction);
                if (!success)
                {
                    MessageBox.Show("该问题的 Word 定位书签已失效，请重新扫描后再应用修改。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                issue.IsFixed = true;
                UpdatePendingCount();
                ShowToast("已应用修改，并保留淡色底纹与 Word 批注。");
            }
            catch (Exception ex)
            {
                MessageBox.Show("应用修改失败：" + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void IgnoreIssue_Click(object sender, RoutedEventArgs e)
        {
            Button btn = sender as Button;
            if (btn == null) return;
            string issueId = btn.Tag as string;
            if (string.IsNullOrEmpty(issueId)) return;

            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            try
            {
                string bookmarkName = CorrectionTracker.Instance.FindBookmarkName(doc, issueId);
                if (doc.Bookmarks.Exists(bookmarkName))
                {
                    Word.Range r = doc.Bookmarks[bookmarkName].Range;
                    r.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;
                    doc.Bookmarks[bookmarkName].Delete();
                }
            }
            catch (Exception) { }

            IssueItem issue = currentIssues.Find(x => x.IssueId == issueId);
            if (issue != null)
            {
                issue.IsIgnored = true;
            }
            UpdatePendingCount();
        }

        private void UndoIssue_Click(object sender, RoutedEventArgs e)
        {
            Button btn = sender as Button;
            if (btn == null) return;
            string issueId = btn.Tag as string;
            if (string.IsNullOrEmpty(issueId)) return;

            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            try
            {
                IssueItem issue = currentIssues.Find(x => x.IssueId == issueId);
                string moduleType = issue != null ? issue.Type : currentModuleType;
                bool success = CorrectionTracker.Instance.ExecuteUndo(doc, issueId, moduleType);
                if (success)
                {
                    if (issue != null)
                    {
                        issue.IsFixed = false;
                    }

                    UpdatePendingCount();
                    ShowToast("成功撤销改写，底纹高亮已恢复。");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("撤销失败：" + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void FixAllModule_Click(object sender, RoutedEventArgs e)
        {
            FixCurrentIssues(false);
        }

        private void FixCurrentIssues(bool isFullFix)
        {
            if (currentIssues == null || currentIssues.Count == 0) return;

            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            if (SettingsManager.IsRuleEnabled("global", "auto_backup_before_bulk_fix"))
            {
                BackupDocument(doc);
            }

            int fixedCount = 0;
            RunWithWordUiPaused(() =>
            {
                foreach (var issue in currentIssues)
                {
                    if (CorrectionTracker.Instance.CanUndo(issue.IssueId)) continue; // Already corrected

                    try
                    {
                        Action<Word.Range> customFormatAction = CorrectionActionResolver.Resolve(issue, doc);
                        if (CorrectionTracker.Instance.ExecuteCorrection(doc, issue.IssueId, issue.Type, issue.RecommendFix, customFormatAction))
                        {
                            issue.IsFixed = true;
                            fixedCount++;
                        }
                    }
                    catch (Exception) { }
                }
            });

            UpdatePendingCount();
            string scope = isFullFix ? "全部专项" : "本模块";
            txtFooterStatus.Text = $"{scope}一键修复成功！共修复 {fixedCount} 处排版问题。底纹已转为淡色留痕。";
            ShowToast($"{scope}一键修复完成，共处理 {fixedCount} 处！");
        }

        private void UpdatePendingCount()
        {
            int pending = 0;
            foreach (var issue in currentIssues)
            {
                if (!issue.IsFixed && !issue.IsIgnored)
                {
                    pending++;
                }
            }

            txtErrCount.Text = pending.ToString();
        }

        private void BackupDocument(Word.Document doc)
        {
            try
            {
                string path = doc.FullName;
                if (File.Exists(path))
                {
                    string dir = Path.GetDirectoryName(path);
                    string name = Path.GetFileNameWithoutExtension(path);
                    string ext = Path.GetExtension(path);
                    string backupPath = Path.Combine(dir, $"{name}_backup_{DateTime.Now:yyyyMMddHHmmss}{ext}");
                    File.Copy(path, backupPath, true);
                }
            }
            catch (Exception) { }
        }

        private void ToggleCardState(Button btn, bool isCorrected)
        {
            Grid parentGrid = FindParent<Grid>(btn);
            if (parentGrid != null)
            {
                StackPanel normal = parentGrid.FindName("panelNormalState") as StackPanel;
                StackPanel corrected = parentGrid.FindName("panelCorrectedState") as StackPanel;
                if (normal != null && corrected != null)
                {
                    normal.Visibility = isCorrected ? Visibility.Collapsed : Visibility.Visible;
                    corrected.Visibility = isCorrected ? Visibility.Visible : Visibility.Collapsed;
                }
            }
        }

        private void IssueCard_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            DependencyObject source = e.OriginalSource as DependencyObject;
            if (source != null && FindParent<Button>(source) != null)
            {
                return;
            }

            FrameworkElement element = sender as FrameworkElement;
            IssueItem issue = element == null ? null : element.DataContext as IssueItem;
            if (issue == null)
            {
                return;
            }

            JumpToIssue(issue);
            e.Handled = true;
        }

        private void JumpToIssue(IssueItem issue)
        {
            try
            {
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                Word.Range range = null;

                string bookmarkName = CorrectionTracker.Instance.FindBookmarkName(doc, issue.IssueId);
                if (!string.IsNullOrEmpty(bookmarkName) && doc.Bookmarks.Exists(bookmarkName))
                {
                    range = doc.Bookmarks[bookmarkName].Range;
                }
                else if (issue.Start >= 0 && issue.End > issue.Start && issue.End <= doc.Content.End)
                {
                    range = doc.Range(issue.Start, issue.End);
                }

                if (range == null)
                {
                    ShowToast("原文定位已失效，请重新扫描。");
                    return;
                }

                range.Select();
                Globals.ThisAddIn.Application.Activate();
                ShowToast("已跳转到原文位置。");
            }
            catch
            {
                ShowToast("原文定位已失效，请重新扫描。");
            }
        }

        private T FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            if (child == null)
            {
                return null;
            }

            DependencyObject parent = VisualTreeHelper.GetParent(child);
            while (parent != null && !(parent is T))
            {
                parent = VisualTreeHelper.GetParent(parent);
            }
            return parent as T;
        }

        private T FindChild<T>(DependencyObject parent, string childName) where T : DependencyObject
        {
            if (parent == null) return null;

            int childrenCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childrenCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T tChild && (string.IsNullOrEmpty(childName) || (child is FrameworkElement fe && fe.Name == childName)))
                {
                    return tChild;
                }

                T childOfChild = FindChild<T>(child, childName);
                if (childOfChild != null)
                    return childOfChild;
            }
            return null;
        }

        private void StartKeywordSearch_Click(object sender, RoutedEventArgs e)
        {
            string keywords = txtKeywordInput.Text;
            if (string.IsNullOrEmpty(keywords))
            {
                MessageBox.Show("请先输入查找关键词！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool caseSensitive = chkCaseSensitive.IsChecked == true;
            bool wholeWord = chkWholeWord.IsChecked == true;

            string colorHex = "#FFFF33"; // Yellow
            if (radColorGreen.IsChecked == true) colorHex = "#5CFF5C";
            if (radColorCyan.IsChecked == true) colorHex = "#00E5E5";
            if (radColorPink.IsChecked == true) colorHex = "#FF66FF";
            SettingsManager.Current.Colors["keyword"] = colorHex;

            Word.Document doc = null;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                MessageBox.Show("获取当前 Word 文档失败。", "错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (doc == null) return;

            try
            {
                RunWithWordUiPaused(() =>
                {
                    int count = 0;
                    using (DocumentScanContext.Begin(doc))
                    {
                        ProtectedRangeService.RefreshProtectedMarkers(doc);

                        ShadingManager.ClearModuleShading(doc, "keyword");

                        string[] words = keywords.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var rawWord in words)
                        {
                            string w = rawWord.Trim();
                            if (string.IsNullOrEmpty(w)) continue;

                            Word.Range range = doc.Content;
                            Word.Find find = range.Find;
                            find.ClearFormatting();
                            find.Text = w;
                            find.MatchCase = caseSensitive;
                            find.MatchWholeWord = wholeWord;
                            find.Forward = true;
                            find.Wrap = Word.WdFindWrap.wdFindStop;

                            while (find.Execute())
                            {
                                if (ProtectedRangeService.IsRangeProtected(range))
                                {
                                    int skipStart = range.End;
                                    if (skipStart >= doc.Content.End)
                                    {
                                        break;
                                    }

                                    range.SetRange(skipStart, doc.Content.End);
                                    find = range.Find;
                                    find.ClearFormatting();
                                    find.Text = w;
                                    find.MatchCase = caseSensitive;
                                    find.MatchWholeWord = wholeWord;
                                    find.Forward = true;
                                    find.Wrap = Word.WdFindWrap.wdFindStop;
                                    continue;
                                }

                                string issueId = Guid.NewGuid().ToString();
                                CorrectionTracker.Instance.CreateBookmark(doc, issueId, range, "keyword");
                                range.Shading.BackgroundPatternColor = SettingsManager.HexToWdColor(colorHex);
                                count++;

                                int nextStart = range.End;
                                if (nextStart >= doc.Content.End)
                                {
                                    break;
                                }

                                range.SetRange(nextStart, doc.Content.End);
                                find = range.Find;
                                find.ClearFormatting();
                                find.Text = w;
                                find.MatchCase = caseSensitive;
                                find.MatchWholeWord = wholeWord;
                                find.Forward = true;
                                find.Wrap = Word.WdFindWrap.wdFindStop;
                            }
                        }
                    }

                    boxSearchStats.Visibility = Visibility.Visible;
                    txtSearchStatsResult.Text = $"共查找到并标记了 {count} 处匹配的关键词。";
                    ShowToast($"查找标记完成，共找到 {count} 处！");
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("关键词查找标记发生错误：" + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadSettingsToUI()
        {
            var settings = SettingsManager.Current;
            SyncLegacySettingsFromRuleToggles(settings);

            txtWhitelistCasing.Text = settings.WhitelistCasing;
            txtWhitelistAcronyms.Text = settings.WhitelistAcronyms;

            sldCasingRatio.Value = settings.CasingRatioThreshold;
            txtCasingRatioVal.Text = settings.CasingRatioThreshold + "%";

            txtAcronymLag.Text = settings.MaxAcronymLagCharacters.ToString();

            // Load Colors
            var colors = settings.Colors;
            if (colors != null)
            {
                if (colors.ContainsKey("punc")) { txtColorPunc.Text = colors["punc"]; SetColorPreview(rectColorPunc, colors["punc"]); }
                if (colors.ContainsKey("cap")) { txtColorCap.Text = colors["cap"]; SetColorPreview(rectColorCap, colors["cap"]); }
                if (colors.ContainsKey("dash")) { txtColorDash.Text = colors["dash"]; SetColorPreview(rectColorDash, colors["dash"]); }
                if (colors.ContainsKey("data")) { txtColorData.Text = colors["data"]; SetColorPreview(rectColorData, colors["data"]); }
                if (colors.ContainsKey("ital")) { txtColorItal.Text = colors["ital"]; SetColorPreview(rectColorItal, colors["ital"]); }
                if (colors.ContainsKey("sub")) { txtColorSub.Text = colors["sub"]; SetColorPreview(rectColorSub, colors["sub"]); }
                if (colors.ContainsKey("word")) { txtColorWord.Text = colors["word"]; SetColorPreview(rectColorWord, colors["word"]); }
                if (colors.ContainsKey("keyword")) { txtColorKeyword.Text = colors["keyword"]; SetColorPreview(rectColorKeyword, colors["keyword"]); }
            }

            BuildRuleTogglePanel(settings);
        }

        private void SaveSettings_Click(object sender, RoutedEventArgs e)
        {
            var settings = SettingsManager.Current;

            settings.UseNativeSubscript = radUseNativeSubscript == null || radUseNativeSubscript.IsChecked == true;

            settings.WhitelistCasing = txtWhitelistCasing.Text;
            settings.WhitelistAcronyms = txtWhitelistAcronyms.Text;

            settings.CasingRatioThreshold = (int)sldCasingRatio.Value;

            if (int.TryParse(txtAcronymLag.Text, out int lag))
            {
                settings.MaxAcronymLagCharacters = lag;
            }

            // Save Colors
            if (settings.Colors == null) settings.Colors = new Dictionary<string, string>();
            settings.Colors["punc"] = txtColorPunc.Text;
            settings.Colors["cap"] = txtColorCap.Text;
            settings.Colors["dash"] = txtColorDash.Text;
            settings.Colors["data"] = txtColorData.Text;
            settings.Colors["ital"] = txtColorItal.Text;
            settings.Colors["sub"] = txtColorSub.Text;
            settings.Colors["word"] = txtColorWord.Text;
            settings.Colors["keyword"] = txtColorKeyword.Text;

            SaveRuleTogglePanel(settings);
            SyncLegacySettingsFromRuleToggles(settings);

            SettingsManager.Save();
            ShowToast("系统高级设置已保存并应用。");
        }

        private void BuildRuleTogglePanel(PluginSettings settings)
        {
            if (panelRuleToggles == null)
            {
                return;
            }

            panelRuleToggles.Children.Clear();
            AddRuleGroup("global", "全局设置", true, settings);

            foreach (DiagnosticModule module in DiagnosticModuleRegistry.AllModules)
            {
                AddRuleGroup(module.ModuleType, module.DisplayName, module.ModuleType == "sub", settings);
            }
        }

        private void AddRuleGroup(string moduleType, string displayName, bool isExpanded, PluginSettings settings)
        {
            Expander expander = new Expander
            {
                Header = displayName,
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(50, 49, 48)),
                Margin = new Thickness(0, 0, 0, 8),
                IsExpanded = isExpanded
            };

            StackPanel ruleStack = new StackPanel
            {
                Margin = new Thickness(8, 8, 0, 2)
            };

            if (moduleType == "sub")
            {
                AddSubscriptOutputModeControls(ruleStack, settings);
            }

            foreach (DiagnosticRuleDefinition rule in DiagnosticRuleRegistry.GetRulesForModule(moduleType))
            {
                Border row = new Border
                {
                    Background = new SolidColorBrush(Color.FromRgb(250, 250, 250)),
                    BorderBrush = new SolidColorBrush(Color.FromRgb(237, 235, 233)),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(4),
                    Padding = new Thickness(10, 8, 10, 8),
                    Margin = new Thickness(0, 0, 0, 8)
                };

                StackPanel rowStack = new StackPanel();
                CheckBox checkBox = new CheckBox
                {
                    Content = rule.Title,
                    Tag = rule.Key,
                    IsChecked = GetRuleEnabled(settings, rule),
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = new SolidColorBrush(Color.FromRgb(50, 49, 48)),
                    ToolTip = rule.Description,
                    Margin = new Thickness(0, 0, 0, 4)
                };

                TextBlock description = new TextBlock
                {
                    Text = rule.Description,
                    FontSize = 10,
                    Foreground = new SolidColorBrush(Color.FromRgb(96, 94, 92)),
                    TextWrapping = TextWrapping.Wrap,
                    LineHeight = 15,
                    Margin = new Thickness(22, 0, 0, 0)
                };

                rowStack.Children.Add(checkBox);
                rowStack.Children.Add(description);
                row.Child = rowStack;
                ruleStack.Children.Add(row);
            }

            expander.Content = ruleStack;
            panelRuleToggles.Children.Add(expander);
        }

        private void AddSubscriptOutputModeControls(StackPanel ruleStack, PluginSettings settings)
        {
            Border row = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(250, 250, 250)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(237, 235, 233)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(10, 9, 10, 9),
                Margin = new Thickness(0, 0, 0, 8)
            };

            StackPanel stack = new StackPanel();
            TextBlock title = new TextBlock
            {
                Text = "正文角标输出方式",
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(50, 49, 48)),
                Margin = new Thickness(0, 0, 0, 7)
            };

            radUseNativeSubscript = new RadioButton
            {
                GroupName = "SubscriptOutputMode",
                Content = "Word 原生角标",
                IsChecked = settings.UseNativeSubscript,
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(Color.FromRgb(216, 124, 0)),
                Margin = new Thickness(0, 0, 0, 5),
                ToolTip = "默认模式。正文保留普通数字字符，并通过 Word Font.Subscript 或 Font.Superscript 显示为角标，适合正式论文排版。"
            };

            TextBlock nativeDescription = new TextBlock
            {
                Text = "保留 WSe2、cm-3 等普通文本字符；修复时通过 Word 原生字体属性显示下标或上标。",
                FontSize = 10,
                Foreground = new SolidColorBrush(Color.FromRgb(96, 94, 92)),
                TextWrapping = TextWrapping.Wrap,
                LineHeight = 15,
                Margin = new Thickness(22, 0, 0, 7)
            };

            radUseUnicodeSubscript = new RadioButton
            {
                GroupName = "SubscriptOutputMode",
                Content = "Unicode 角标字符",
                IsChecked = !settings.UseNativeSubscript,
                FontSize = 11,
                Foreground = new SolidColorBrush(Color.FromRgb(96, 94, 92)),
                Margin = new Thickness(0, 0, 0, 5),
                ToolTip = "纯文本兼容模式。修复时直接输出 ¹、²、³、₂、₃ 等特殊 Unicode 角标字符。"
            };

            TextBlock unicodeDescription = new TextBlock
            {
                Text = "直接写入 Unicode 角标字符，仅建议在纯文本兼容或无法保留 Word 字体属性的场景中使用。",
                FontSize = 10,
                Foreground = new SolidColorBrush(Color.FromRgb(96, 94, 92)),
                TextWrapping = TextWrapping.Wrap,
                LineHeight = 15,
                Margin = new Thickness(22, 0, 0, 0)
            };

            stack.Children.Add(title);
            stack.Children.Add(radUseNativeSubscript);
            stack.Children.Add(nativeDescription);
            stack.Children.Add(radUseUnicodeSubscript);
            stack.Children.Add(unicodeDescription);
            row.Child = stack;
            ruleStack.Children.Add(row);
        }

        private bool GetRuleEnabled(PluginSettings settings, DiagnosticRuleDefinition rule)
        {
            if (settings.EnabledRules == null)
            {
                settings.EnabledRules = new Dictionary<string, bool>();
            }

            if (!settings.EnabledRules.ContainsKey(rule.Key))
            {
                settings.EnabledRules[rule.Key] = rule.DefaultEnabled;
            }

            return settings.EnabledRules[rule.Key];
        }

        private void SaveRuleTogglePanel(PluginSettings settings)
        {
            if (settings.EnabledRules == null)
            {
                settings.EnabledRules = new Dictionary<string, bool>();
            }

            SaveRuleTogglePanelRecursive(panelRuleToggles, settings);
        }

        private void SaveRuleTogglePanelRecursive(DependencyObject parent, PluginSettings settings)
        {
            if (parent == null)
            {
                return;
            }

            foreach (object logicalChild in LogicalTreeHelper.GetChildren(parent))
            {
                DependencyObject child = logicalChild as DependencyObject;
                if (child == null)
                {
                    continue;
                }

                CheckBox checkBox = child as CheckBox;
                if (checkBox != null && checkBox.Tag is string key)
                {
                    settings.EnabledRules[key] = checkBox.IsChecked == true;
                }

                SaveRuleTogglePanelRecursive(child, settings);
            }
        }

        private void SyncLegacySettingsFromRuleToggles(PluginSettings settings)
        {
            if (settings.EnabledRules == null)
            {
                return;
            }

            settings.PreserveUserHighlights = GetStoredRuleEnabled(settings, "global.preserve_user_highlights", settings.PreserveUserHighlights);
            settings.AutoBackup = GetStoredRuleEnabled(settings, "global.auto_backup_before_bulk_fix", settings.AutoBackup);
            settings.UnifyGreekMu = GetStoredRuleEnabled(settings, "punc.greek_mu_encoding", settings.UnifyGreekMu);
            settings.EquationPunctuation =
                GetStoredRuleEnabled(settings, "punc.equation_spacing", settings.EquationPunctuation) ||
                GetStoredRuleEnabled(settings, "punc.equation_terminal_punctuation", settings.EquationPunctuation);
            settings.CrossRefCapitalization = GetStoredRuleEnabled(settings, "cap.crossref_capitalization", settings.CrossRefCapitalization);
            settings.VariableLock = GetStoredRuleEnabled(settings, "cap.physical_variable_casing_lock", settings.VariableLock);
            settings.DetectExistingItalics = GetStoredRuleEnabled(settings, "ital.existing_italics_review", settings.DetectExistingItalics);
            settings.EnableElementSubscriptConversion = GetStoredRuleEnabled(settings, "sub.element_formula_subscript", settings.EnableElementSubscriptConversion);
        }

        private bool GetStoredRuleEnabled(PluginSettings settings, string key, bool fallback)
        {
            if (settings.EnabledRules != null && settings.EnabledRules.ContainsKey(key))
            {
                return settings.EnabledRules[key];
            }

            return fallback;
        }

        private void SeedRuleTogglesFromLegacySettings(PluginSettings settings)
        {
            if (settings.EnabledRules == null)
            {
                settings.EnabledRules = new Dictionary<string, bool>();
            }

            settings.EnabledRules["global.filter_references_and_citations"] = true;
            settings.EnabledRules["global.skip_latex_formula_regions"] = true;
            settings.EnabledRules["global.mark_skipped_formula_regions"] = false;
            settings.EnabledRules["global.preserve_user_highlights"] = settings.PreserveUserHighlights;
            settings.EnabledRules["global.auto_backup_before_bulk_fix"] = settings.AutoBackup;
            settings.EnabledRules["punc.greek_mu_encoding"] = settings.UnifyGreekMu;
            settings.EnabledRules["punc.equation_spacing"] = settings.EquationPunctuation;
            settings.EnabledRules["punc.equation_terminal_punctuation"] = settings.EquationPunctuation;
            settings.EnabledRules["cap.crossref_capitalization"] = settings.CrossRefCapitalization;
            settings.EnabledRules["cap.physical_variable_casing_lock"] = settings.VariableLock;
            settings.EnabledRules["ital.existing_italics_review"] = settings.DetectExistingItalics;
            settings.EnabledRules["sub.element_formula_subscript"] = settings.EnableElementSubscriptConversion;
        }

        private void SldCasingRatio_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (txtCasingRatioVal != null)
            {
                txtCasingRatioVal.Text = (int)sldCasingRatio.Value + "%";
            }
        }

        private void ExportSettings_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "JSON Files (*.json)|*.json",
                DefaultExt = "json",
                FileName = "manuscript_settings.json"
            };

            if (sfd.ShowDialog() == true)
            {
                try
                {
                    var settings = SettingsManager.Current;
                    var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                    string json = serializer.Serialize(settings);
                    File.WriteAllText(sfd.FileName, json);
                    ShowToast("配置成功导出。");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导出失败：" + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ImportSettings_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "JSON Files (*.json)|*.json",
                DefaultExt = "json"
            };

            if (ofd.ShowDialog() == true)
            {
                try
                {
                    string json = File.ReadAllText(ofd.FileName);
                    var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                    PluginSettings settings = serializer.Deserialize<PluginSettings>(json);
                    if (settings != null)
                    {
                        var current = SettingsManager.Current;
                        current.UseNativeSubscript = settings.UseNativeSubscript;
                        current.EnableElementSubscriptConversion = settings.EnableElementSubscriptConversion;
                        current.UnifyGreekMu = settings.UnifyGreekMu;
                        current.PreserveUserHighlights = settings.PreserveUserHighlights;
                        current.EquationPunctuation = settings.EquationPunctuation;
                        current.CrossRefCapitalization = settings.CrossRefCapitalization;
                        current.AutoBackup = settings.AutoBackup;
                        current.VariableLock = settings.VariableLock;
                        current.DetectExistingItalics = settings.DetectExistingItalics;
                        current.WhitelistCasing = settings.WhitelistCasing;
                        current.WhitelistAcronyms = settings.WhitelistAcronyms;
                        current.CasingRatioThreshold = settings.CasingRatioThreshold;
                        current.MaxAcronymLagCharacters = settings.MaxAcronymLagCharacters;
                        if (settings.Colors != null)
                        {
                            current.Colors = settings.Colors;
                        }
                        if (settings.EnabledRules != null)
                        {
                            current.EnabledRules = new Dictionary<string, bool>(settings.EnabledRules);
                        }
                        else
                        {
                            SeedRuleTogglesFromLegacySettings(current);
                        }

                        SettingsManager.Save();
                        LoadSettingsToUI();
                        ShowToast("配置成功导入并应用！");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导入失败：" + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void RectColor_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Border rect = sender as Border;
            if (rect == null) return;
            string key = rect.Tag as string;
            if (string.IsNullOrEmpty(key)) return;

            TextBox txt = FindName("txtColor" + char.ToUpper(key[0]) + key.Substring(1)) as TextBox;
            if (txt == null) return;

            using (var dialog = new System.Windows.Forms.ColorDialog())
            {
                try
                {
                    var color = (Color)ColorConverter.ConvertFromString(txt.Text);
                    dialog.Color = System.Drawing.Color.FromArgb(color.R, color.G, color.B);
                }
                catch { }

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string hex = string.Format("#{0:X2}{1:X2}{2:X2}", dialog.Color.R, dialog.Color.G, dialog.Color.B);
                    txt.Text = hex;
                    rect.Background = new SolidColorBrush(Color.FromRgb(dialog.Color.R, dialog.Color.G, dialog.Color.B));
                }
            }
        }

        private void ColorTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox txt = sender as TextBox;
            if (txt == null) return;
            string key = txt.Tag as string;
            if (string.IsNullOrEmpty(key)) return;

            Border rect = FindName("rectColor" + char.ToUpper(key[0]) + key.Substring(1)) as Border;
            if (rect == null) return;

            try
            {
                string hex = txt.Text;
                if (!string.IsNullOrEmpty(hex) && hex.StartsWith("#") && hex.Length == 7)
                {
                    var color = (Color)ColorConverter.ConvertFromString(hex);
                    rect.Background = new SolidColorBrush(color);
                }
            }
            catch { }
        }

        private void SetColorPreview(Border border, string hex)
        {
            try
            {
                if (border == null || string.IsNullOrEmpty(hex)) return;
                var color = (Color)ColorConverter.ConvertFromString(hex);
                border.Background = new SolidColorBrush(color);
            }
            catch { }
        }

        private void RunWithWordUiPaused(Action action)
        {
            Word.Application app = Globals.ThisAddIn.Application;
            bool originalScreenUpdating = app.ScreenUpdating;

            try
            {
                app.ScreenUpdating = false;
                action();
            }
            finally
            {
                app.ScreenUpdating = originalScreenUpdating;
            }
        }

        private void ShowToast(string message)
        {
            txtToastMessage.Text = message;
            borderToast.Visibility = Visibility.Visible;

            var timer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(3)
            };
            timer.Tick += (s, ev) =>
            {
                borderToast.Visibility = Visibility.Collapsed;
                timer.Stop();
            };
            timer.Start();
        }
    }
}
