using Markdig;
using Microsoft.Win32;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Navigation;
using Application = System.Windows.Application;
using Color = System.Windows.Media.Color;
using ColorConverter = System.Windows.Media.ColorConverter;
using ListBox = System.Windows.Controls.ListBox;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace MarkdownKnowledgeBase
{
    public partial class MainWindow : Window
    {
        private static readonly Regex TaskListRegex = new(@"^\s*[-*+]\s+\[(?<state>[ xX])\]", RegexOptions.Multiline);
        private readonly string _rootPath;
        private readonly string _metadataPath;
        private MetadataStore _metadata = new();
        private readonly Stack<NavigationEntry> _navigationStack = new();
        private readonly MarkdownPipeline _pipeline;
        private NoteItem? _currentNote;
        private bool _isEditorVisible = true;
        private bool _isPreviewVisible = true;
        private bool _isDarkMode;
        private System.Windows.Forms.NotifyIcon? _trayIcon;
        private bool _isExitRequested;
        private ScrollViewer? _editorScrollViewer;
        private bool _isSyncingScroll;
        private double _lastEditorScrollPercent;
        private bool _isPreviewReady;

        public MainWindow()
        {
            InitializeComponent();
            ApplyTheme(false);
            InitializeTrayIcon();
            _rootPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "MarkdownKnowledgeBase");
            _metadataPath = Path.Combine(_rootPath, ".metadata.json");
            _pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
            EnsureRoot();
            LoadMetadata();
            LoadCategories();
            RefreshMarkersAndLinks();
            UpdateEditorVisibility();
            UpdatePreviewVisibility();
            EditorBox.Loaded += OnEditorLoaded;
            PreviewBrowser.LoadCompleted += OnPreviewLoadCompleted;
            PreviewBrowser.ObjectForScripting = new PreviewBridge(this);
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            SetWindowDarkMode(_isDarkMode);
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            base.OnClosing(e);

            if (_isExitRequested)
            {
                return;
            }

            e.Cancel = true;
            Hide();
            ShowInTaskbar = false;
        }

        protected override void OnClosed(EventArgs e)
        {
            if (_trayIcon is not null)
            {
                _trayIcon.Visible = false;
                _trayIcon.Dispose();
                _trayIcon = null;
            }

            base.OnClosed(e);
        }

        private void EnsureRoot()
        {
            if (!Directory.Exists(_rootPath))
            {
                Directory.CreateDirectory(_rootPath);
            }
        }

        private void LoadMetadata()
        {
            _metadata = MetadataStore.Load(_metadataPath);
        }

        private void SaveMetadata()
        {
            _metadata.Save(_metadataPath);
            RefreshMarkersAndLinks();
        }

        private void LoadCategories()
        {
            CategoryList.ItemsSource = Directory.GetDirectories(_rootPath)
                .Select(Path.GetFileName)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .OrderBy(name => name)
                .ToList();
        }

        private void LoadNotes(string categoryName)
        {
            var categoryPath = Path.Combine(_rootPath, categoryName);
            var notes = Directory.Exists(categoryPath)
                ? Directory.GetFiles(categoryPath, "*.md").Select(path => new NoteItem(Path.GetFileNameWithoutExtension(path), path)).ToList()
                : new List<NoteItem>();
            NoteList.ItemsSource = notes;
        }

        private void RefreshMarkersAndLinks()
        {
            SourceMarkerList.ItemsSource = _metadata.Markers.Select(marker => DisplayMarker(marker)).ToList();
            TargetMarkerList.ItemsSource = _metadata.Markers.Select(marker => DisplayMarker(marker)).ToList();
            LinkList.ItemsSource = _metadata.Links.Select(link => DisplayLink(link)).ToList();
        }

        private string DisplayMarker(Marker marker)
        {
            var noteName = Path.GetFileNameWithoutExtension(marker.NotePath);
            return $"{marker.Name} ({noteName})";
        }

        private string DisplayLink(Link link)
        {
            var source = _metadata.Markers.FirstOrDefault(marker => marker.Id == link.SourceMarkerId);
            var target = _metadata.Markers.FirstOrDefault(marker => marker.Id == link.TargetMarkerId);
            var sourceLabel = source is null ? "未知" : DisplayMarker(source);
            var targetLabel = target is null ? "未知" : DisplayMarker(target);
            return $"{sourceLabel} -> {targetLabel}";
        }

        private void OnAddCategory(object sender, RoutedEventArgs e)
        {
            var name = Prompt("请输入分类名称");
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            var categoryPath = Path.Combine(_rootPath, name);
            if (!Directory.Exists(categoryPath))
            {
                Directory.CreateDirectory(categoryPath);
            }

            LoadCategories();
        }

        private void OnRefreshCategories(object sender, RoutedEventArgs e)
        {
            LoadCategories();
        }

        private void OnCategorySelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CategoryList.SelectedItem is string category)
            {
                LoadNotes(category);
            }
        }

        private void OnAddNote(object sender, RoutedEventArgs e)
        {
            if (CategoryList.SelectedItem is not string category)
            {
                MessageBox.Show("请先选择分类。");
                return;
            }

            var name = Prompt("请输入笔记名称");
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            var notePath = Path.Combine(_rootPath, category, $"{name}.md");
            if (!File.Exists(notePath))
            {
                File.WriteAllText(notePath, $"# {name}{Environment.NewLine}");
            }

            LoadNotes(category);
        }

        private void OnSaveNote(object sender, RoutedEventArgs e)
        {
            if (_currentNote is null)
            {
                MessageBox.Show("请先选择笔记。");
                return;
            }

            File.WriteAllText(_currentNote.Path, EditorBox.Text, Encoding.UTF8);
            UpdatePreview();
        }

        private void OnImportNote(object sender, RoutedEventArgs e)
        {
            if (CategoryList.SelectedItem is not string category)
            {
                MessageBox.Show("请先选择分类。");
                return;
            }

            var dialog = new OpenFileDialog
            {
                Filter = "Markdown 文件 (*.md)|*.md|所有文件 (*.*)|*.*",
                Title = "选择要导入的 Markdown 文件"
            };

            dialog.Multiselect = true;
            if (dialog.ShowDialog() != true)
            {
                return;
            }

            foreach (var sourcePath in dialog.FileNames)
            {
                if (!File.Exists(sourcePath))
                {
                    MessageBox.Show("选择的文件不存在。");
                    continue;
                }

                var fileName = Path.GetFileNameWithoutExtension(sourcePath);
                var destinationPath = Path.Combine(_rootPath, category, $"{fileName}.md");
                if (File.Exists(destinationPath))
                {
                    var overwrite = MessageBox.Show("目标笔记已存在，是否覆盖？", "确认覆盖", MessageBoxButton.YesNo);
                    if (overwrite != MessageBoxResult.Yes)
                    {
                        continue;
                    }
                }

                File.Copy(sourcePath, destinationPath, true);
            }
            LoadNotes(category);
        }

        private void OnDeleteNote(object sender, RoutedEventArgs e)
        {
            if (CategoryList.SelectedItem is not string category)
            {
                MessageBox.Show("请先选择分类。");
                return;
            }

            if (NoteList.SelectedItem is not NoteItem note)
            {
                MessageBox.Show("请先选择要删除的笔记。");
                return;
            }

            var confirm = MessageBox.Show($"确定删除笔记 \"{note.Name}\" 吗？", "确认删除", MessageBoxButton.YesNo);
            if (confirm != MessageBoxResult.Yes)
            {
                return;
            }

            if (File.Exists(note.Path))
            {
                File.Delete(note.Path);
            }

            var relativePath = GetRelativeNotePath(note.Path);
            var markersToRemove = _metadata.Markers.Where(marker => marker.NotePath == relativePath).ToList();
            RemoveMarkersAndLinks(markersToRemove);

            if (_currentNote?.Path == note.Path)
            {
                _currentNote = null;
                EditorBox.Text = string.Empty;
                UpdatePreview();
            }

            LoadNotes(category);
        }

        private void OnNoteSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (NoteList.SelectedItem is NoteItem note)
            {
                _currentNote = note;
                EditorBox.Text = File.Exists(note.Path) ? File.ReadAllText(note.Path) : string.Empty;
                UpdatePreview();
            }
        }

        private void OnEditorTextChanged(object sender, TextChangedEventArgs e)
        {
            UpdatePreview();
        }

        private void OnToggleEditor(object sender, RoutedEventArgs e)
        {
            _isEditorVisible = !_isEditorVisible;
            UpdateEditorVisibility();
        }

        private void OnTogglePreview(object sender, RoutedEventArgs e)
        {
            _isPreviewVisible = !_isPreviewVisible;
            UpdatePreviewVisibility();
        }

        private void UpdateEditorVisibility()
        {
            EditorBorder.Visibility = _isEditorVisible ? Visibility.Visible : Visibility.Collapsed;
            EditorColumn.Width = _isEditorVisible ? new GridLength(1, GridUnitType.Star) : new GridLength(0);
            ToggleEditorButton.Content = _isEditorVisible ? "隐藏编辑" : "显示编辑";
            UpdateEditorPreviewMargins();
        }

        private void UpdatePreviewVisibility()
        {
            PreviewBorder.Visibility = _isPreviewVisible ? Visibility.Visible : Visibility.Collapsed;
            PreviewColumn.Width = _isPreviewVisible ? new GridLength(1, GridUnitType.Star) : new GridLength(0);
            TogglePreviewButton.Content = _isPreviewVisible ? "隐藏预览" : "显示预览";
            UpdateEditorPreviewMargins();
        }

        private void UpdateEditorPreviewMargins()
        {
            EditorBorder.Margin = _isEditorVisible && _isPreviewVisible ? new Thickness(0, 0, 6, 0) : new Thickness(0);
            PreviewBorder.Margin = _isPreviewVisible && _isEditorVisible ? new Thickness(6, 0, 0, 0) : new Thickness(0);
        }

        private void UpdatePreview()
        {
            var html = Markdig.Markdown.ToHtml(EditorBox.Text, _pipeline);
            var palette = GetPreviewPalette();
            _isPreviewReady = false;
            var page = $@"<html><head><meta charset=""utf-8""><style>
                body{{font-family:'Segoe UI', sans-serif; padding:18px; background:{palette.Background}; color:{palette.Text};}}
                h1,h2,h3{{color:{palette.Heading};}}
                a{{color:{palette.Accent};}}
                pre{{background:{palette.CodeBackground}; padding:12px; border-radius:8px; border:1px solid {palette.Border};}}
                code{{background:{palette.InlineCodeBackground}; padding:2px 6px; border-radius:6px;}}
                blockquote{{border-left:4px solid {palette.Border}; padding-left:12px; color:{palette.MutedText};}}
                img{{max-width:100%;}}
                table{{border-collapse:collapse;}}
                th,td{{border:1px solid {palette.Border}; padding:6px 10px;}}
            </style></head><body>{html}
            <script>
                function getScrollPercent() {{
                    const doc = document.documentElement;
                    const body = document.body;
                    const scrollHeight = Math.max(body.scrollHeight, doc.scrollHeight);
                    const clientHeight = doc.clientHeight;
                    if (scrollHeight <= clientHeight) {{
                        return 0;
                    }}
                    const scrollTop = doc.scrollTop || body.scrollTop;
                    return scrollTop / (scrollHeight - clientHeight);
                }}
                function notifyScroll() {{
                    if (window.external && window.external.NotifyPreviewScroll) {{
                        window.external.NotifyPreviewScroll(getScrollPercent());
                    }}
                }}
                function setScroll(percent) {{
                    const doc = document.documentElement;
                    const body = document.body;
                    const scrollHeight = Math.max(body.scrollHeight, doc.scrollHeight);
                    const clientHeight = doc.clientHeight;
                    const top = percent * Math.max(0, scrollHeight - clientHeight);
                    doc.scrollTop = top;
                    body.scrollTop = top;
                }}
                function initTasks() {{
                    const boxes = document.querySelectorAll('input[type=""checkbox""]');
                    boxes.forEach((box, index) => {{
                        box.disabled = false;
                        box.dataset.taskIndex = index;
                        box.addEventListener('click', () => {{
                            if (window.external && window.external.ToggleTask) {{
                                window.external.ToggleTask(index, box.checked);
                            }}
                        }});
                    }});
                }}
                window.addEventListener('scroll', () => {{
                    window.requestAnimationFrame(notifyScroll);
                }});
                window.onload = () => {{
                    initTasks();
                    notifyScroll();
                }};
            </script></body></html>";
            PreviewBrowser.NavigateToString(page);
        }

        private void OnToggleTheme(object sender, RoutedEventArgs e)
        {
            _isDarkMode = !_isDarkMode;
            ApplyTheme(_isDarkMode);
            UpdatePreview();
        }

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int attributeValue, int attributeSize);

        private void SetWindowDarkMode(bool isDark)
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            if (hwnd == IntPtr.Zero)
            {
                return;
            }

            int useDark = isDark ? 1 : 0;
            const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
            const int DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19;

            if (DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ref useDark, sizeof(int)) != 0)
            {
                DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1, ref useDark, sizeof(int));
            }
        }

        private void ApplyTheme(bool isDark)
        {
            var resources = Application.Current.Resources;
            if (isDark)
            {
                SetBrush(resources, "AppBackground", "#0F172A");
                SetBrush(resources, "PanelBackground", "#111827");
                SetBrush(resources, "PanelBorder", "#1F2937");
                SetBrush(resources, "PrimaryText", "#E5E7EB");
                SetBrush(resources, "SecondaryText", "#9CA3AF");
                SetBrush(resources, "ButtonBackground", "#3B82F6");
                SetBrush(resources, "ButtonHover", "#2563EB");
                SetBrush(resources, "ButtonPressed", "#1D4ED8");
                SetBrush(resources, "ButtonBorder", "#1D4ED8");
                SetBrush(resources, "InputBackground", "#0B1220");
                SetBrush(resources, "InputBorder", "#233049");
                SetBrush(resources, "ListBackground", "#0B1220");
                SetBrush(resources, "ListItemHover", "#1F2937");
                SetBrush(resources, "SelectionBackground", "#1E3A8A");
                SetBrush(resources, "SelectionText", "#E5E7EB");
                ToggleThemeButton.Content = "日间模式";
            }
            else
            {
                SetBrush(resources, "AppBackground", "#F1F3F6");
                SetBrush(resources, "PanelBackground", "#FFFFFF");
                SetBrush(resources, "PanelBorder", "#E3E7EF");
                SetBrush(resources, "PrimaryText", "#111827");
                SetBrush(resources, "SecondaryText", "#6B7280");
                SetBrush(resources, "ButtonBackground", "#2563EB");
                SetBrush(resources, "ButtonHover", "#1D4ED8");
                SetBrush(resources, "ButtonPressed", "#1E40AF");
                SetBrush(resources, "ButtonBorder", "#1D4ED8");
                SetBrush(resources, "InputBackground", "#FFFFFF");
                SetBrush(resources, "InputBorder", "#D7DDE8");
                SetBrush(resources, "ListBackground", "#F9FAFB");
                SetBrush(resources, "ListItemHover", "#EEF2FF");
                SetBrush(resources, "SelectionBackground", "#DBEAFE");
                SetBrush(resources, "SelectionText", "#111827");
                ToggleThemeButton.Content = "夜间模式";
            }
        
            SetWindowDarkMode(isDark);
        }

        private static void SetBrush(ResourceDictionary resources, string key, string hex)
        {
            resources[key] = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
        }

        private PreviewPalette GetPreviewPalette()
        {
            return _isDarkMode
                ? new PreviewPalette("#0B1220", "#E5E7EB", "#F8FAFC", "#60A5FA", "#1F2937", "#111827", "#CBD5F5")
                : new PreviewPalette("#FFFFFF", "#1F2937", "#111827", "#2563EB", "#E5E7EB", "#F3F4F6", "#6B7280");
        }

        private readonly record struct PreviewPalette(
            string Background,
            string Text,
            string Heading,
            string Accent,
            string Border,
            string CodeBackground,
            string MutedText)
        {
            public string InlineCodeBackground => CodeBackground;
        }

        private void OnAddSourceMarker(object sender, RoutedEventArgs e)
        {
            AddMarker("请输入源标记名称", SourceMarkerList);
        }

        private void OnAddTargetMarker(object sender, RoutedEventArgs e)
        {
            AddMarker("请输入目标标记名称", TargetMarkerList);
        }

        private void AddMarker(string prompt, ListBox listBox)
        {
            if (_currentNote is null)
            {
                MessageBox.Show("请先选择笔记。");
                return;
            }

            var name = Prompt(prompt);
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            var marker = new Marker(Guid.NewGuid().ToString(), name, GetRelativeNotePath(_currentNote.Path), EditorBox.SelectionStart);
            _metadata.Markers.Add(marker);
            SaveMetadata();

            var display = DisplayMarker(marker);
            var markers = _metadata.Markers.Select(DisplayMarker).ToList();
            listBox.ItemsSource = markers;
            listBox.SelectedItem = display;
        }

        private void OnDeleteMarker(object sender, RoutedEventArgs e)
        {
            var marker = GetSelectedMarker(SourceMarkerList.SelectedItem as string)
                ?? GetSelectedMarker(TargetMarkerList.SelectedItem as string);
            if (marker is null)
            {
                MessageBox.Show("请选择要删除的标记。");
                return;
            }

            var confirm = MessageBox.Show($"确定删除标记 \"{marker.Name}\" 吗？", "确认删除", MessageBoxButton.YesNo);
            if (confirm != MessageBoxResult.Yes)
            {
                return;
            }

            RemoveMarkersAndLinks(new List<Marker> { marker });
        }

        private void OnCreateLink(object sender, RoutedEventArgs e)
        {
            var sourceMarker = GetSelectedMarker(SourceMarkerList.SelectedItem as string);
            var targetMarker = GetSelectedMarker(TargetMarkerList.SelectedItem as string);
            if (sourceMarker is null || targetMarker is null)
            {
                MessageBox.Show("请选择源标记和目标标记。");
                return;
            }

            var link = new Link(Guid.NewGuid().ToString(), sourceMarker.Id, targetMarker.Id);
            _metadata.Links.Add(link);
            SaveMetadata();
        }

        private void OnDeleteLink(object sender, RoutedEventArgs e)
        {
            if (LinkList.SelectedItem is not string linkDisplay)
            {
                MessageBox.Show("请选择要删除的跳转关系。");
                return;
            }

            var link = _metadata.Links.FirstOrDefault(item => DisplayLink(item) == linkDisplay);
            if (link is null)
            {
                return;
            }

            var confirm = MessageBox.Show("确定删除该跳转关系吗？", "确认删除", MessageBoxButton.YesNo);
            if (confirm != MessageBoxResult.Yes)
            {
                return;
            }

            _metadata.Links.Remove(link);
            SaveMetadata();
        }

        private void OnJump(object sender, RoutedEventArgs e)
        {
            if (LinkList.SelectedItem is not string linkDisplay)
            {
                MessageBox.Show("请选择跳转关系。");
                return;
            }

            var link = _metadata.Links.FirstOrDefault(item => DisplayLink(item) == linkDisplay);
            if (link is null)
            {
                return;
            }

            var targetMarker = _metadata.Markers.FirstOrDefault(marker => marker.Id == link.TargetMarkerId);
            if (targetMarker is null)
            {
                return;
            }

            if (_currentNote is not null)
            {
                _navigationStack.Push(new NavigationEntry(_currentNote.Path, EditorBox.SelectionStart));
            }

            NavigateToMarker(targetMarker);
        }

        private void OnBack(object sender, RoutedEventArgs e)
        {
            if (_navigationStack.Count == 0)
            {
                return;
            }

            var entry = _navigationStack.Pop();
            OpenNoteAt(entry.NotePath, entry.Position);
        }

        private void NavigateToMarker(Marker marker)
        {
            var notePath = Path.Combine(_rootPath, marker.NotePath);
            OpenNoteAt(notePath, marker.Position);
        }

        private void OpenNoteAt(string notePath, int position)
        {
            var category = Directory.GetParent(notePath)?.Name;
            if (category is null)
            {
                return;
            }

            CategoryList.SelectedItem = category;
            LoadNotes(category);
            var notes = NoteList.ItemsSource as List<NoteItem>;
            var note = notes?.FirstOrDefault(item => item.Path == notePath);
            if (note is null)
            {
                return;
            }

            NoteList.SelectedItem = note;
            _currentNote = note;
            EditorBox.Text = File.Exists(note.Path) ? File.ReadAllText(note.Path) : string.Empty;
            EditorBox.SelectionStart = Math.Clamp(position, 0, EditorBox.Text.Length);
            EditorBox.Focus();
            UpdatePreview();
        }

        private Marker? GetSelectedMarker(string? display)
        {
            if (string.IsNullOrWhiteSpace(display))
            {
                return null;
            }

            return _metadata.Markers.FirstOrDefault(marker => DisplayMarker(marker) == display);
        }

        private string GetRelativeNotePath(string notePath)
        {
            return Path.GetRelativePath(_rootPath, notePath);
        }

        private void RemoveMarkersAndLinks(List<Marker> markersToRemove)
        {
            if (markersToRemove.Count == 0)
            {
                return;
            }

            var markerIds = new HashSet<string>(markersToRemove.Select(marker => marker.Id));
            _metadata.Markers = _metadata.Markers.Where(marker => !markerIds.Contains(marker.Id)).ToList();
            _metadata.Links = _metadata.Links
                .Where(link => !markerIds.Contains(link.SourceMarkerId) && !markerIds.Contains(link.TargetMarkerId))
                .ToList();
            SaveMetadata();
        }

        private string? Prompt(string prompt)
        {
            var dialog = new InputDialog(prompt) { Owner = this };
            return dialog.ShowDialog() == true ? dialog.Response : null;
        }

        private void InitializeTrayIcon()
        {
            var menu = new System.Windows.Forms.ContextMenuStrip();
            var exitItem = new System.Windows.Forms.ToolStripMenuItem("退出");
            exitItem.Click += (_, _) => ExitApplication();
            menu.Items.Add(exitItem);

            _trayIcon = new System.Windows.Forms.NotifyIcon
            {
                Text = "Markdown Knowledge Base",
                Icon = LoadTrayIcon(),
                ContextMenuStrip = menu,
                Visible = true
            };
            _trayIcon.DoubleClick += (_, _) => ShowMainWindow();
        }

        private System.Drawing.Icon LoadTrayIcon()
        {
            try
            {
                var exePath = Process.GetCurrentProcess().MainModule?.FileName;
                if (!string.IsNullOrWhiteSpace(exePath))
                {
                    var icon = System.Drawing.Icon.ExtractAssociatedIcon(exePath);
                    if (icon is not null)
                    {
                        return icon;
                    }
                }
            }
            catch
            {
            }

            return System.Drawing.SystemIcons.Application;
        }

        private void ShowMainWindow()
        {
            if (!IsVisible)
            {
                Show();
            }

            if (WindowState == WindowState.Minimized)
            {
                WindowState = WindowState.Normal;
            }

            ShowInTaskbar = true;
            Activate();
        }

        private void ExitApplication()
        {
            _isExitRequested = true;
            Close();
        }

        private void OnEditorLoaded(object sender, RoutedEventArgs e)
        {
            _editorScrollViewer = FindChildScrollViewer(EditorBox);
            if (_editorScrollViewer is null)
            {
                return;
            }

            _editorScrollViewer.ScrollChanged += OnEditorScrollChanged;
        }

        private void OnPreviewLoadCompleted(object? sender, NavigationEventArgs e)
        {
            _isPreviewReady = true;
            SyncPreviewScroll(_lastEditorScrollPercent);
        }

        private void OnEditorScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (_isSyncingScroll || _editorScrollViewer is null)
            {
                return;
            }

            var percent = GetScrollPercent(_editorScrollViewer);
            _lastEditorScrollPercent = percent;
            SyncPreviewScroll(percent);
        }

        private void SyncPreviewScroll(double percent)
        {
            if (_isSyncingScroll)
            {
                return;
            }

            _isSyncingScroll = true;
            if (_isPreviewReady && PreviewBrowser.Document is not null)
            {
                try
                {
                    PreviewBrowser.InvokeScript("setScroll", percent);
                }
                catch (COMException)
                {
                    _isPreviewReady = false;
                }
                finally
                {
                    _isSyncingScroll = false;
                }
            }
            else
            {
                _isSyncingScroll = false;
            }
        }

        private void SyncEditorScroll(double percent)
        {
            if (_editorScrollViewer is null)
            {
                return;
            }

            if (_isSyncingScroll)
            {
                return;
            }

            _isSyncingScroll = true;
            var offset = percent * Math.Max(0, _editorScrollViewer.ExtentHeight - _editorScrollViewer.ViewportHeight);
            _editorScrollViewer.ScrollToVerticalOffset(offset);
            _lastEditorScrollPercent = percent;
            _isSyncingScroll = false;
        }

        private static double GetScrollPercent(ScrollViewer viewer)
        {
            if (viewer.ExtentHeight <= viewer.ViewportHeight)
            {
                return 0;
            }

            return viewer.VerticalOffset / (viewer.ExtentHeight - viewer.ViewportHeight);
        }

        private static ScrollViewer? FindChildScrollViewer(DependencyObject parent)
        {
            for (var i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is ScrollViewer viewer)
                {
                    return viewer;
                }

                var nested = FindChildScrollViewer(child);
                if (nested is not null)
                {
                    return nested;
                }
            }

            return null;
        }

        private void ToggleTask(int index, bool isChecked)
        {
            var matches = TaskListRegex.Matches(EditorBox.Text);
            if (index < 0 || index >= matches.Count)
            {
                return;
            }

            var match = matches[index];
            var stateGroup = match.Groups["state"];
            if (!stateGroup.Success)
            {
                return;
            }

            var text = EditorBox.Text;
            var chars = text.ToCharArray();
            chars[stateGroup.Index] = isChecked ? 'x' : ' ';
            var selectionStart = EditorBox.SelectionStart;
            var selectionLength = EditorBox.SelectionLength;
            EditorBox.Text = new string(chars);
            EditorBox.SelectionStart = Math.Clamp(selectionStart, 0, EditorBox.Text.Length);
            EditorBox.SelectionLength = selectionLength;
        }

        [ComVisible(true)]
        public sealed class PreviewBridge
        {
            private readonly MainWindow _owner;

            public PreviewBridge(MainWindow owner)
            {
                _owner = owner;
            }

            public void NotifyPreviewScroll(double percent)
            {
                _owner.SyncEditorScroll(percent);
            }

            public void ToggleTask(int index, bool isChecked)
            {
                _owner.ToggleTask(index, isChecked);
            }
        }
    }
}
