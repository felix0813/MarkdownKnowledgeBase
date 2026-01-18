using Markdig;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using Application = System.Windows.Application;
using Button = System.Windows.Controls.Button;
using Color = System.Windows.Media.Color;
using ColorConverter = System.Windows.Media.ColorConverter;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace MarkdownKnowledgeBase
{
    public partial class MainWindow : Window
    {
        private readonly string _rootPath;
        private readonly string _metadataPath;
        private MetadataStore _metadata = new();
        private readonly Stack<NavigationEntry> _navigationStack = new();
        private readonly MarkdownPipeline _pipeline;
        private NoteItem? _currentNote;
        private bool _isEditorVisible = true;
        private bool _isPreviewVisible = true;
        private bool _isDarkMode;
        private NotifyIcon? _trayIcon;
        private bool _isExitRequested;
        private ScrollViewer? _editorScrollViewer;
        private ScrollViewer? _previewScrollViewer;
        private bool _isSyncingScroll;

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
            RefreshLineMarkers();
            UpdateEditorVisibility();
            UpdatePreviewVisibility();
            Loaded += OnWindowLoaded;
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            SetWindowDarkMode(_isDarkMode);
        }

        private void OnWindowLoaded(object sender, RoutedEventArgs e)
        {
            EditorBox.ApplyTemplate();
            PreviewViewer.ApplyTemplate();

            _editorScrollViewer = FindDescendantScrollViewer(EditorBox);
            _previewScrollViewer = FindDescendantScrollViewer(PreviewViewer);

            if (_editorScrollViewer is not null)
            {
                _editorScrollViewer.ScrollChanged += OnEditorScrollChanged;
            }

            if (_previewScrollViewer is not null)
            {
                _previewScrollViewer.ScrollChanged += OnPreviewScrollChanged;
            }

            EditorBox.SizeChanged += (_, _) => RefreshLineMarkers();
            RefreshLineMarkers();
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
            RefreshLineMarkers();
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

        private void RefreshLineMarkers()
        {
            MarkerOverlay.Children.Clear();

            if (_currentNote is null)
            {
                return;
            }

            var relativePath = GetRelativeNotePath(_currentNote.Path);
            var links = _metadata.LineLinks.Where(link => link.SourceNotePath == relativePath).ToList();
            foreach (var link in links)
            {
                var lineIndex = Math.Clamp(link.SourceLineIndex, 0, Math.Max(0, EditorBox.LineCount - 1));
                var charIndex = EditorBox.GetCharacterIndexFromLineIndex(lineIndex);
                var rect = EditorBox.GetRectFromCharacterIndex(charIndex);
                if (rect.IsEmpty)
                {
                    continue;
                }

                var button = CreateLineMarkerButton(link);
                Canvas.SetLeft(button, 2);
                Canvas.SetTop(button, rect.Top + 2);
                MarkerOverlay.Children.Add(button);
            }
        }

        private Button CreateLineMarkerButton(LineLink link)
        {
            var button = new Button
            {
                Content = "🔖",
                Width = 18,
                Height = 18,
                Padding = new Thickness(0),
                Margin = new Thickness(0),
                ToolTip = $"跳转到: {GetNoteDisplayName(link.TargetNotePath)}",
                Tag = link.Id
            };
            button.Click += OnLineMarkerClicked;
            button.ContextMenu = BuildLineMarkerContextMenu(link.Id);
            return button;
        }

        private ContextMenu BuildLineMarkerContextMenu(string linkId)
        {
            var menu = new ContextMenu();
            var deleteItem = new MenuItem { Header = "删除跳转标记" };
            deleteItem.Click += (_, _) => DeleteLineLink(linkId);
            menu.Items.Add(deleteItem);
            return menu;
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
            RemoveLineLinksForNote(relativePath);

            if (_currentNote?.Path == note.Path)
            {
                _currentNote = null;
                EditorBox.Text = string.Empty;
                UpdatePreview();
            }

            LoadNotes(category);
        }

        private void OnRenameNote(object sender, RoutedEventArgs e)
        {
            if (CategoryList.SelectedItem is not string category)
            {
                MessageBox.Show("Please select a category first.");
                return;
            }

            if (NoteList.SelectedItem is not NoteItem note)
            {
                MessageBox.Show("Please select a note to rename.");
                return;
            }

            var newName = Prompt("Enter new note name");
            if (string.IsNullOrWhiteSpace(newName))
            {
                return;
            }

            newName = newName.Trim();
            if (string.Equals(newName, note.Name, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            if (newName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                MessageBox.Show("The note name contains invalid characters.");
                return;
            }

            var newPath = Path.Combine(_rootPath, category, $"{newName}.md");
            if (File.Exists(newPath))
            {
                MessageBox.Show("A note with the same name already exists.");
                return;
            }

            var oldPath = note.Path;
            var oldRelative = GetRelativeNotePath(oldPath);
            var newRelative = GetRelativeNotePath(newPath);

            if (File.Exists(oldPath))
            {
                File.Move(oldPath, newPath);
            }

            if (_currentNote?.Path == oldPath)
            {
                _currentNote = new NoteItem(newName, newPath);
                File.WriteAllText(newPath, EditorBox.Text, Encoding.UTF8);
            }

            _metadata.LineLinks = _metadata.LineLinks
                .Select(link => link.SourceNotePath == oldRelative
                    ? link with { SourceNotePath = newRelative }
                    : link.TargetNotePath == oldRelative
                        ? link with { TargetNotePath = newRelative }
                        : link)
                .ToList();
            SaveMetadata();

            LoadNotes(category);
            if (NoteList.ItemsSource is List<NoteItem> notes)
            {
                var renamed = notes.FirstOrDefault(item => item.Path == newPath);
                if (renamed is not null)
                {
                    NoteList.SelectedItem = renamed;
                }
            }
        }

        private void OnNoteSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (NoteList.SelectedItem is NoteItem note)
            {
                _currentNote = note;
                EditorBox.Text = File.Exists(note.Path) ? File.ReadAllText(note.Path) : string.Empty;
                UpdatePreview();
                RefreshLineMarkers();
            }
        }

        private void OnEditorTextChanged(object sender, TextChangedEventArgs e)
        {
            UpdatePreview();
            RefreshLineMarkers();
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
            PreviewViewer.FontFamily = new System.Windows.Media.FontFamily("Segoe UI Emoji, Segoe UI Symbol, Segoe UI");
            PreviewViewer.Pipeline = _pipeline;
            PreviewViewer.Markdown = ReplaceKeycapDigits(EditorBox.Text);

            _previewScrollViewer ??= FindDescendantScrollViewer(PreviewViewer);
            SyncScroll(_editorScrollViewer, _previewScrollViewer);
        }

        private static string ReplaceKeycapDigits(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            var builder = new StringBuilder(text.Length);
            for (var i = 0; i < text.Length; i++)
            {
                var current = text[i];
                if (current is >= '0' and <= '9')
                {
                    var nextIndex = i + 1;
                    var hasVariant = nextIndex < text.Length && text[nextIndex] == '\uFE0F';
                    var keycapIndex = hasVariant ? nextIndex + 1 : nextIndex;
                    if (keycapIndex < text.Length && text[keycapIndex] == '\u20E3')
                    {
                        builder.Append(current switch
                        {
                            '0' => '⓪',
                            '1' => '①',
                            '2' => '②',
                            '3' => '③',
                            '4' => '④',
                            '5' => '⑤',
                            '6' => '⑥',
                            '7' => '⑦',
                            '8' => '⑧',
                            '9' => '⑨',
                            _ => current
                        });
                        i = keycapIndex;
                        continue;
                    }
                }

                builder.Append(current);
            }

            return builder.ToString();
        }

        private void OnEditorScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (_isSyncingScroll)
            {
                return;
            }

            SyncScroll(_editorScrollViewer, _previewScrollViewer);
            RefreshLineMarkers();
        }

        private void OnPreviewScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (_isSyncingScroll)
            {
                return;
            }

            SyncScroll(_previewScrollViewer, _editorScrollViewer);
        }

        private void SyncScroll(ScrollViewer? source, ScrollViewer? target)
        {
            if (source is null || target is null || target.ScrollableHeight <= 0)
            {
                return;
            }

            var ratio = source.ScrollableHeight <= 0
                ? 0
                : source.VerticalOffset / source.ScrollableHeight;

            _isSyncingScroll = true;
            target.ScrollToVerticalOffset(ratio * target.ScrollableHeight);
            _isSyncingScroll = false;
        }

        private static ScrollViewer? FindDescendantScrollViewer(DependencyObject root)
        {
            for (var i = 0; i < VisualTreeHelper.GetChildrenCount(root); i++)
            {
                var child = VisualTreeHelper.GetChild(root, i);
                if (child is ScrollViewer viewer)
                {
                    return viewer;
                }

                var descendant = FindDescendantScrollViewer(child);
                if (descendant is not null)
                {
                    return descendant;
                }
            }

            return null;
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

        private void OnAddLineLink(object sender, RoutedEventArgs e)
        {
            if (_currentNote is null)
            {
                MessageBox.Show("请先选择笔记。");
                return;
            }

            var options = GetAllNoteOptions();
            if (options.Count == 0)
            {
                MessageBox.Show("暂无可用的目标文档。");
                return;
            }

            var dialog = new NoteSelectionDialog(options) { Owner = this };
            if (dialog.ShowDialog() != true || dialog.SelectedNote is null)
            {
                return;
            }

            var sourceRelative = GetRelativeNotePath(_currentNote.Path);
            var targetRelative = GetRelativeNotePath(dialog.SelectedNote.Path);
            var lineIndex = EditorBox.GetLineIndexFromCharacterIndex(EditorBox.SelectionStart);

            var lineLink = new LineLink(Guid.NewGuid().ToString(), sourceRelative, lineIndex, targetRelative, 0);
            _metadata.LineLinks.Add(lineLink);
            SaveMetadata();
            RefreshLineMarkers();
        }

        private void OnLineMarkerClicked(object sender, RoutedEventArgs e)
        {
            if (sender is not Button button || button.Tag is not string linkId)
            {
                return;
            }

            var link = _metadata.LineLinks.FirstOrDefault(item => item.Id == linkId);
            if (link is null)
            {
                return;
            }

            if (_currentNote is not null)
            {
                _navigationStack.Push(new NavigationEntry(_currentNote.Path, EditorBox.SelectionStart));
            }

            var targetPath = Path.Combine(_rootPath, link.TargetNotePath);
            OpenNoteAt(targetPath, link.TargetPosition);
        }

        private void DeleteLineLink(string linkId)
        {
            var link = _metadata.LineLinks.FirstOrDefault(item => item.Id == linkId);
            if (link is null)
            {
                return;
            }

            _metadata.LineLinks.Remove(link);
            SaveMetadata();
            RefreshLineMarkers();
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
            RefreshLineMarkers();
        }

        private string GetRelativeNotePath(string notePath)
        {
            return Path.GetRelativePath(_rootPath, notePath);
        }

        private void RemoveLineLinksForNote(string relativePath)
        {
            var originalCount = _metadata.LineLinks.Count;
            _metadata.LineLinks = _metadata.LineLinks
                .Where(link => link.SourceNotePath != relativePath && link.TargetNotePath != relativePath)
                .ToList();
            if (_metadata.LineLinks.Count != originalCount)
            {
                SaveMetadata();
            }
        }

        private List<NoteOption> GetAllNoteOptions()
        {
            var options = new List<NoteOption>();
            if (!Directory.Exists(_rootPath))
            {
                return options;
            }

            foreach (var categoryPath in Directory.GetDirectories(_rootPath))
            {
                var categoryName = Path.GetFileName(categoryPath);
                if (string.IsNullOrWhiteSpace(categoryName))
                {
                    continue;
                }

                foreach (var notePath in Directory.GetFiles(categoryPath, "*.md"))
                {
                    var noteName = Path.GetFileNameWithoutExtension(notePath);
                    if (string.IsNullOrWhiteSpace(noteName))
                    {
                        continue;
                    }

                    options.Add(new NoteOption($"{categoryName} / {noteName}", notePath));
                }
            }

            return options.OrderBy(option => option.DisplayName).ToList();
        }

        private string GetNoteDisplayName(string relativeNotePath)
        {
            var trimmed = relativeNotePath.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
            var parts = trimmed.Split(Path.DirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 2)
            {
                var name = Path.GetFileNameWithoutExtension(parts[^1]);
                return $"{parts[^2]} / {name}";
            }

            return Path.GetFileNameWithoutExtension(relativeNotePath);
        }

        private string? Prompt(string prompt)
        {
            var dialog = new InputDialog(prompt) { Owner = this };
            return dialog.ShowDialog() == true ? dialog.Response : null;
        }

        private void InitializeTrayIcon()
        {
            var menu = new ContextMenuStrip();
            var exitItem = new ToolStripMenuItem("退出");
            exitItem.Click += (_, _) => ExitApplication();
            menu.Items.Add(exitItem);

            _trayIcon = new NotifyIcon
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
    }
}
