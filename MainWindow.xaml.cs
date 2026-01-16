using Markdig;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

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

        public MainWindow()
        {
            InitializeComponent();
            _rootPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "MarkdownKnowledgeBase");
            _metadataPath = Path.Combine(_rootPath, ".metadata.json");
            _pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
            EnsureRoot();
            LoadMetadata();
            LoadCategories();
            RefreshMarkersAndLinks();
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

        private void UpdatePreview()
        {
            var html = Markdig.Markdown.ToHtml(EditorBox.Text, _pipeline);
            var page = $@"<html><head><meta charset=""utf-8""><style>body{{font-family:'Segoe UI', sans-serif; padding:16px;}} pre{{background:#f4f4f4; padding:12px;}}</style></head><body>{html}</body></html>";
            PreviewBrowser.NavigateToString(page);
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

        private string? Prompt(string prompt)
        {
            var dialog = new InputDialog(prompt) { Owner = this };
            return dialog.ShowDialog() == true ? dialog.Response : null;
        }
    }
}
