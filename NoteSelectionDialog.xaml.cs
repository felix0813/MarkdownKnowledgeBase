using System.Collections.Generic;
using System.Linq;
using System.Windows;
using MessageBox = System.Windows.MessageBox;

namespace MarkdownKnowledgeBase
{
    public partial class NoteSelectionDialog : Window
    {
        private readonly List<NoteOption> _options;

        public NoteSelectionDialog(List<NoteOption> options)
        {
            InitializeComponent();
            _options = options;
            NoteList.ItemsSource = _options;
            NoteList.SelectedItem = _options.FirstOrDefault();
        }

        public NoteOption? SelectedNote => NoteList.SelectedItem as NoteOption;

        private void OnOkClicked(object sender, RoutedEventArgs e)
        {
            if (SelectedNote is null)
            {
                MessageBox.Show("请选择目标文档。");
                return;
            }

            DialogResult = true;
        }

        private void OnCancelClicked(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
