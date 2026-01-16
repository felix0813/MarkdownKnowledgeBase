using System.Windows;

namespace MarkdownKnowledgeBase
{
    public partial class InputDialog : Window
    {
        public InputDialog(string prompt)
        {
            InitializeComponent();
            PromptText.Text = prompt;
            InputBox.Focus();
        }

        public string? Response => InputBox.Text;

        private void OnOkClicked(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void OnCancelClicked(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
