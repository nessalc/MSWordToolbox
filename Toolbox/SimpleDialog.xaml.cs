using System.Windows;

namespace Toolbox
{
    /// <summary>
    /// Interaction logic for SimpleDialog.xaml
    /// </summary>
    public partial class SimpleDialog : Window
    {
        public SimpleDialog(string prompt, string option1, string option2, string title, bool includeCancel = true)
        {
            InitializeComponent();
            btnOpt1.Content = option1;
            btnOpt2.Content = option2;
            txtPrompt.Content = prompt;
            Title = title;
            if (!includeCancel)
            {
                grdMain.ColumnDefinitions.RemoveAt(2);
                btnCancel.Visibility = Visibility.Hidden;
            }
            Result = null;
        }
        private void Option1Click(object sender, RoutedEventArgs e)
        {
            Result = btnOpt1.Content.ToString();
            DialogResult = true;
            Close();
        }
        private void Option2Click(object sender, RoutedEventArgs e)
        {
            Result = btnOpt2.Content.ToString();
            DialogResult = true;
            Close();
        }
        private void CancelClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
        public string Result { get; set; }
    }
}
