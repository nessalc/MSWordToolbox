using System;
using System.Globalization;
using System.Windows;
using Toolbox.Properties;

namespace Toolbox
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class SettingsDialog : Window
    {
        public SettingsDialog()
        {
            InitializeComponent();
            dgCharacters.ItemsSource = Settings.Default.SpecialCharacters;
        }

        private void cbPreferDocVars_Click(object sender, RoutedEventArgs e)
        {
            Settings.Default.PreferDocumentVariables = (bool)cbPreferDocVars.IsChecked;
            Settings.Default.Save();
        }

        private void cbIncludeBookmarks_Click(object sender, RoutedEventArgs e)
        {
            Settings.Default.IncludeBookmarks = (bool)cbIncludeBookmarks.IsChecked;
            Settings.Default.Save();
        }

        private void cbIncludeVariables_Click(object sender, RoutedEventArgs e)
        {
            Settings.Default.IncludeVariables = (bool)cbIncludeVariables.IsChecked;
            Settings.Default.Save();
        }

        private void txtTextSize_LostFocus(object sender, RoutedEventArgs e)
        {
            Settings.Default.FontSize = (int)txtTextSize.Value;
            Settings.Default.Save();
        }

        private void txtIconSize_LostFocus(object sender, RoutedEventArgs e)
        {
            Settings.Default.IconSize = (int)txtIconSize.Value;
            Settings.Default.Save();
        }

        private void txtCustomUnits_LostFocus(object sender, RoutedEventArgs e)
        {
            Settings.Default.CustomUnits = txtCustomUnits.Text;
            Settings.Default.Save();
        }
    }
}
