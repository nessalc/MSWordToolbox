using System;
using System.Windows;
using System.Windows.Controls;

namespace Toolbox
{
    /// <summary>
    /// Interaction logic for EnumPrompt.xaml
    /// </summary>
    public partial class EnumPrompt : Window
    {
        public EnumPrompt()
        {
            InitializeComponent();
        }
        private void CancelClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
        private void OKClick(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
        private void DataTypeChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                txtValue.Visibility = Visibility.Hidden;
                dtpValue.Visibility = Visibility.Hidden;
                chkValue.Visibility = Visibility.Hidden;
                dblValue.Visibility = Visibility.Hidden;
                intValue.Visibility = Visibility.Hidden;
                switch (((ComboBoxItem)cboDataType.SelectedItem).Content.ToString())
                {
                    case "Integer":
                        intValue.Visibility = Visibility.Visible;
                        break;
                    case "DateTime":
                        dtpValue.Visibility = Visibility.Visible;
                        break;
                    case "Number":
                        dblValue.Visibility = Visibility.Visible;
                        break;
                    case "Yes/No":
                        chkValue.Visibility = Visibility.Visible;
                        break;
                    case "String":
                    default:
                        txtValue.Visibility = Visibility.Visible;
                        break;
                }
            }
            catch (NullReferenceException)
            {

            }
        }
    }
}
