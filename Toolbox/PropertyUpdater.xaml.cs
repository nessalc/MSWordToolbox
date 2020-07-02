using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Office.Core;
using Toolbox.Properties;
using Interop = Microsoft.Office.Interop.Word;

namespace Toolbox
{
    /// <summary>
    /// Interaction logic for PropertyUpdater.xaml
    /// </summary>
    public partial class PropertyUpdater : Window
    {
        public PropertyUpdater()
        {
            InitializeComponent();
        }
        private void SaveClick(object sender, RoutedEventArgs e)
        {
            string selectedType = ((ListBoxItem)itemType.SelectedItem).Content.ToString();
            string selectedAttribute = (string)itemList.SelectedItem;
            switch (selectedType)
            {
                case "Document Property":
                    DocumentProperties properties;
                    try
                    {
                        properties = (DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties;
                        if (txtContents.IsVisible)
                            properties[selectedAttribute].Value = txtContents.Text;
                        else if (dtpDateTime.IsVisible)
                            properties[selectedAttribute].Value = dtpDateTime.Value;
                        else if (intValue.IsVisible)
                            properties[selectedAttribute].Value = intValue.Value;
                        else if (dblValue.IsVisible)
                            properties[selectedAttribute].Value = dblValue.Value;
                        else if (chkValue.IsVisible)
                            properties[selectedAttribute].Value = chkValue.IsChecked;
                    }
                    catch (System.ArgumentException)
                    {
                        properties = (DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties;
                        if (txtContents.IsVisible)
                            properties[selectedAttribute].Value = txtContents.Text;
                        else if (dtpDateTime.IsVisible)
                            properties[selectedAttribute].Value = dtpDateTime.Value;
                        else if (intValue.IsVisible)
                            properties[selectedAttribute].Value = intValue.Value;
                        else if (dblValue.IsVisible)
                            properties[selectedAttribute].Value = dblValue.Value;
                        else if (chkValue.IsVisible)
                            properties[selectedAttribute].Value = chkValue.IsChecked;
                    }
                    break;
                case "Document Variable":
                    Interop.Variables variables = Globals.ThisAddIn.Application.ActiveDocument.Variables;
                    variables[selectedAttribute].Value = txtContents.Text;
                    break;
                case "Bookmark":
                    break;
            }
        }
        private void CloseClick(object sender, RoutedEventArgs e)
        {
            Close();
        }
        private void ChooseType(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                string selectedType = itemType.SelectedItem.ToString();
                itemList.Items.Clear();
                switch (selectedType)
                {
                    case "Document Property":
                        foreach (DocumentProperty property in (DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties)
                        {
                            itemList.Items.Add(property.Name);
                        }
                        foreach (DocumentProperty property in (DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties)
                        {
                            itemList.Items.Add(property.Name);
                        }
                        break;
                    case "Document Variable":
                        foreach (Interop.Variable variable in Globals.ThisAddIn.Application.ActiveDocument.Variables)
                        {
                            itemList.Items.Add(variable.Name);
                        }
                        break;
                    case "Bookmark":
                        foreach (Interop.Bookmark bookmark in Globals.ThisAddIn.Application.ActiveDocument.Bookmarks)
                        {
                            itemList.Items.Add(bookmark.Name);
                        }
                        break;
                }
                itemList.Items.SortDescriptions.Add(new System.ComponentModel.SortDescription("", System.ComponentModel.ListSortDirection.Ascending));
            }
            catch (NullReferenceException)
            {

            }
        }
        private void ChooseAttribute(object sender, SelectionChangedEventArgs e)
        {
            if (itemList.SelectedItem != null)
            {
                btnSave.IsEnabled = true;
                string selectedItem = itemList.SelectedItem.ToString();
                string category = itemType.SelectedItem.ToString();
                object retval;
                if (category == "Document Property")
                {
                    try
                    {
                        DocumentProperties properties = (DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties;
                        retval = properties[selectedItem];
                    }
                    catch (ArgumentException)
                    {
                        DocumentProperties properties = (DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties;
                        retval = properties[selectedItem];
                    }
                    MsoDocProperties valueType = ((DocumentProperty)retval).Type;
                    txtContents.Visibility = Visibility.Hidden;
                    dtpDateTime.Visibility = Visibility.Hidden;
                    chkValue.Visibility = Visibility.Hidden;
                    dblValue.Visibility = Visibility.Hidden;
                    intValue.Visibility = Visibility.Hidden;
                    try
                    {
                        object test = ((DocumentProperty)retval).Value;
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        valueType = 0;
                    }
                    switch (valueType)
                    {
                        case MsoDocProperties.msoPropertyTypeString:
                            txtContents.Visibility = Visibility.Visible;
                            txtContents.Text = (string)((DocumentProperty)retval).Value;
                            break;
                        case MsoDocProperties.msoPropertyTypeDate:
                            dtpDateTime.Visibility = Visibility.Visible;
                            dtpDateTime.Value = (DateTime)((DocumentProperty)retval).Value;
                            break;
                        case MsoDocProperties.msoPropertyTypeBoolean:
                            chkValue.Visibility = Visibility.Visible;
                            chkValue.IsChecked = (bool)((DocumentProperty)retval).Value;
                            break;
                        case MsoDocProperties.msoPropertyTypeFloat:
                            dblValue.Visibility = Visibility.Visible;
                            dblValue.Value = (float)((DocumentProperty)retval).Value;
                            break;
                        case MsoDocProperties.msoPropertyTypeNumber:
                            intValue.Visibility = Visibility.Visible;
                            intValue.Value = (int)((DocumentProperty)retval).Value;
                            break;
                        case 0:
                            txtContents.Visibility = Visibility.Visible;
                            txtContents.Text = "";
                            break;
                    }
                }
                else if (category == "Document Variable")
                {
                    retval = Globals.ThisAddIn.Application.ActiveDocument.Variables[selectedItem].Value;
                    txtContents.Text = (string)retval;
                }
                else if (category == "Bookmark")
                {

                }
            }
            else
            {
                btnSave.IsEnabled = false;
            }
        }
        private void DeleteClick(object sender, RoutedEventArgs e)
        {
            string selectedItem = itemList.SelectedItem.ToString();
            string[] builtins = Array.Empty<string>();
            foreach (DocumentProperty documentProperty in (DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties)
            {
                builtins.Append(documentProperty.Name);
            }
            if (!builtins.Contains(selectedItem))
            {
                (((DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties)[itemList.SelectedItem.ToString()]).Delete();
            }
            else
            {
                MessageBox.Show("Cannot delete built-in document property!", "Error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            itemList.Items.Remove(selectedItem);
            txtContents.Text = "";
            dtpDateTime.Value = null;
            intValue.Value = null;
            dblValue.Value = null;
            chkValue.IsChecked = false;
        }
        private void NewClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if ((ListBoxItem)itemType.SelectedItem != null)
                {
                    string selectedType = ((ListBoxItem)itemType.SelectedItem).Content.ToString();
                    EnumPrompt enumPrompt = new EnumPrompt
                    {
                        FontSize = Settings.Default.FontSize
                    };
                    if (selectedType == "Document Variable")
                    {
                        enumPrompt.cboDataType.IsEnabled = false;
                    }
                    enumPrompt.Owner = this;
                    enumPrompt.ShowDialog();
                    if (itemList.Items.Contains(enumPrompt.txtName.Text))
                    {
                        MessageBox.Show(selectedType + " already exists! Cannot duplicate " + selectedType + " name!", "Error!", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                    else
                    {
                        switch (selectedType)
                        {
                            case "Document Property":
                                MsoDocProperties newValueType;
                                object newValue;
                                switch (enumPrompt.cboDataType.Text)
                                {
                                    case "Integer":
                                        newValueType = MsoDocProperties.msoPropertyTypeNumber;
                                        newValue = enumPrompt.intValue.Value;
                                        break;
                                    case "Yes/No":
                                        newValueType = MsoDocProperties.msoPropertyTypeBoolean;
                                        newValue = enumPrompt.chkValue.IsChecked;
                                        break;
                                    case "Number":
                                        newValueType = MsoDocProperties.msoPropertyTypeFloat;
                                        newValue = enumPrompt.dblValue.Value;
                                        break;
                                    case "DateTime":
                                        newValueType = MsoDocProperties.msoPropertyTypeDate;
                                        newValue = enumPrompt.dtpValue.Value;
                                        break;
                                    case "String":
                                    default:
                                        newValueType = MsoDocProperties.msoPropertyTypeString;
                                        newValue = enumPrompt.txtValue.Text;
                                        break;
                                }
                                ((DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties).Add(enumPrompt.txtName.Text, false, newValueType, newValue);
                                itemList.Items.Add(enumPrompt.txtName.Text);
                                break;
                            case "Document Variable":
                                Globals.ThisAddIn.Application.ActiveDocument.Variables.Add(enumPrompt.txtName.Text, enumPrompt.txtValue.Text);
                                itemList.Items.Add(enumPrompt.txtName.Text);
                                break;
                            case "Bookmark":
                                break;
                        }
                        itemList.Items.SortDescriptions.Add(new System.ComponentModel.SortDescription("", System.ComponentModel.ListSortDirection.Ascending));
                    }
                }
                else
                {
                    MessageBox.Show("Please select a property type.", "Error!", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {

            }
        }
    }
}
