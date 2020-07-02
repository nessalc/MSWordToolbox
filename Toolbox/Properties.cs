using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System.Globalization;
using Toolbox.Properties;

namespace Toolbox
{
    class PropertyFunctions
    {
        public static void SetDocumentValue(string PropOrVariableName, object value)
        {
            if (Settings.Default.PreferDocumentVariable)
            {
                Variables variables = Globals.ThisAddIn.Application.ActiveDocument.Variables;
                variables[PropOrVariableName].Value = value.ToString();
            }
            else
            {
                DocumentProperties properties;
                try
                {
                    properties = (DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties;
                    properties[PropOrVariableName].Value = Convert.ChangeType(value, properties[PropOrVariableName].GetType(), CultureInfo.CurrentCulture);
                }
                catch (ArgumentException)
                {
                    try
                    {
                        properties = (DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties;
                        object testval = properties[PropOrVariableName].Value;
                        if (testval is string)
                        {
                            properties[PropOrVariableName].Value = value.ToString();
                        }
                        else if (testval is bool)
                        {
                            properties[PropOrVariableName].Value = bool.Parse(value.ToString());
                        }
                        else if (testval is int)
                        {
                            properties[PropOrVariableName].Value = int.Parse(value.ToString(), CultureInfo.CurrentCulture);
                        }
                        else if (testval is double)
                        {
                            properties[PropOrVariableName].Value = float.Parse(value.ToString(), CultureInfo.CurrentCulture);
                        }
                        else if (testval is DateTime)
                        {
                            properties[PropOrVariableName].Value = DateTime.Parse(value.ToString(), CultureInfo.CurrentCulture);
                        }
                    }
                    catch (ArgumentException)
                    {
                        switch (value.GetType().Name)
                        {
                            case "String":
                                ((DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties).Add(PropOrVariableName,
                                    false, MsoDocProperties.msoPropertyTypeString, value);
                                break;
                            case "SByte":
                            case "Int16":
                            case "UInt16":
                            case "Int32":
                            case "UInt32":
                            case "Int64":
                            case "UInt64":
                                ((DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties).Add(PropOrVariableName,
                                    false, MsoDocProperties.msoPropertyTypeNumber, value);
                                break;
                            case "Single":
                            case "Double":
                            case "Decimal":
                                ((DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties).Add(PropOrVariableName,
                                    false, MsoDocProperties.msoPropertyTypeFloat, value);
                                break;
                            case "Boolean":
                                ((DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties).Add(PropOrVariableName,
                                    false, MsoDocProperties.msoPropertyTypeBoolean, value);
                                break;
                            case "DateTime":
                                ((DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties).Add(PropOrVariableName,
                                    false, MsoDocProperties.msoPropertyTypeDate, value);
                                break;
                        }
                    }
                }
            }
        }
        public static object GetDocumentValue(string PropOrVariableName, object DefaultValue = null)
        {
            object retval;
            try
            {
                Variables variables = Globals.ThisAddIn.Application.ActiveDocument.Variables;
                retval = variables[PropOrVariableName].Value;
            }
            catch (ArgumentException)
            {
                DocumentProperties properties;
                try
                {
                    properties = (DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties;
                    retval = properties[PropOrVariableName].Value;
                }
                catch (ArgumentException)
                {
                    try
                    {
                        properties = (DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties;
                        retval = properties[PropOrVariableName].Value;
                    }
                    catch (ArgumentException)
                    {
                        retval = DefaultValue;
                    }
                }
            }
            return retval;
        }
    }
}