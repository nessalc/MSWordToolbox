using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using System.Linq;
using System.Globalization;
using StringExtensions;

namespace Toolbox
{
    [ComVisible(true)]
    public partial class Ribbon1: IRibbonExtensibility
    {
        private IRibbonUI ribbon;
        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Toolbox.Ribbon.xml");
        }
        private void Ribbon1_Load(IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
        }
        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
        private const string quantity = @"[-−‐-―]?(?:\d+\.\d+|\d+|\.\d+|\d+\.)";
        private const string whitespace = @"[ \t  -   ]*";
        private const string SIPrefixes = @"da|[YZEPTGMkhdcmμnpfazy]";
        private const string SIUnits = @"Wb|Sv|Hz|sr|mol|lm|lx|cd|rad|Pa|Bq|Da|eV|ua|Gy|kat|°C|[gmsulAKNJWCVFSTHLΩ]";
        private static readonly string SIPrefixedUnits = $"(?:{SIPrefixes})?(?:{SIUnits})";
        private const string SILog = @"m?Np|B|dB(?:FS|iC|m0s?|mV|ov|pp|rnC|sm|TP|μV|μ0s|VU| HL| Q| SIL| SPL| SWL|\/K|-Hz|[ABcCdefGiJkKmoOqruvVWZμ])?";
        private const string binary = @"[KMGTPEZY]i[bB]";
        private const string additional = @"[Mkdcm]?bar|mmHg|ha|min|[Åbhdt]";
        private const string custom = @"kts?|ft|lbs?|inHg|n?mi|psi|atm|°F|VDC";
        private const string pm = @"±|\+\/[-−‐-―]";
        private const string preamble = @"[≤≥<>\+±]?";
        private static readonly string units = $@"(?:(?:{SIPrefixedUnits}|{SILog}|{binary}|{additional}|{custom})\b)|%";
        private static readonly string full = $@"({preamble}){whitespace}({quantity}){whitespace}(?:(?:({units})|(\w+\b)){whitespace})?(?:(?:{pm}){whitespace}({quantity}){whitespace}(?:({units})|(\w+\b))|\+{whitespace}({quantity}){whitespace}(?:({units})|(\w+\b)){whitespace}\/{whitespace}[-−‐-―]{whitespace}({quantity}){whitespace}(?:({units})|(\w+\b)))?";
        private readonly Regex regex = new Regex($@"{full}", RegexOptions.Compiled);
        public void FixQuantity(IRibbonControl control)
        {
#if !DEBUG
            Globals.ThisAddIn.Application.ScreenUpdating = false;
#endif
            Selection selection = Globals.ThisAddIn.Application.Selection;
            string sel = selection.Text.Trim();
            string result = "";
            if (sel.Length != 0 && regex.IsMatch(sel) && selection.Range.OMaths.Count == 0)
            {
                Match match = regex.Match(sel);
                string[] unitMatches = new int[] { 3, 4, 6, 7, 9, 10, 12, 13 }.Select(x => match.Groups[x].Value).Where(s => s.Length != 0 && s != null).ToArray();
                bool unitsMatch = unitMatches.All(s => s == unitMatches[0]);
                string units = unitMatches.FirstOrDefault(s => !string.IsNullOrEmpty(s)) ?? "";
                int decimalCount = new int[] { 2, 5, 8, 11 }.Select(x => match.Groups[x].Value).Select(x => x.Length - (x.IndexOf('.') < 0 ? x.Length : x.IndexOf('.') + 1)).Max();
                string format = $"F{decimalCount}";
                decimal mainQuantity = decimal.Parse(match.Groups[2].Value.ReplaceMany("−‐‑‒–—―", '-'), CultureInfo.CurrentCulture);
                decimal symTolerance = Math.Abs(decimal.Parse(match.Groups[5].Length != 0 ? match.Groups[5].Value : "0", CultureInfo.CurrentCulture));
                decimal posTolerance = Math.Abs(decimal.Parse(match.Groups[8].Length != 0 ? match.Groups[8].Value : "0", CultureInfo.CurrentCulture));
                decimal negTolerance = Math.Abs(decimal.Parse(match.Groups[11].Length != 0 ? match.Groups[11].Value : "0", CultureInfo.CurrentCulture));
                if (match.Groups[2].Length != 0 || (posTolerance == negTolerance && match.Groups[5].Length != 0))
                {
                    if (match.Groups[5].Length != 0)
                    {
                        symTolerance = posTolerance;
                    }
                    string tolerance = $"± {symTolerance} {units}";
                } else if (posTolerance != negTolerance)
                {
                    string tolerance = $"+ {posTolerance} {units}/− {negTolerance} {units}";
                }
                // Is there one provided unit?
                if (unitsMatch && units.Length != 0)
                {

                }
                // There is more than one (mismatched) unit
                else if (units.Length!=0)
                {

                }
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
    }
}

namespace StringExtensions
{
    public static class StringExtensionClass
    {
        public static string ReplaceMany(this string input, char[] oldValues, char newValue)
        {
            string result = input;
            foreach (char c in oldValues)
            {
                result.Replace(c, newValue);
            }
            return result;
        }
        public static string ReplaceMany(this string input, string oldValues, char newValue)
        {
            return ReplaceMany(input, oldValues.ToArray(), newValue);
        }
    }
}