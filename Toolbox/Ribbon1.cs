using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

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
        private const string quantity = @"[\-−–]?(?:\d+\.\d+|\d+|\.\d+|\d+\.)";
        private const string whitespace = @"[ \t  -   ]*";
        private const string SIPrefixes = @"da|[YZEPTGMkhdcmμnpfazy]";
        private const string SIUnits = @"Wb|Sv|Hz|sr|mol|lm|lx|cd|rad|Pa|Bq|Da|eV|ua|Gy|kat|°C|[gmsulAKNJWCVFSTHLΩ]";
        private static readonly string SIPrefixedUnits = $"(?:{SIPrefixes})?(?:{SIUnits})";
        private const string SILog = @"m?Np|B|dB(?:FS|iC|m0s?|mV|ov|pp|rnC|sm|TP|μV|μ0s|VU| HL| Q| SIL| SPL| SWL|\/K|-Hz|[ABcCdefGiJkKmoOqruvVWZμ])?";
        private const string binary = @"[KMGTPEZY]i[bB]";
        private const string additional = @"[Mkdcm]?bar|mmHg|ha|min|[Åbhdt]";
        private const string custom = @"kts?|ft|lbs?|inHg|n?mi|psi|atm|°F|VDC";
        private const string pm = @"±|\+\/[-−]";
        private const string preamble = @"[≤≥<>\+\-−±]?";
        private static readonly string units = $@"(?:(?:(?:{SIPrefixedUnits}|(?:{SILog})|(?:{binary})|(?:{additional})|(?:{custom}))\b)|%)";
        private static readonly string full = $@"((?:{preamble}))(?:{whitespace})((?:{quantity}))(?:{whitespace})(?:((?:{units}))?|(\w+\b))?(?:(?:{whitespace})(?:{pm})(?:{whitespace})((?:{quantity}))(?:{whitespace})(?:((?:{units}))?|(\w+\b))?|\+(?:{whitespace})((?:{quantity}))(?:{whitespace})(?:((?:{units}))?|(\w+\b))?(?:{whitespace})\/(?:{whitespace})[-−](?:{whitespace})((?:{quantity}))(?:{whitespace})(?:((?:{units}))?|(\w+\b))?)?";
        private readonly Regex regex = new Regex($@"{full}", RegexOptions.Compiled);
        public void FixQuantity(IRibbonControl control)
        {
#if !DEBUG
            Globals.ThisAddIn.Application.ScreenUpdating = false;
#endif
            Selection selection = Globals.ThisAddIn.Application.Selection;
            string s = selection.Text.Trim();
            if (s.Length!=0)
            {
                MatchCollection matchCollection = regex.Matches(selection.Text.Trim());
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
    }
}
