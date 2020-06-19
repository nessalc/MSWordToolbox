using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using System.Linq;
using System.Globalization;
using StringExtensions;
using System.Windows.Controls;

namespace Toolbox
{
    [ComVisible(true)]
    public partial class Ribbon1 : IRibbonExtensibility
    {
        private IRibbonUI ribbon;
        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Toolbox.Ribbon.xml");
        }
        public void Ribbon1_Load(IRibbonUI ribbonUI)
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
        private const string unknownUnits = @"\p{L}\w*\b";
        private static readonly string units = $@"(?:(?:{SIPrefixedUnits}|{SILog}|{binary}|{additional}|{custom})\b)|%";
        private static readonly string full = $@"({preamble}){whitespace}({quantity}){whitespace}(?:(?:({units})|({unknownUnits})|\b){whitespace})?(?:(?:{pm}){whitespace}({quantity}){whitespace}(?:({units})|({unknownUnits})|\b)|\+{whitespace}({quantity}){whitespace}(?:({units})|({unknownUnits})|\b){whitespace}\/{whitespace}[-−‐-―]{whitespace}({quantity}){whitespace}(?:({units})|({unknownUnits})|\b))?";
        private readonly Regex regexAll = new Regex($@"{full}", RegexOptions.Compiled);
        private readonly Regex regexSingle = new Regex($@"^{full}$", RegexOptions.Compiled);
        private object missing = Type.Missing;
        private List<Range> GetRange(Range range, Regex regex)
        {
            List<Range> ranges = new List<Range>();
            MatchCollection matchCollection = regex.Matches(range.Text);
            if (matchCollection.Count != 0)
            {
                Range testRange = Globals.ThisAddIn.Application.ActiveDocument.Range(range.Start, range.End);
                foreach (Match match in matchCollection)
                {
                    testRange.End = range.Start + match.Index + match.Length;
                    testRange.Start = range.Start + match.Index;
                    int i = 1;
                    while ((testRange.Text != match.Groups[0].Value
                            || (bool)testRange.Information[WdInformation.wdInFieldCode]
                            || (bool)testRange.Information[WdInformation.wdInFieldResult])
                           && i <= range.Fields.Count)
                    {
                        Field field = range.Fields[i];
                        int adjustment = field.ShowCodes ? field.Result.Text.Length + 1 : field.Code.Text.Length + 3;
                        testRange.End += adjustment;
                        testRange.Start += adjustment;
                        i++;
                    }
                    // Okay, so Word is funky with its fields. Who knew? Aside from everyone? I give up. Let's just
                    // move forward a few characters until we find a match...or don't.
                    i = 0;
                    while (testRange.Text != match.Groups[0].Value && i < 5)
                    {
                        testRange.End += 1;
                        testRange.Start += 1;
                        i++;
                    }
                    if (testRange.Text == match.Groups[0].Value
                        && !(bool)testRange.Information[WdInformation.wdInFieldCode]
                        && !(bool)testRange.Information[WdInformation.wdInFieldResult])
                    {
                        Range rangeCopy = Globals.ThisAddIn.Application.ActiveDocument.Range(testRange.Start,
                                                                                             testRange.End);
                        ranges.Add(rangeCopy);
                    }
                }
            }
            return ranges;
        }
        private List<Range> GetRange(Range range, string regex)
        {
            return GetRange(range, new Regex(regex));
        }
        private void REReplace(Range range, Regex regex, string replace)
        {
            List<Range> ranges = GetRange(range, regex);
            foreach (Range rangeIter in ranges)
            {
                rangeIter.Text = regex.Replace(rangeIter.Text, replace);
            }
        }
        private void REReplace(Range range, string regex, string replace)
        {
            REReplace(range, new Regex(regex), replace);
        }
        public void FixQuantities(IRibbonControl control)
        {
#if !DEBUG
            Globals.ThisAddIn.Application.ScreenUpdating = false;
#endif
            IEnumerable<Paragraph> query = from Paragraph paragraph in Globals.ThisAddIn.Application.ActiveDocument.Paragraphs where paragraph.Range.Comments.Count == 0 && paragraph.Range.Bookmarks.Count == 0 select paragraph;
            FormProgress progressForm = new FormProgress();
            ProgressBar progressBar = progressForm.prgProgress;
            progressBar.Maximum = query.Count();
            progressBar.Value = 0;
            progressBar.IsIndeterminate = true;
            progressForm.BringIntoView();
            progressForm.Show();
            progressForm.Activate();
            foreach (Paragraph paragraph in query)
            {
                // get normal font name:
                Font normalFont = Globals.ThisAddIn.Application.ActiveDocument.Styles["Normal"].Font;
                paragraph.Range.Find.Execute("µ", missing, missing, missing, missing, missing, missing, WdFindWrap.wdFindContinue, missing, "μ", WdReplace.wdReplaceAll, missing, missing, missing, missing);
                Find findObject = Globals.ThisAddIn.Application.Selection.Find;
                findObject.ClearFormatting();
                findObject.Text = "\uF06D";
                findObject.Font.Name = "Symbol";
                findObject.Replacement.ClearFormatting();
                findObject.Replacement.Text = "μ";
                findObject.Replacement.Font = normalFont;
                findObject.Execute(missing, missing, missing, missing, missing, missing, missing, WdFindWrap.wdFindContinue, missing, missing, WdReplace.wdReplaceAll, missing, missing, missing, missing);
                paragraph.Range.Find.Execute("Ω", missing, missing, missing, missing, missing, missing, WdFindWrap.wdFindContinue, true, "Ω", WdReplace.wdReplaceAll, missing, missing, missing, missing);
                findObject.ClearFormatting();
                findObject.Text = "\uF057";
                findObject.Font.Name = "Symbol";
                findObject.Replacement.ClearFormatting();
                findObject.Replacement.Text = "Ω";
                findObject.Replacement.Font = normalFont;
                findObject.Execute(missing, missing, missing, missing, missing, missing, missing, WdFindWrap.wdFindContinue, missing, missing, WdReplace.wdReplaceAll, missing, missing, missing, missing);
                REReplace(paragraph.Range, $@"\bu({SIUnits})\b", "μ$1");
                REReplace(paragraph.Range, @"(?<!A-Za-z])(?:([m\u03BC]?)(?:Sec|sec)|([m\u03BC])S)\b", "$1s");
                REReplace(paragraph.Range, $@"(?<!A-Za-z])({SIPrefixes})?(?i)ohm\b", "$1Ω");
                REReplace(paragraph.Range, $@"(\d+{whitespace})[oº]([FC]|{whitespace})", "$1°$2");
                REReplace(paragraph.Range, $@"\bK({SIUnits})\b", "k$1");
                REReplace(paragraph.Range, $@"(?<!\w)({whitespace})[\-−‐-―]{whitespace}(\d)", "$1-$2");
                List<Range> ranges = GetRange(paragraph.Range, regexAll);
                foreach (Range range in ranges)
                {
                    //Fix Units
                    FixQuantityInRange(range);
                }
                progressBar.Value++;
                progressBar.IsIndeterminate = false;
                progressBar.InvalidateVisual();
            }
            progressForm.Hide();
            progressForm.Close();
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        private void FixQuantityInRange(Range range)
        {
            range.Find.Execute("µ", missing, missing, missing, missing, missing, missing, WdFindWrap.wdFindStop, missing, "μ", WdReplace.wdReplaceAll, missing, missing, missing, missing);
            Find findObject = Globals.ThisAddIn.Application.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = "m";
            findObject.Font.Name = "Symbol";
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = "μ";
            findObject.Execute(missing, missing, missing, missing, missing, missing, missing, WdFindWrap.wdFindStop, missing, missing, WdReplace.wdReplaceAll, missing, missing, missing, missing);
            range.Find.Execute("Ω", missing, missing, missing, missing, missing, missing, WdFindWrap.wdFindStop, true, "Ω", WdReplace.wdReplaceAll, missing, missing, missing, missing);
            findObject.ClearFormatting();
            findObject.Text = "W";
            findObject.Font.Name = "Symbol";
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = "Ω";
            findObject.Execute(missing, missing, missing, missing, missing, missing, missing, WdFindWrap.wdFindStop, missing, missing, WdReplace.wdReplaceAll, missing, missing, missing, missing);
            REReplace(range, $@"\bu({SIUnits})\b", "μ$1");
            REReplace(range, @"(?<!A-Za-z])(?:([m\u03BC]?)(?:Sec|sec)|([m\u03BC])S)\b", "$1s");
            REReplace(range, $@"(?<!A-Za-z])({SIPrefixes})?(?i)ohm\b", "$1Ω");
            REReplace(range, $@"(\d+{whitespace})[oº]([FC]|{whitespace})", "$1°$2");
            REReplace(range, $@"\bK({SIUnits})\b", "k$1");
            REReplace(range, $@"(?<!\w)({whitespace})[\-−‐-―]{whitespace}(\d)", "$1-$2");
            range.MoveStartWhile("[({             ​  \r\a\n\t", WdConstants.wdForward);
            range.MoveEndWhile("             ​  \r\a\n\t})]", WdConstants.wdBackward);
            if (range.Text is null)
            {
                return;
            }
            string sel = range.Text.Trim();
            string result = "";
            if (sel.Length != 0 && regexSingle.IsMatch(sel) && range.OMaths.Count == 0)
            {
                Match match = regexSingle.Match(sel);
                string[] unitMatches = new int[] { 3, 4, 6, 7, 9, 10, 12, 13 }.Select(x => match.Groups[x].Value).Where(s => s.Length != 0 && s != null).ToArray();
                bool unitsMatch = unitMatches.All(s => s == unitMatches[0]);
                bool knownUnit = new int[] { 3, 6, 9, 12 }.Select(x => match.Groups[x].Value).Where(s => s.Length != 0 && s != null).Any();
                string units = unitMatches.FirstOrDefault(s => !string.IsNullOrEmpty(s)) ?? "";
                int decimalCount = new int[] { 2, 5, 8, 11 }.Select(x => match.Groups[x].Value)
                                                            .Select(x => x.Length - (x.IndexOf('.') < 0 ? x.Length : x.IndexOf('.') + 1))
                                                            .Max();
                string format = $"F{decimalCount}";
                decimal mainQuantity = decimal.Parse(match.Groups[2].Value.ReplaceMany("−‐‑‒–—―", '-'), CultureInfo.CurrentCulture);
                decimal symTolerance = Math.Abs(decimal.Parse(match.Groups[5].Length != 0 ? match.Groups[5].Value : "0", CultureInfo.CurrentCulture));
                decimal posTolerance = Math.Abs(decimal.Parse(match.Groups[8].Length != 0 ? match.Groups[8].Value : "0", CultureInfo.CurrentCulture));
                decimal negTolerance = Math.Abs(decimal.Parse(match.Groups[11].Length != 0 ? match.Groups[11].Value : "0", CultureInfo.CurrentCulture));
                string main = "";
                string tolerance = "";
                if (match.Groups[5].Length != 0 || (posTolerance == negTolerance && match.Groups[8].Length != 0))
                {
                    if (match.Groups[8].Length != 0)
                    {
                        symTolerance = posTolerance;
                    }
                    tolerance = "± " + symTolerance.ToString(format, CultureInfo.CurrentCulture) + $" {units}";
                }
                else if (posTolerance != negTolerance)
                {
                    tolerance = "+ "
                                + posTolerance.ToString(format, CultureInfo.CurrentCulture)
                                + $" {units}/− "
                                + negTolerance.ToString(format, CultureInfo.CurrentCulture)
                                + $" {units}";
                }
                // Is there one provided unit?
                if (unitsMatch && units.Length != 0)
                {
                    main = mainQuantity.ToString(format, CultureInfo.CurrentCulture).Replace("-", "−") + $" {units}";
                    if (match.Groups[1].Value.Length > 0 && match.Groups[1].Value.IndexOfAny(new char[] { '+', '±' }) > -1)
                    {
                        main = $"{match.Groups[1].Value}{main}";
                    }
                    else if (match.Groups[1].Value.Length != 0)
                    {
                        main = $"{match.Groups[1].Value} {main}";
                    }
                }
                if (unitsMatch && units.Length != 0 && tolerance.Length != 0 && main.Length != 0)
                // include (!knownUnit || knownUnit) in future if settings enable this
                {
                    result = $"{main} {tolerance}";
                }
                else if (unitsMatch && units.Length != 0 && tolerance.Length == 0 && main.Length != 0 && knownUnit)
                {
                    result = main;
                }
                // There is more than one (mismatched) unit
                else if (!unitsMatch && unitMatches.Length != 0)
                {
                    Comment comment = range.Comments.Add(range, "Mismatched Units");
                    comment.Author = "Toolbox";
                    comment.Initial = "TBX";
                }
#if DEBUG
                else if (unitsMatch && units.Length != 0 && !knownUnit && tolerance.Length == 0)
                {
                    // this should be ignored
                    // Dropping in here means that a number is seen with a "unit", but the unit is unknown.
                    //   This can cause numbers followed by "any old word" to get caught up in the unit
                    //   processing code, which we don't want. If, on the other hand, the tolerance exists,
                    //   and the units match, go ahead and process it like it's just an unrecogniezed unit.
                    //   That's caught in a previous if statement. That behavior should be a setting in a
                    //   future build.
                    // this branch exists for debug purposes and to show what is otherwise omitted.
                    ((Action)(() => { }))();
                }
#endif
            }
            if (result.Length != 0)
            {
                range.Text = result;
            }
        }
        public void FixSingleQuantity(IRibbonControl control)
        {
#if !DEBUG
            Globals.ThisAddIn.Application.ScreenUpdating = false;
#endif
            FixQuantityInRange(Globals.ThisAddIn.Application.Selection.Range);
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        public bool GetPageBreakBefore(IRibbonControl control)
        {
            return Globals.ThisAddIn.Application.Selection.ParagraphFormat.PageBreakBefore != 0;
        }
        public bool IsSelection()
        {
            return Globals.ThisAddIn.Application.Selection.End != Globals.ThisAddIn.Application.Selection.Start;
        }
        public bool IsUnit(IRibbonControl control)
        {
            if (IsSelection())
            {
                Range range = Globals.ThisAddIn.Application.Selection.Range;
                range.MoveStartWhile("[({             ​  \r\a\n\t", WdConstants.wdForward);
                range.MoveEndWhile("             ​  \r\a\n\t})]", WdConstants.wdBackward);
                return regexSingle.IsMatch(range.Text);
            }
            return false;
        }
        public void TogglePageBreakBefore(IRibbonControl control, bool pressed)
        {
            ParagraphFormat format = Globals.ThisAddIn.Application.Selection.ParagraphFormat;
            if (format.PageBreakBefore == 0)
                format.PageBreakBefore = -1;
            else
                format.PageBreakBefore = 0;
            InvalidateToggle();
        }
        public void InvalidateToggle()
        {
            ribbon.InvalidateControl("tglTogglePgBrk");
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
                result = result.Replace(c, newValue);
            }
            return result;
        }
        public static string ReplaceMany(this string input, string oldValues, char newValue)
        {
            return ReplaceMany(input, oldValues.ToArray(), newValue);
        }
    }
}