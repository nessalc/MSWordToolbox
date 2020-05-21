using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Controls;

namespace Toolbox
{
    public partial class ToolboxRibbon
    {
        private void Toolbox_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnFixQuantity_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            int total = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs.Count;
            FormProgress progressForm = new FormProgress();
            progressForm.BringIntoView();
            progressForm.Show();
            progressForm.Activate();
            ProgressBar progressBar = progressForm.prgProgress;
            progressBar.Maximum = total;
            int idx = 0;
            progressBar.Value = idx;
            Find find = Globals.ThisAddIn.Application.ActiveDocument.Content.Find;
            object missing = Type.Missing;
            find.Execute("\u2126", missing, missing, missing, missing, missing, missing, missing, missing, "\u03A9",
                    WdReplace.wdReplaceAll, missing, missing, missing, missing);
            string quantityExp = @"(?:\d+\.\d+|\d+|\.\d+|\d+\.)";
            string spaceExp = "[ \\t\xA0\u2000-\u200A\u202F\u205F]";
            //All SI units with prefixes as well as derived units--capital K will be replaced
            string unitSIExp = "(?:(?:[YZEPTGMkhdcmµμnpfazy]|da)?\u03A9|(?:[YZEPTGMKkhdcm\xB5\u03BCnpfazy]|da)?(?:Wb|Sv|Hz|sr|mol|lm|lx|cd|rad|Pa|Wb|Bq|Da|eV|ua|Gy|kat|°C|[gmsulAKNJWCVFSTHL])\\b)";
            string unitOtherNISTExp = "(?:\xC5|b|[Mkdcm]?bar|mmHg|ha|min|[hdt])\\b";
            //Some traditional (ICAO) units
            string unitOtherICAOExp = "(?:kts|ft|lbs?|inHg|n?mi|psi|mmHg|atm|°F|VDC)\\b";
            //Scalars
            string unitScalarsExp = "(?:°|%|B\\b|dB(?:[ABcCdefGiJkKmoOqruvVWZ\xB5\u03BC]|FS|iC|m0s?|mV|ov|pp|rnC|sm|TP|u0s|uV|[\xB5\u03BC]V|VU| HL| Q| SIL| SPL| SWL|/K|-Hz)?\\b|m?Np\\b)";
            //Binary units
            string unitBinExp = "(?:[KMGTPEZY]i[bB])\\b";
            //All possible combinations using Latin/Greek letters
            //string unitAllExp = $"[A-Za-z\xB5\u0391-\u03a9\u03B1-\u03C9°]+";
            //string unitExp = $"{unitSIExp}"; //SI only
            //string unitExp = $"{unitBinExp}|{unitSIExp}"; //binary plus SI
            //string unitExp = $"{unitOtherICAOExp}|{unitOtherNISTExp}|{unitSIExp}"; //some traditional, some SI
            string unitExp = $"{unitOtherICAOExp}|{unitBinExp}|{unitSIExp}|{unitScalarsExp}|{unitOtherNISTExp}";
            //string unitExp = $"[A-Za-z\xB5\u0391-\u03a9\u03B1-\u03C9]+"; //accept basically any (Latin/Greek letter combination) as a unit
            string pmExp = "(?:±|\\+\\/[-\u2212])";
            string inequalityExp = "[±+\\-<>\u2264\u2265]";
            string toleranceExp = $"(?:{pmExp}{spaceExp}*({quantityExp}){spaceExp}*(?:({unitExp}))?|\\+({quantityExp}){spaceExp}*(?:({unitExp}))?\\/[-\u2212]({quantityExp}){spaceExp}*(?:({unitExp}))?)";
            string full = $"(?:({inequalityExp}){spaceExp}*)?([-\u2212]?{quantityExp}){spaceExp}*(?:({unitExp}))?{spaceExp}*{toleranceExp}?";
            Regex regex = new Regex($"(?<=[[\\s\\(]){full}");
            Regex regexb = new Regex($"^{full}");
            Range rng;
            System.Collections.Generic.IEnumerable<Paragraph> query = from Paragraph paragraph in Globals.ThisAddIn.Application.ActiveDocument.Paragraphs where paragraph.Range.Comments.Count == 0 && paragraph.Range.Bookmarks.Count == 0 select paragraph;
            foreach (Paragraph paragraph in query)
            {
                //First, clean up common mistakes:
                //millisecond: mS, msec, mSec [ensure previous character is not a letter] ==> ms
                Regex badUnit = new Regex(@"(?<![A-Za-z])m(?:[Ss](?:ec)|S)\b");
                foreach (Match match in badUnit.Matches(paragraph.Range.Text))
                {
                    rng = paragraph.Range;
                    ReplaceText(rng, match.Index, match.Groups[0].Value, "ms");
                }
                //second: sec, second, seconds [ensure previous character is not a letter] ==> s
                badUnit = new Regex(@"(?<![A-Za-z])sec(?:ond)?s?\b");
                foreach (Match match in badUnit.Matches(paragraph.Range.Text))
                {
                    rng = paragraph.Range;
                    ReplaceText(rng, match.Index, match.Groups[0].Value, "s");
                }
                badUnit = new Regex(@"(\bus\b|\busec\b|\buS\b|\buSec\b)");
                foreach (Match match in badUnit.Matches(paragraph.Range.Text))
                {
                    rng = paragraph.Range;
                    ReplaceText(rng, match.Index, match.Groups[0].Value, "\u03BCs");
                }
                badUnit = new Regex(@"(ohm\b)", RegexOptions.IgnoreCase);
                foreach (Match match in badUnit.Matches(paragraph.Range.Text))
                {
                    rng = paragraph.Range;
                    ReplaceText(rng, match.Index, match.Groups[0].Value, "\u03A9");
                }
                //perform actual fixes
                MatchCollection matchCollection = regex.Matches(paragraph.Range.Text);
                foreach (Match match in matchCollection)
                {
                    FixUnits(match, paragraph);
                }
                //do it again for stuff inside tables, lists
                matchCollection = regexb.Matches(paragraph.Range.Text);
                foreach (Match match in matchCollection)
                {
                    FixUnits(match, paragraph);
                }
                progressBar.Value = idx++;
            }
            progressForm.Hide();
            progressForm.Close();
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        private void FixUnits(Match match, Paragraph paragraph)
        {
            /*
             * match.Groups:
             * 0	Entire match
             * 1	Inequality
             * 2	Main quantity
             * 3	Main units
             * 4	Tolerance quantity
             * 5	Tolerance units
             * 6	Positive tolerance quantity
             * 7	Positive tolerance units
             * 8	Negative tolerance quantity
             * 9	Negative tolerance units
             * */
            Globals.ThisAddIn.Application.ActiveDocument.RemovePersonalInformation = false;
            Range rng;
            // skip if no units
            if (match.Groups[0].Value.Trim() == match.Groups[2].Value ||
                (match.Groups[1].Value.IndexOf("±+-", StringComparison.CurrentCulture) != -1 &&
                 !match.Groups.Cast<string>().SkipWhile((x, i) => i <= 2).Any(x => x.Length == 0)) ||
                 // skip if an equation
                 paragraph.Range.OMaths.Count != 0) //no units
            {
                return;
            }
            string units = "X";
            string[] unitMatches = new string[] { match.Groups[3].Value,
                match.Groups[5].Value,
                match.Groups[7].Value,
                match.Groups[9].Value }.Where(s => (s.Length != 0 && s != null)).ToArray();
            int decimalCount = (new string[] { match.Groups[2].Value,
                match.Groups[4].Value,
                match.Groups[6].Value,
                match.Groups[8].Value }).Select(x => x.Length - (x.IndexOf('.') < 0 ? x.Length : x.IndexOf('.') + 1)).Max();
            string format = "F" + decimalCount.ToString(CultureInfo.CurrentCulture);
            //Make sure all the units detected are the same
            //If they are, and there's more than one of them, use that
            if (unitMatches.Length > 0 && unitMatches.All(s => s == unitMatches[0]))
            {
                units = unitMatches.FirstOrDefault(s => !string.IsNullOrEmpty(s)) ?? "";
            }
            //If they're not, and there's more than one of them, flag it
            else if (unitMatches.Length > 1)
            {
                rng = paragraph.Range;
                int start = match.Index;
                int length = match.Groups[0].Length;
                SimpleDialog dialog = new SimpleDialog($"Quantity\n\n{match.Groups[0]}\n\ndoes not have matching units.\n\nContext: \"{SafeSubstring(paragraph.Range.Text, start - 10, length + 20).Trim()}\"",
                    "Fix", "Skip", "Invalid Units", false);
                dialog.ShowDialog();
                if (dialog.Result == "Skip")
                {
                    SetComment(rng, "The units don't match!", start, length);
                    return;
                }
                else
                {
                    string response = Microsoft.VisualBasic.Interaction.InputBox("Enter correct units", "Fix Units", unitMatches[0]);
                    units = response;
                }
                SetComment(rng, "The units don't match!", start, length);
                return;
            }
            //If they're not (because there are no units), flag it
            else if (unitMatches.Length == 0)
            {
                rng = paragraph.Range;
                int start = match.Index;
                int length = match.Groups[0].Length;
                SimpleDialog dialog = new SimpleDialog($"Quantity\n\n{match.Groups[0]}\n\ndoes not have any units.\n\nContext: \"{SafeSubstring(paragraph.Range.Text, start - 10, length + 20).Trim()}\"",
                    "Fix", "Skip", "Invalid Units", false);
                dialog.ShowDialog();
                if (dialog.Result == "Skip")
                {
                    SetComment(rng, "This doesn't have any units!", start, length);
                    return;
                }
                else
                {
                    string response = Microsoft.VisualBasic.Interaction.InputBox("Enter correct units", "Fix Units");
                    units = response;
                }
            }
            //fix units that begin with a capital K--should be lowercase for the "kilo" prefix
            if (units.First() == 'K')
            {
                units = "k" + units.Substring(1);
            }
            //Get main quantity
            decimal quantity = decimal.Parse(match.Groups[2].Value.Replace('\u2212', '-'), CultureInfo.CurrentCulture);
            if (match.Groups[4].Value.Length != 0 || (match.Groups[6].Value == match.Groups[8].Value && match.Groups[6].Value.Length != 0))
            {
                string prefix = match.Groups[1].Value + (match.Groups[1].Value.Length == 0 ? "" : "\u2009");
                decimal tolerance = decimal.Parse(match.Groups[match.Groups[4].Value.Length != 0 ? 4 : 6].Value.Replace('\u2212', '-'), CultureInfo.CurrentCulture);
                tolerance = Math.Abs(tolerance);
                rng = paragraph.Range;
                ReplaceText(rng, match.Index, match.Groups[0].Value,
                    $"{prefix}{quantity.ToString(format, CultureInfo.CurrentCulture)}\u2009{units}\u2009±\u2009{tolerance.ToString(format, CultureInfo.CurrentCulture)}\u2009{units}");
            }
            else if (match.Groups[6].Value.Length != 0 && match.Groups[8].Value.Length != 0)
            {
                float posTolerance = float.Parse(match.Groups[6].Value.Replace('\u2212', '-'), CultureInfo.CurrentCulture);
                posTolerance = Math.Abs(posTolerance);
                float negTolerance = float.Parse(match.Groups[8].Value.Replace('\u2212', '-'), CultureInfo.CurrentCulture);
                negTolerance = Math.Abs(negTolerance);
                rng = paragraph.Range;
                ReplaceText(rng, match.Index, match.Groups[0].Value,
                    $"{quantity.ToString(format, CultureInfo.CurrentCulture)}\u2009{units}\u2009+{posTolerance.ToString(format, CultureInfo.CurrentCulture)}\u2009{units}/\u2212{negTolerance.ToString(format, CultureInfo.CurrentCulture)}\u2009{units}");
            }
            else if (match.Groups[4].Value.Length == 0 && match.Groups[6].Value.Length == 0 && match.Groups[8].Value.Length == 0)
            {
                rng = paragraph.Range;
                if (match.Groups[1].Value.Length == 0)
                {
                    ReplaceText(rng, match.Index, match.Groups[0].Value, $"{quantity.ToString(format, CultureInfo.CurrentCulture)}\u2009{units}");
                }
                else
                {
                    ReplaceText(rng, match.Index, match.Groups[0].Value, $"{match.Groups[1].Value}\u2009{quantity.ToString(format, CultureInfo.CurrentCulture)}\u2009{units}");
                }
            }
        }
        private static Comment SetComment(Range rng, string commentText, int index = 0, int selectionLength = 1)
        {
            rng.Start = rng.Paragraphs[1].Range.Start + index;
            rng.End = rng.Paragraphs[1].Range.Start + index + selectionLength;
            rng.Select();
            Comment comment = Globals.ThisAddIn.Application.ActiveDocument.Comments.Add(Globals.ThisAddIn.Application.Selection.Range, commentText);
            comment.Author = "Toolbox";
            comment.Initial = "TBX";
            return comment;
        }
        private static void ReplaceText(Range rng, int index, string original, string replacement)
        {
            original = original.Trim();
            rng.Start = rng.Paragraphs[1].Range.Start + index;
            rng.End = rng.Paragraphs[1].Range.Start + index + original.Length;
            if (rng.Text != original)
            {
                rng.Start = rng.Start;
                rng.End = rng.Start + original.Length;
            }
            while (rng.Text != original && rng.End < rng.Paragraphs[1].Range.End)
            {
                rng.MoveStart(WdUnits.wdCharacter, 1);
                rng.MoveEnd(WdUnits.wdCharacter, 1);
            }
            rng.Select();
            if (rng.Text == original)
                Globals.ThisAddIn.Application.Selection.TypeText(replacement);
        }
        private string SafeSubstring(string input, int start, int length)
        {
            string result;
            if (start <= 0)
            {
                start = 0;
            }
            if (input.Length < start + length)
            {
                result = input.Substring(start);
            }
            else
            {
                result = input.Substring(start, length);
            }
            return result;
        }
    }
}
