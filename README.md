# MS Word Toolbox

While working at an aerospace firm, I spend far more time writing documents in Microsoft Word than doing things I enjoy, like coding. But I found that some of my tasks could be automated in Word through judicious use of macros. I spent...far too long writing those macros in VBA, which, in addition to being [dog-slow](https://ell.stackexchange.com/questions/140818/etymology-of-dog-slow?answertab=votes#tab-top), leaves much to be desired in terms of things like error handling and dialog design. Sure, there are a couple of advantages, but those can be co-opted or overcome easily in a different language. There do remain a few issues which are more related to Microsoft Word itself than any language used to interface _with_ it. Plus, it gave me an excuse to learn a new language, and to code!

TL;DR

I wrote some of what I consider to be useful macros for MS Word in C#. This is them.

**_Issues and pull requests are welcome!_**

## Fix Quantities

One pet peeve of mine is the expression of quantites with tolerances in engineering documents. They should match (to the degree possible, even according to our own rules) [NIST SP811](https://www.nist.gov/pml/special-publication-811). I see all to often things like the following:

|Incorrect|Correct|Note|
|-|-|-|
|100+5mS/-5mSec|100 ms ± 5 ms|I often see "millisecond" abbreviated as "mS", "mSec", and/or "msec". While I sometimes see multiple abbreviations in the same document, they're rarely in the same quantity as I've displayed here. Due to an...overzealous prescriptivist, some of our documents have positive and negative tolerances specified separately, despite being the same value, when there's a perfectly good "±" symbol available for use.|
|15 KOhm|15 kΩ|The "kilo-" prefix should be lowercase, especially when abbreviated. Additionally "Ohm" should not be spelled out if a prefix isn't likewise spelled out: the symbol "Ω" [U+03A9] (and not "Ω" [U+2126]) should be used instead.|
|65 KB|64 KiB|One exception is for binary units which are sometimes ignored (despite guidance) and SI prefixes are used instead.|
|115uA|115 μA|I often see the "micro-" prefix abbreviated with a "u" instead of "μ" [U+03BC] or even "µ" [U+00B5].|
|28±0.50V|28.00 V ± 0.50 V|While not perfect, the number of decimal places can be used as a proxy for significant figures in much of our documentation.|
|1mA±5μA|1.000 mA ± 0.0005 mA|Units must match, regardless of how many decimal places this adds to one or the other of the quantities.|

Sometimes the errors are out of laziness: "I can't be bothered to figure out how to type '±', 'μ', or 'Ω', let me just type what I know and move on." I can understand this sentiment, but it takes very little effort to rectify the situation. Nonetheless, this macro fixes (most) of the above. I can't accurately correct the KB/KiB conversion, so that is ignored, partly by technical limitations, and partly by design.

### Method

A breakdown of my algorithm:

1. Fix common mistakes
    1. Replace instances of "µ" [U+00B5] with "μ" [U+03BC]
    2. Replace instances of "" [U+F06D] in Symbol font with "μ" [U+03BC] in the font of the "Normal" style.
    3. Replace "Ω" [U+2126] with "Ω" [U+03A9] throughout document (per Note 2 of Unicode Standard Character 2126)
    4. Replace "" [U+F057] in Symbol font with "Ω" [U+03A9] in the font of the "Normal" style.
    5. Replace a "u" preceding any SI unit with "μ" [U+03BC]
    6. Replace "msec", "mSec", and "mS" with "ms"
    7. Replace "μsec", "μSec", "μS", "usec", "uSec", and "uS" with "μs"
    8. Replace the word "ohm" if it follows an SI prefix character
    9. Replace a superscript "o", "º" [U+00A7] or "⁰" [U+2070] (sometimes incorrectly used as degree signs) with "°" [U+00B0]
    10. Replace "K" with "k" if used as an SI prefix
    11. TODO: Replace certain dashes with a hyphen-minus (for use in float.Parse)
2. Iterate through matches of the [regex](#the-regex)
    1. Capture all units in the regex (see groups)
    2. See if the units are in agreement (or if there's only one)
    3. See if the units are recognized
    4. Set the unit string to be used
    5. Count the decimal places of the provided numbers (used as a proxy for significant figures) and prepare to format all numbers with that number of digits after the decimal place
    6. Parse the tolerance and decide whether symmetric or asymmetric
    7. If units match (from step 2) and a tolerance was included, build result string
    8. If units match (from step 2) but no tolerance was included
        1. Checks if the units are recognized (step 3)
        2. Build string if they are
        3. Ignore if not
    9. If units do not match, flag with a comment and move to next match (skip step 10)
    10. Replace match with "fixed" string

### The Regex

My first stab at a regular expression is as follows (newlines and tabs added for clarity):

```csharp
private const string quantity = @"[-−‐-―]?(?:\d+\.\d+|\d+\.|\.\d+|\d+)";
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
```

This regular expression matches a number with any word that follows. Obviously, this is too generous. I could eliminate the `|(\w+\b)` piece, but it serves a useful purpose. In those cases where a non-standard unit is used, I don't want it to be neglected. If there is a tolerance listed, I wish to either format the quantiy properly, or flag if mixed units are present. If someone writes "123foo+456bar/-789baz", that should be flagged. If, instead, someone wrote "999±1bacon", it should be parsed and reformatted as "999 bacon ± 1 bacon". Delicious. On the other hand, if someone writes "Section 22 Category R", I definitely *do not* want it to be parsed as "22 Category" and reformatted as "22 Category". This is why it's separated from the `units` definition—so it's *captured* separately. So a good portion of the logic will come after a match has been made, based on which groups actually capture something.

In the full expanded form, it looks like this:

`([≤≥<>\+±]?)[ \t  -   ]*([-−‐-―]?(?:\d+\.\d+|\d+|\.\d+|\d+\.))[ \t  -   ]*(?:(?:((?:(?:(?:da|[YZEPTGMkhdcmμnpfazy])?(?:Wb|Sv|Hz|sr|mol|lm|lx|cd|rad|Pa|Bq|Da|eV|ua|Gy|kat|°C|[gmsulAKNJWCVFSTHLΩ])|m?Np|B|dB(?:FS|iC|m0s?|mV|ov|pp|rnC|sm|TP|μV|μ0s|VU| HL| Q| SIL| SPL| SWL|\/K|-Hz|[ABcCdefGiJkKmoOqruvVWZμ])?|[KMGTPEZY]i[bB]|[Mkdcm]?bar|mmHg|ha|min|[Åbhdt]|kts?|ft|lbs?|inHg|n?mi|psi|atm|°F|VDC)\b)|%)|(\w+\b))[ \t  -   ]*)?(?:(?:±|\+\/[-−‐-―])[ \t  -   ]*([-−‐-―]?(?:\d+\.\d+|\d+|\.\d+|\d+\.))[ \t  -   ]*(?:((?:(?:(?:da|[YZEPTGMkhdcmμnpfazy])?(?:Wb|Sv|Hz|sr|mol|lm|lx|cd|rad|Pa|Bq|Da|eV|ua|Gy|kat|°C|[gmsulAKNJWCVFSTHLΩ])|m?Np|B|dB(?:FS|iC|m0s?|mV|ov|pp|rnC|sm|TP|μV|μ0s|VU| HL| Q| SIL| SPL| SWL|\/K|-Hz|[ABcCdefGiJkKmoOqruvVWZμ])?|[KMGTPEZY]i[bB]|[Mkdcm]?bar|mmHg|ha|min|[Åbhdt]|kts?|ft|lbs?|inHg|n?mi|psi|atm|°F|VDC)\b)|%)|(\w+\b))|\+[ \t  -   ]*([-−‐-―]?(?:\d+\.\d+|\d+|\.\d+|\d+\.))[ \t  -   ]*(?:((?:(?:(?:da|[YZEPTGMkhdcmμnpfazy])?(?:Wb|Sv|Hz|sr|mol|lm|lx|cd|rad|Pa|Bq|Da|eV|ua|Gy|kat|°C|[gmsulAKNJWCVFSTHLΩ])|m?Np|B|dB(?:FS|iC|m0s?|mV|ov|pp|rnC|sm|TP|μV|μ0s|VU| HL| Q| SIL| SPL| SWL|\/K|-Hz|[ABcCdefGiJkKmoOqruvVWZμ])?|[KMGTPEZY]i[bB]|[Mkdcm]?bar|mmHg|ha|min|[Åbhdt]|kts?|ft|lbs?|inHg|n?mi|psi|atm|°F|VDC)\b)|%)|(\w+\b))[ \t  -   ]*\/[ \t  -   ]*[-−‐-―][ \t  -   ]*([-−‐-―]?(?:\d+\.\d+|\d+|\.\d+|\d+\.))[ \t  -   ]*(?:((?:(?:(?:da|[YZEPTGMkhdcmμnpfazy])?(?:Wb|Sv|Hz|sr|mol|lm|lx|cd|rad|Pa|Bq|Da|eV|ua|Gy|kat|°C|[gmsulAKNJWCVFSTHLΩ])|m?Np|B|dB(?:FS|iC|m0s?|mV|ov|pp|rnC|sm|TP|μV|μ0s|VU| HL| Q| SIL| SPL| SWL|\/K|-Hz|[ABcCdefGiJkKmoOqruvVWZμ])?|[KMGTPEZY]i[bB]|[Mkdcm]?bar|mmHg|ha|min|[Åbhdt]|kts?|ft|lbs?|inHg|n?mi|psi|atm|°F|VDC)\b)|%)|(\w+\b)))?`

There are 14 capture groups here, as follows:

|Group|Contains|
|:-:|-|
|0|Full Match|
|1|Preamble (leading inequality, plus sign, or plus/minus sign)|
|2|Main Quantity (number)|
|3|Recognized Main Unit|
|4|Unrecognized Main Unit|
|5|Symmetric Tolerance Quantity|
|6|Symmetric Tolerance Recognized Unit|
|7|Symmetric Tolerance Unrecognized Unit|
|8|Asymmetric Positive Tolerance Quantity|
|9|Asymmetric Positive Tolerance Recognized Unit|
|10|Asymmetric Positive Tolerance Unrecognized Unit|
|11|Asymmetric Negative Tolerance Quantity|
|12|Asymmetric Negative Tolerance Recognized Unit|
|13|Asymmetric Negative Tolerance Unrecognized Unit|

This allows some checks to be done. As there are many, many possibilities, I've left the details in the code for now.

### Limitations

1. Word "fields" are weird (see [here](https://findingmyname.com/microsoft-word-again/) for some of my research).
2. Comments, bookmarks, and (as discussed) fields, count as special characters in Word and thus, if surrounding or interrupting a quantity when the macro is run, those values will not be matched or altered. In fact, if a bookmark or comment is encountered, the entire paragraph in which it occurs is ignored. Fixing a single quantity via the context menu will still work in most cases.

### Interesting Tidbits

- Word (2016-2019, at least) treats a thin space [U+2009] as a non-breaking space, so I utilized it instead of the narrow no-break space [U+202F]. Stylistically, quantities and their units and/or tolerances should not be broken across lines, so if Word ever changes this behavior, an alteration might be necessary. It might be valuable to include this as an option in a settings dialog.

## TODO

- Fix Quantities
  - [x] Only fix quantities in highlighted section, if applicable (entire document if nothing's highlighted)
    - [ ] My fix is rather simplistic; could/should be improved.
  - [ ] Fix progress bar: needs to run in separate thread or otherwise force visual updates
  - [ ] Add settings/options dialog
  - [ ] Replace `'` with `′` and `"` with `″` where appropriate
  - [ ] Allow use of planar angles (e.g. 15° ± 1°)
  - [ ] Allow compound units and units with exponents, e.g. m<sup>2</sup>, m / s, V / m, N • m
- Add other macros
  - [ ] Edit Properties dialog
    - [x] Get Document Value
    - [x] Set Document Value
  - [x] Import/Export Properties
  - [x] ~~Toggle Page Break Before~~
  - [x] Selection To Link
    - [ ] Improve dialog asking what type of reference to insert
    - [ ] Allow linking to numbered lists under sections given full context
  - [x] ~~Find Broken Links~~
  - [ ] Create Acronym Table
  - [x] ~~Character Gallery~~
    - [ ] Add selection of dashes and spaces (with labels)
    - [ ] Dynamically generate images with current font? (investigate feasibility)
