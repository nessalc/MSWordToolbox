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
|115uA|115 μA|I often see the "micro-" prefix abbreviated with a "u" instead of "μ" [U+03BC] or "µ" [U+00B5].|
|28±0.50V|28.00 V ± 0.50 V|While not perfect, the number of decimal places can be used as a proxy for significant figures in much of our documentation.|
|1mA±5μA|1.000 mA ± 0.0005 mA|Units must match, regardless of how many decimal places this adds to one or the other of the quantities.|

Sometimes the errors are out of laziness: "I can't be bothered to figure out how to type '±', 'μ', or 'Ω', let me just type what I know and move on." I can understand this sentiment, but it takes very little effort to rectify the situation. Nonetheless, this macro fixes (most) of the above. I can't accurately correct the KB/KiB conversion, so that is ignored, partly by technical limitations, and partly by design.

### Method

A breakdown of my algorithm:

1. Replace "Ω" [U+2126] with "Ω" [U+03A9] throughout document (per Note 2 of Unicode Standard Character 2126)
2. Construct regular expression to find probable quantity definitions
   - Due to limitations in MS Word regular expressions, two searches are necessary: one preceded by `(?<=[[\s\(])` for all quantities that follow a word or are in the middle of a paragrah, and one preceded by `^` for those that begin a paragraph or are inside a table.
3. Iterate over paragraphs:
   1. Replace "mS", "msec", "mSec" with "ms", provided previous character is not a letter.
   2. Replace "sec", "second", "seconds" with "s", provided previous character is not a letter. **NOTE: This may cause issues!**
   3. Replace "us", "usec", uS", and "uSec" with μs", provided "u" is preceded by a boundary.
   4. Replace "ohm" (regardless of case) with "Ω", provided "ohm" is succeeded by a boundary. **NOTE: This may cause issues with units that are fully spelled out.**
   5. (Finally) perform actual "fixes":
      1. The regular expression provides the following match groups:
         |Match Group|Description|
         |-|-|
         |0|Entire Match|
         |1|Inequality, or one of [±+-]|
         |2|Main quantity|
         |3|Main unit|
         |4|(Symmetric) Tolerance quantity|
         |5|(Symmetric) Tolerance unit|
         |6|Positive tolerance quantity|
         |7|Positive tolerance unit
         |8|Negative tolerance quantity|
         |9|Negative tolerance unit|
      2. If value only without units or tolerance (or match occurred in an equation), ignore and move on (could be, say, a zip code or other innocent bystander number that, honestly, shouldn't have triggered a match in the first place, but this is error checking).
      3. Check that provided unit(s) are all the same; if not, flag mismatch in comment.
      4. Count decimal places (as proxy for significant figures) in quantity, and use maximum value.
      5. If value had a tolerance, but no units attached to it, flag in comment.
      6. If unit prefix is "K", replace with "k" **NOTE: This _will_ mess up "KiB"**
      7. Replace text with appropriate reformatted value.

### Challenges

1. Word's implementation of regular expressions (and this may be a POSIX thing that I wasn't aware of) does not necessarily match the beginning of a line as a boundary character.
2. `Ω`, however, despite (and possibly because) being a Greek letter, means a boundary character _does_ exist between `m` and `Ω` in `mΩ`. If the Unicode flag is active, the boundary is handled as desired; need to test if RegexOptions.CultureInvariant is functionally equivalent in C#.
3. Order very much matters in regular expression, as the first match wins, so there was a _lot_ of reordering of my expression to get to the current version.

### Limitations

1. Word "fields" do not count as characters in text (or, rather, the entire field counts as a single character...I think), so when selecting text to replace, selection would not necessarily be in the expected location.
2. Comments, bookmarks, and (as discussed) fields, count as special characters in Word and thus, if surrounding or interrupting a quantity when the macro is run, those values will not be matched or altered.

### Interesting Tidbits

- Word (2016-2019, at least) treats a thin space [U+2009] as a non-breaking space, so I utilized it instead of the narrow no-break space [U+202F]. Stylistically, quantities and their units and/or tolerances should not be broken across lines, so if Word ever changes this behavior, an alteration might be necessary. It might be valuable to include this as an option in a settings dialog.

## TODO

- Fix Quantities
  - [ ] Add settings/options dialog
  - [ ] Add right-click menu for updating a single quantity
  - [ ] Allow planar angle markings (°′″) to be printed _without_ thin space preceding them.
  - [ ] Replace `'` with `′` and `"` with `″` where appropriate
  - [ ] Allow compound units and units with exponents, e.g. m<sup>2</sup>, m / s, V / m, N • m
  - [ ] Fix issues noted above under [Method](#method)
- Add other macros
  - Edit Properties dialog
    - Get Document Value
    - Set Document Value
  - Import/Export Properties
  - Toggle Page Break Before
  - Selection To Link
  - Find Broken Links
  - Create Acronym Table
  - (they exist, but are...not ready, even for GitHub, yet)
