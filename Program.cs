using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Collections.Generic;
using System.IO;
using System.Linq;


namespace PdfCopy
{
    static class Utils
    {
        public static LinkedList<T> ToLinkedList<T>(this IEnumerable<T> list)
        {
            return new LinkedList<T>(list);
        }
    }
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            // \u2019 is single curved right quotation mark
            // \u201d is double curved right quotation mark
            // (Book Title, 2010) <- don't want to change that to (Book Title,[2010])
            // avoid doing this to years by assuming footnotes will be 1-3 digits only
            Func<string, string> fixEndnotes = s => Regex.Replace(s, @"
                (?<=
                  (?<!\d)[,]
                  |
                  [a-zA-Z;]
                  |
                  [\])!.?""'\u2019\u201d]\ ?
                )
                (?<!
                  \b(?:pp?|vol|no|cf|ca|vv?|chap|[IVXLCDM]+)\.?\ ?
                  |
                  \d\.
                )
                (\d{1,3})
                (?![.:]\d|\w|:)", "[$1]", RegexOptions.IgnorePatternWhitespace); // second line was (?<!\d)[:,]
			Func<string, string> removeSomeEndnoteSpaces = s => Regex.Replace(s, @"(?<=[.""'\u2019\u201d]) (\[\d+\])", "$1");
            Func<string, string> stripNewlines = s => Regex.Replace(s, @"\s*\r?\n", " ");
            Func<string, string> fixEmDash = s => Regex.Replace(s, @"(?<=\w+)(\u2014) |-- ?", "\u2014");
            Func<string, string> fixHyphen = s => Regex.Replace(s, @"(?<!-)(-) ", "$1");
            Func<string, string> fixParen = s => Regex.Replace(s, @"\( ([^()]+[^ ])\)", "($1)");
            Func<string, string> makeEndnote = s => Regex.Replace(s, @"^\s*(\d+)\. ", "[$1] ");
            Func<string, string> fixSoftHyphens = s => Regex.Replace(s, @"\u00ad ?", "");
			Func<string, string> fixSpacedApostrophes = s => Regex.Replace(s, @" (?<x>\u2019)|(?<x>\u2018) ", "$1");
			Func<string, string> fixSpacedQuotes = s => Regex.Replace(s, @" (?<x>\u201d)|(?<x>\u201c) ", "$1");
            Func<string, string> ManInRevolt = s => s;// Regex.Replace(s, " (?<x>[;:])| ?(?<x>\u2014) ?", "$1");
            Func<string, string> JourneyOfModernTheology = s => Regex.Replace(s, @"(?<![.!?]) \.(?! ?\.)", ".");
            var inp = Clipboard.GetText();
            var modifications = new Func<string, string>[]
            {
                ConvertFirstLineAllCaps,
                s => stripNewlines(s).Trim(),
                fixSoftHyphens,
                fixEmDash,
                fixHyphen,
                fixParen,
                fixEndnotes,
                removeSomeEndnoteSpaces,
                makeEndnote,
                FixInterWordSpacing,
                ManInRevolt,
                fixSpacedApostrophes,
                fixSpacedQuotes,
                JourneyOfModernTheology,
            };
            var ret = inp;
            foreach (var f in modifications)
                ret = f(ret);
            //string debug = "justifi";
            //Clipboard.SetText(string.Join(" ", ret.Substring(ret.IndexOf(debug) + debug.Length - 2, 7).Select(c => string.Format("0x{0:x} ('{1}')", (int)c, c)).ToArray()));
            Clipboard.SetText(ret);
            //Console.WriteLine(ret);

            // Doing this via PowerShell, e.g.:
            //Console.InputEncoding = Console.OutputEncoding = Encoding.UTF8;
            //...
            //Console.WriteLine(ret);
            // $ [Console]::OutputEncoding = [Console]::InputEncoding = [System.Text.Encoding]::UTF8
            // $ Get-Clipboard | ./PdfCopy.exe | Set-Clipboard
            // did not work; emdash resulted in 0xFFFD insetad of 0x2014, for some reason
        }

        [Flags]
        private enum CharacterCasing
        {
            None = 0,
            Upper = 1,
            Lower = 2,
            Mixed = Upper | Lower
        }

        private static CharacterCasing GetCasing(string s)
        {
            var lower = Regex.IsMatch(s, "[a-z]") ? CharacterCasing.Lower : CharacterCasing.None;
            var upper = Regex.IsMatch(s, "[A-Z]") ? CharacterCasing.Upper : CharacterCasing.None;
            return lower | upper;
        }

        private static string FixInterWordSpacing(string ret)
        {
            // http://stackoverflow.com/questions/837488/how-can-i-get-the-applications-path-in-a-net-console-application/837501#837501
            var dictFullName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wordsEn.txt");
            HashSet<string> dict = new HashSet<string>(File.ReadAllLines(dictFullName), StringComparer.OrdinalIgnoreCase);
            // could have done this with string[], but that wouldn't have immutable semantics
            var tokens = Regex.Matches(ret, @"(\w+)|(\W+)").Cast<Match>()
                .Select(m => new { V = m.Value, IsWord = m.Groups[1].Success })
                .ToLinkedList();

            if (dict.Contains("th"))
                dict.Remove("th");

            // combine two adjacent parts of a word if either part is a spelling error
            // ignores proper nouns (spelling error if first letter not capitalized)
            for (var n = tokens.First; n != null && n.Next != null && n.Next.Next != null; n = n.Next)
            {
                var n0v = n.Previous != null ? n.Previous.Value.V : "";
                var n1 = n.Value;
                var n2 = n.Next.Value;
                var n3 = n.Next.Next.Value;
                bool sepSpace = n2.V == " " || n2.V == "- " || n2.V == "\u00ad ";
                bool eitherNotWord = !dict.Contains(n1.V) || !dict.Contains(n3.V);

                // \u00ad is soft hyphen; \u2019 is right single quotation mark
                // replace hyphenated words only if the second part isn't in the dictionary
                // guess that [Ww]ord ABBREVIATION should not be merged
                if (n0v != "'" && n0v != "\u2019" && // ignore the "t" and "s" in "don't", "dog's"
                    n1.IsWord && n3.IsWord &&
                    (sepSpace && eitherNotWord || n2.V == "-" && !dict.Contains(n3.V)) &&
                    (GetCasing(n1.V) == CharacterCasing.Upper || GetCasing(n3.V) == CharacterCasing.Lower) &&
                    dict.Contains(n1.V + n3.V))
                {
                    var ins = tokens.AddAfter(n.Next.Next, new { V = n1.V + n3.V, IsWord = true });
                    tokens.Remove(n.Next.Next);
                    tokens.Remove(n.Next);
                    tokens.Remove(n);
                    n = ins;
                }
            }

            // split words combined by f-kerning
            for (var n = tokens.First; n != null; n = n.Next)
            {
                var v = n.Value;

                if (!v.IsWord || dict.Contains(v.V))
                    continue;

                foreach (var t in PossibleFKerns(v.V))
                {
                    if (dict.Contains(t.Item1) && dict.Contains(t.Item2))
                    {
                        var ins = tokens.AddAfter(n, new { V = t.Item1 + " " + t.Item2, IsWord = true });
                        tokens.Remove(n);
                        n = ins;
                        break;
                    }
                }
            }

            return string.Join("", tokens.Select(e => e.V).ToArray());
        }

        private static IEnumerable<Tuple<string, string>> PossibleFKerns(string s)
        {
            // i = 1 means minimum first word length = 2
            var i = 1;
            while ((i = s.IndexOf('f', i) + 1) > 0)
                yield return Tuple.Create(s.Substring(0, i), s.Substring(i));
        }

        private static string ConvertFirstLineAllCaps(string s)
        {
            var nl = Regex.Match(s, @"\r\n?|\n");
            string newLine = nl.Value;
            var x = nl.Success ? s.Split(new[] { newLine }, 2, StringSplitOptions.None) : new[] { s };
            var firstLine = x[0];
            var rest = nl.Success ? x[1] : string.Empty;
            Predicate<string> sloppyIsRomanNumeral = r => Regex.IsMatch(r, @"^\W*[IVXLCDM]+\W*$", RegexOptions.IgnoreCase);


            // \u2013\u2014 are en, en dash
            // \u2018-\u201f are various curved quotation marks
            if (!Regex.IsMatch(firstLine, @"^[0-9A-Z .,:'""\-\u2013\u2014\u2018-\u201f&?]{2,}$"))
                return s;

            var noCap = new[] { "A", "THE", "IN", "FOR", "AND", "OF", "WITH", "TO", "THAT", "ON", "AS", "IS" };

            firstLine = Regex.Replace(firstLine, @"([\w'\u2019])(\w*)", m => sloppyIsRomanNumeral(m.Value)
                ? m.Value
                : noCap.Contains(m.Value)
                ? m.Value.ToLower()
                : m.Groups[1].Value + m.Groups[2].Value.ToLower());

            // we might have lower-cased the first character (or subtitle, etc.)
            firstLine = Regex.Replace(firstLine, @"(?<=^\W*|[:""']\s*)\w", m => m.Value.ToUpper());

            return nl.Success
                ? firstLine + nl.Value + rest
                : firstLine;
        }
    }
}
