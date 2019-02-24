using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelReport.Parsers
{
    public abstract class RegexParser
    {
        public abstract Regex Regex { get; }

        public IEnumerable<string> Parse(string content)
        {
            MatchCollection matches = Regex.Matches(content);
            foreach (Match match in matches)
            {
                yield return match.Value;
            }
        }
    }
}