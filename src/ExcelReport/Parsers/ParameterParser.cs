using System.Text.RegularExpressions;

namespace ExcelReport.Parsers
{
    public sealed class ParameterParser : RegexParser
    {
        private static readonly Regex regex = new Regex(@"(?<=\$\[)([\w]*)(?=\])");

        public override Regex Regex => regex;
    }
}