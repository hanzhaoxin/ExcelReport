using System.Text;

namespace ExcelReport.Driver.CSV
{
    internal interface ICsvBuilder
    {
        void AppendTo(StringBuilder builder);
    }
}