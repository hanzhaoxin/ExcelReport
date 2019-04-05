using ExcelReport.Driver.CSV.Extends;
using System.Text;

namespace ExcelReport.Driver.CSV
{
    public class Cell : ICell, ICsvBuilder
    {
        private readonly Row _row;

        private readonly int _columnIndex;

        public Cell(Row row, int columnIndex)
        {
            _row = row;
            _columnIndex = columnIndex;
        }

        public int RowIndex => _row.RowIndex;

        public int ColumnIndex => _columnIndex;

        public object Value { get; set; }

        public void AppendTo(StringBuilder builder)
        {
            var value = Value.ToString();
            var needEscape = value.ContainAny(Constant.NEED_ESCAPE_CHARS);
            if (needEscape)
            {
                builder.Append(Constant.ESCAPE);
            }
            foreach (var c in value)
            {
                if (Constant.ESCAPE.Equals(c))
                {
                    builder.Append(Constant.ESCAPE);
                }
                builder.Append(c);
            }
            if (needEscape)
            {
                builder.Append(Constant.ESCAPE);
            }
        }

        public object GetOriginal()
        {
            return this;
        }
    }
}