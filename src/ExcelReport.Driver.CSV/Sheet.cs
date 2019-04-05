using ExcelReport.Driver.CSV.Extends;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReport.Driver.CSV
{
    public class Sheet : ISheet, ICsvBuilder
    {
        private readonly LinkedList<Row> _rows = new LinkedList<Row>();

        public string SheetName { get; private set; }

        public Sheet(string sheetName)
        {
            SheetName = sheetName;
        }

        public IRow this[int rowIndex]
        {
            get
            {
                var row = _rows.SingleOrDefault(r => r.RowIndex.Equals(rowIndex));
                if (null == row)
                {
                    row = new Row(rowIndex);
                    _rows.AddLast(row);
                }
                return row;
            }
        }

        public int CopyRows(int start, int end)
        {
            var shiftCount = 1 + end - start;
            var copyRows = new LinkedList<Row>();
            foreach (var row in _rows)
            {
                if (row.RowIndex > end)
                {
                    row.RowIndex += shiftCount;
                }
                else if (row.RowIndex >= start)
                {
                    var newRow = (Row)row.Clone();
                    newRow.RowIndex += shiftCount;
                    copyRows.AddLast(newRow);
                }
            }
            _rows.AppendAll(copyRows);
            return copyRows.Count;
        }

        public IEnumerator<IRow> GetEnumerator()
        {
            return _rows.OrderBy(row => row.RowIndex).GetEnumerator();
        }

        public int RemoveRows(int start, int end)
        {
            var shiftCount = 1 + end - start;
            var removeRows = new LinkedList<Row>();
            foreach (var row in _rows)
            {
                if (row.RowIndex > end)
                {
                    row.RowIndex -= shiftCount;
                }
                else if (row.RowIndex >= start)
                {
                    removeRows.AddLast(row);
                }
            }
            _rows.RemoveAll(removeRows);
            return removeRows.Count;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _rows.OrderBy(row => row.RowIndex).GetEnumerator();
        }

        public void AppendTo(StringBuilder builder)
        {
            foreach (Row row in this)
            {
                row.AppendTo(builder);
            }
        }

        public object GetOriginal()
        {
            return this;
        }
    }
}