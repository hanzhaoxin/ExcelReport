using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace ExcelReport.Driver.CSV
{
    public class Row : IRow, ICloneable, ICsvBuilder
    {
        private readonly Dictionary<int, Cell> _cellDic = new Dictionary<int, Cell>();

        public int RowIndex { get; set; }

        public Row(int rowIndex)
        {
            RowIndex = rowIndex;
        }

        public ICell this[int columnIndex]
        {
            get
            {
                if (_cellDic.ContainsKey(columnIndex))
                {
                    return _cellDic[columnIndex];
                }

                var cell = new Cell(this, columnIndex);
                _cellDic[cell.ColumnIndex] = cell;
                return cell;
            }
        }

        public object Clone()
        {
            var row = new Row(RowIndex);
            foreach (var cell in _cellDic.Values)
            {
                row[cell.ColumnIndex].Value = cell.Value;
            }
            return row;
        }

        public IEnumerator<ICell> GetEnumerator()
        {
            foreach (var cell in _cellDic.Values)
            {
                yield return cell;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _cellDic.Values.GetEnumerator();
        }

        public void AppendTo(StringBuilder builder)
        {
            var index = 0;
            var maxIndex = _cellDic.Count - 1;
            foreach (Cell cell in this)
            {
                cell.AppendTo(builder);
                if (index++ < maxIndex)
                {
                    builder.Append(Constant.DELIMITER);
                }
            }
            builder.Append(Constant.ROW_END);
        }

        public object GetOriginal()
        {
            return this;
        }
    }
}