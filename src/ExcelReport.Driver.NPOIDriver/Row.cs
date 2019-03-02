using ExcelReport.Driver.NPOIDriver.Extends;
using System.Collections;
using System.Collections.Generic;
using NpoiCell = NPOI.SS.UserModel.ICell;
using NpoiRow = NPOI.SS.UserModel.IRow;

namespace ExcelReport.Driver.NPOIDriver
{
    public class Row : IRow
    {
        public NpoiRow NpoiRow { get; private set; }

        public Row(NpoiRow npoiRow)
        {
            NpoiRow = npoiRow;
        }

        public ICell this[int columnIndex] => NpoiRow.GetCell(columnIndex).GetAdapter();

        public IEnumerator<ICell> GetEnumerator()
        {
            foreach (NpoiCell npoiCell in NpoiRow)
            {
                yield return npoiCell.GetAdapter();
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}