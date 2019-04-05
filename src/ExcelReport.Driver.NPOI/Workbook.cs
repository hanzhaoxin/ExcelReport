using ExcelReport.Driver.NPOI.Extends;
using NPOI.Extend;
using System.Collections;
using System.Collections.Generic;
using NpoiSheet = NPOI.SS.UserModel.ISheet;
using NpoiWorkbook = NPOI.SS.UserModel.IWorkbook;

namespace ExcelReport.Driver.NPOI
{
    public class Workbook : IWorkbook
    {
        public NpoiWorkbook NpoiWorkbook { get; private set; }

        public Workbook(NpoiWorkbook npoiWorkbook)
        {
            NpoiWorkbook = npoiWorkbook;
        }

        public ISheet this[string sheetName] => NpoiWorkbook.GetSheet(sheetName).GetAdapter();

        public byte[] SaveToBuffer()
        {
            return NpoiWorkbook.SaveToBuffer();
        }

        public IEnumerator<ISheet> GetEnumerator()
        {
            foreach (NpoiSheet npoiSheet in NpoiWorkbook)
            {
                yield return npoiSheet.GetAdapter();
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public object GetOriginal()
        {
            return NpoiWorkbook;
        }
    }
}