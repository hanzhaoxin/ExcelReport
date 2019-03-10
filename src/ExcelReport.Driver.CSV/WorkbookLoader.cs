using AxinLib.IO.CSV;
using ExcelReport.Driver.CSV.Extends;
using System.IO;

namespace ExcelReport.Driver.CSV
{
    public class WorkbookLoader : IWorkbookLoader
    {
        public IWorkbook Load(string filePath)
        {
            var workbook = new Workbook();
            var sheet = workbook[Path.GetFileNameWithoutExtension(filePath)];
            using (var streamReader = new StreamReader(filePath, EncodingExtend.GB2312))
            {
                var csvReader = new CsvReader(streamReader);
                Field field;
                while (null != (field = csvReader.ReadField()))
                {
                    sheet[field.RowIndex][field.FieldIndex].Value = field.Value;
                }
            }
            return workbook;
        }
    }
}