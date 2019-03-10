using AxinLib.IO.CSV;
using System.IO;
using System.Text;

namespace ExcelReport.Driver.CSV
{
    public class WorkbookLoader : IWorkbookLoader
    {
        public IWorkbook Load(string filePath)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var workbook = new Workbook();
            var sheet = workbook[Path.GetFileNameWithoutExtension(filePath)];
            using (var streamReader = new StreamReader(filePath, Encoding.GetEncoding("GB2312")))
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