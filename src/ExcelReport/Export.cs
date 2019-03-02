using ExcelReport.Contexts;
using ExcelReport.Driver;
using ExcelReport.Renderers;
using System.IO;

namespace ExcelReport
{
    public sealed class Export
    {
        public static byte[] ExportToBuffer(string templateFile, params SheetRenderer[] sheetRenderers)
        {
            var str = Path.GetExtension(templateFile);
            IWorkbookLoader workbookLoader = Configurator.Get(str);
            IWorkbook workbook = workbookLoader.Load(templateFile);
            var workbookContext = new WorkbookContext(workbook);
            foreach (SheetRenderer sheetRenderer in sheetRenderers)
            {
                sheetRenderer.Render(workbookContext);
            }
            return workbook.SaveToBuffer();
        }
    }
}