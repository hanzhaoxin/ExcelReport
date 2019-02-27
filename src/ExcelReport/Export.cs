using ExcelReport.Contexts;
using ExcelReport.Renderers;
using NPOI.Extend;
using NPOI.SS.UserModel;

namespace ExcelReport
{
    public sealed class Export
    {
        public static byte[] ExportToBuffer(string templateFile, params SheetRenderer[] sheetRenderers)
        {
            IWorkbook workbook = NPOIHelper.LoadWorkbook(templateFile);
            var workbookContext = new WorkbookContext(workbook);
            foreach (SheetRenderer sheetRenderer in sheetRenderers)
            {
                sheetRenderer.Render(workbookContext);
            }
            return workbook.SaveToBuffer();
        }
    }
}