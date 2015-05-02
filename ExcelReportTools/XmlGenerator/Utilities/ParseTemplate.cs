/*
 类：ParseTemplate
 描述：解析模板
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

using System.Drawing;
using System.Text.RegularExpressions;
using ExcelReport;
using NPOI.SS.UserModel;

namespace XmlGenerator
{
    internal static class ParseTemplate
    {
        public static ParameterCollection Parse(string templatePath)
        {
            ParameterCollection workbookParameter = new ParameterCollection();
            IWorkbook workbook = NPOIHelper.LoadWorkbook(templatePath);
            foreach (ISheet sheet in workbook)
            {
                foreach (IRow row in sheet)
                {
                    foreach (ICell cell in row.Cells)
                    {
                        if (cell.CellType.Equals(CellType.String))
                        {
                            string cellText = cell.StringCellValue;
                            MatchCollection matches = new Regex(@"(?<=\$\[)([\w]*)(?=\])").Matches(cellText);
                            foreach (Match match in matches)
                            {
                                workbookParameter[sheet.SheetName, match.Value] = new Point(cell.RowIndex, cell.ColumnIndex);
                            }
                        }
                    }
                }
            }
            return workbookParameter;
        }
    }
}