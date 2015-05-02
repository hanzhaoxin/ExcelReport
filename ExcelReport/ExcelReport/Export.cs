/*
 类：Export
 描述：导出
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

namespace ExcelReport
{
    public static class Export
    {
        /// 导出格式化处理后的文件到二进制文件流
        /// <param name="templateFile"></param>
        /// <param name="containers"></param>
        /// <returns></returns>
        public static byte[] ExportToBuffer(string templateFile, params SheetFormatterContainer[] containers)
        {
            var workbook = NPOIHelper.LoadWorkbook(templateFile);
            foreach (var container in containers)
            {
                var sheet = workbook.GetSheet(container.SheetName);
                var context = new SheetFormatterContext(sheet, container.Formatters);
                context.Format();
            }
            return workbook.SaveToBuffer();
        }
    }
}