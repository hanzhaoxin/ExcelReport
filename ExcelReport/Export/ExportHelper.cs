/*
 类：ExportHelper
 描述：导出助手类
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

using System.IO;
using NPOI.SS.UserModel;

namespace ExcelReport
{
    internal static class ExportHelper
    {
        #region 加载模板,获取IWorkbook对象

        private static IWorkbook LoadTemplateWorkbook(string templateFile)
        {
            using (var fileStream = new FileStream(templateFile, FileMode.Open, FileAccess.Read)) //读入excel模板
            {
                return WorkbookFactory.Create(fileStream);
            }
        }

        #endregion 加载模板,获取IWorkbook对象

        #region 将IWorkBook对象转换成二进制文件流

        private static byte[] SaveToBuffer(this IWorkbook workbook)
        {
            using (var ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                return ms.GetBuffer();
            }
        }

        #endregion 将IWorkBook对象转换成二进制文件流

        #region 导出格式化处理后的文件到二进制文件流

        public static byte[] ExportToBuffer(string templateFile, params SheetFormatterContainer[] containers)
        {
            var workbook = LoadTemplateWorkbook(templateFile);
            foreach (var container in containers)
            {
                var sheet = workbook.GetSheet(container.SheetName);    //加载第一个sheet
                var context = new SheetFormatterContext(sheet, container.Formatters);
                context.Format();
            }
            return workbook.SaveToBuffer();
        }

        #endregion 导出格式化处理后的文件到二进制文件流
    }
}