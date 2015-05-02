/*
 类：ExcelHelper
 描述：Excel相关操作助手类
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

using System.IO;
using NPOI.SS.UserModel;

namespace ERTool.Utilities
{
    public static class ExcelHelper
    {
        public static IWorkbook LoadWorkbook(string fileName)
        {
            using (var fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read)) //读入excel模板
            {
                return WorkbookFactory.Create(fileStream);
            }
        }
    }
}