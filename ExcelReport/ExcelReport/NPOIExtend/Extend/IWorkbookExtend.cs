/*
 类：IWorkbookExtend
 描述：IWorkbook扩展方法
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

using System.IO;
using NPOI.SS.UserModel;

namespace ExcelReport
{
    internal static class IWorkbookExtend
    {
        /// 将workbook转换成二进制文件流
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static byte[] SaveToBuffer(this IWorkbook workbook)
        {
            using (var ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                return ms.GetBuffer();
            }
        }
    }
}
