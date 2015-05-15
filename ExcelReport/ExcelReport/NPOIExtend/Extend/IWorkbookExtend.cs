/*
 类：IWorkbookExtend
 描述：IWorkbook扩展方法
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：
    1: 修改人：JsonRuby  日期：2015年05月14日
       修改内容：SaveToBuffer()，修复Bug（xlsx作为模板时，导出到Buffer异常[无法访问已关闭的流]）。
    2: 修改人：韩兆新  日期：2015年05月16日
       修改内容：SaveToBuffer()，修复Bug（xlsx作为模板时，导出文档提示文档中存在不可读取内容，需要修复。]）。
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
                return ms.ToArray();
            }
        }
    }
}
