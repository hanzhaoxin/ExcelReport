/*
 类：NPOIHelper
 描述：NPOI操作助手类
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/
using System.IO;
using NPOI.SS.UserModel;

namespace XmlGenerator
{
    internal static class NPOIHelper
    { 
        /// 加载模板,获取IWorkbook对象
        /// <param name="file"></param>
        /// <returns></returns>
        public static IWorkbook LoadWorkbook(string file)
        {
            using (var fileStream = new FileStream(file, FileMode.Open, FileAccess.Read)) //读入excel模板
            {
                return WorkbookFactory.Create(fileStream);
            }
        }
    }
}
