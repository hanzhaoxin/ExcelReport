/*
 类：ExportHelper
 描述：导出助手类
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

using System;
using System.IO;
using System.Text;
using System.Web;

namespace ExcelReport
{
    public static class ExportHelper
    {
        /// <summary>
        ///     导出到本地
        /// </summary>
        /// <param name="templateFile"></param>
        /// <param name="targetFile"></param>
        /// <param name="sheetFormatters"></param>
        public static void ExportToLocal(string templateFile, string targetFile, params SheetFormatter[] sheetFormatters)
        {
            #region 参数验证

            if (string.IsNullOrWhiteSpace(templateFile))
            {
                throw new ArgumentNullException("templateFile");
            }
            if (string.IsNullOrWhiteSpace(targetFile))
            {
                throw new ArgumentNullException("targetFile");
            }
            if (!File.Exists(templateFile))
            {
                throw new FileNotFoundException(templateFile + " 文件不存在!");
            }

            #endregion 参数验证

            using (FileStream fs = File.OpenWrite(targetFile))
            {
                byte[] buffer = Export.ExportToBuffer(templateFile, sheetFormatters);
                fs.Write(buffer, 0, buffer.Length);
                fs.Flush();
            }
        }

        /// <summary>
        ///     导出到Web
        /// </summary>
        /// <param name="templateFile"></param>
        /// <param name="targetFile"></param>
        /// <param name="sheetFormatters"></param>
        public static void ExportToWeb(string templateFile, string targetFile, params SheetFormatter[] sheetFormatters)
        {
            #region 参数验证

            if (string.IsNullOrWhiteSpace(templateFile))
            {
                throw new ArgumentNullException("templateFile");
            }
            if (string.IsNullOrWhiteSpace(targetFile))
            {
                throw new ArgumentNullException("targetFile");
            }
            if (!File.Exists(templateFile))
            {
                throw new FileNotFoundException(templateFile + " 文件不存在!");
            }

            #endregion 参数验证

            HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";
            HttpContext.Current.Response.AppendHeader("Content-Disposition",
                "attachment;filename=" + HttpUtility.UrlEncode(targetFile, Encoding.UTF8));
            HttpContext.Current.Response.BinaryWrite(Export.ExportToBuffer(templateFile, sheetFormatters));
            HttpContext.Current.Response.End();
        }
    }
}