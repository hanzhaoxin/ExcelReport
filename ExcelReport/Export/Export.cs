/*
 类：Export
 描述：导出
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

using System;
using System.IO;
using System.Web;

namespace ExcelReport
{
    public static class Export
    {
        #region 导出到本地

        public static void ExportToLocal(string templateFile, string targetFile, params SheetFormatterContainer[] containers)
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
                var buffer = ExportHelper.ExportToBuffer(templateFile, containers);
                fs.Write(buffer, 0, buffer.Length);
                fs.Flush();
            }
        }

        #endregion 导出到本地

        #region 导出到Web

        public static void ExportToWeb(string templateFile, string targetFile, params SheetFormatterContainer[] containers)
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
            HttpContext.Current.Response.AppendHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(targetFile, System.Text.Encoding.UTF8));
            HttpContext.Current.Response.BinaryWrite(ExportHelper.ExportToBuffer(templateFile, containers));
            HttpContext.Current.Response.End();
        }

        #endregion 导出到Web
    }
}