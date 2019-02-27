using ExcelReport.Renderers;
using System;
using System.IO;

namespace ExcelReport
{
    public static class ExportHelper
    {
        public static void ExportToLocal(string templateFile, string targetFile, params SheetRenderer[] sheetRenderers)
        {
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

            using (FileStream fs = File.OpenWrite(targetFile))
            {
                byte[] buffer = Export.ExportToBuffer(templateFile, sheetRenderers);
                fs.Write(buffer, 0, buffer.Length);
                fs.Flush();
            }
        }
    }
}