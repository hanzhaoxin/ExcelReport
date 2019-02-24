using System.Drawing;
using System.IO;

namespace ExcelReport.Extends
{
    public static class ImageExtend
    {
        public static byte[] ToBuffer(this Image img)
        {
            using (var ms = new MemoryStream())
            {
                img.Save(ms, img.RawFormat);
                ms.Flush();
                ms.Position = 0;
                return ms.GetBuffer();
            }
        }
    }
}