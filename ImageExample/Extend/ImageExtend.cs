using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;

namespace Extend
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
