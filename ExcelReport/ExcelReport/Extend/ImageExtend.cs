/*
 类：ImageExtend
 描述：Image扩展方法
 编 码 人：韩兆新 日期：2015年05月01日
 修改记录：

*/
using System.IO;

namespace System.Drawing
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
