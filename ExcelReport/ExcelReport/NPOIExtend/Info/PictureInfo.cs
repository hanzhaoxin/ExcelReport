/*
 类：PictureInfo
 描述：图片信息
 编 码 人：韩兆新 日期：2015年04月11日
 修改记录：

*/
using System;

namespace ExcelReport
{
    public class PictureInfo
    {
        public int MinRow { get; set; }
        public int MaxRow { get; set; }
        public int MinCol { get; set; }
        public int MaxCol { get; set; }
        public Byte[] PictureData { get; set; }
        public PictureStyle PicturesStyle { get; set; }

        public PictureInfo(int minRow, int maxRow, int minCol, int maxCol, Byte[] pictureData, PictureStyle pictureStyle)
        {
            this.MinRow = minRow;
            this.MaxRow = maxRow;
            this.MinCol = minCol;
            this.MaxCol = maxCol;
            this.PictureData = pictureData;
            this.PicturesStyle = pictureStyle;
        }
    }
}
