using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReport
{
    public class PicturesInfo
    {
        public int MinRow { get; set; }
        public int MaxRow { get; set; }
        public int MinCol { get; set; }
        public int MaxCol { get; set; }
        public Byte[] PictureData { get; set; }
        public PicturesStyle PicturesStyle { get; set; }

        public PicturesInfo(int minRow, int maxRow, int minCol, int maxCol, Byte[] pictureData, PicturesStyle picturesStyle)
        {
            this.MinRow = minRow;
            this.MaxRow = maxRow;
            this.MinCol = minCol;
            this.MaxCol = maxCol;
            this.PictureData = pictureData;
            this.PicturesStyle = picturesStyle;
        }
    }
}
