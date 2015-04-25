using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelReport
{
    public static class ISheetExtend
    {
        #region ISheetExtend
        public static void AddPicture(this ISheet sheet, PicturesInfo picInfo)
        {
            var pictureIdx = sheet.Workbook.AddPicture(picInfo.PictureData, PictureType.PNG);
            var anchor = sheet.Workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Col1 = picInfo.MinCol;
            anchor.Col2 = picInfo.MaxCol;
            anchor.Row1 = picInfo.MinRow;
            anchor.Row2 = picInfo.MaxRow;
            anchor.Dx1 = picInfo.PicturesStyle.AnchorDx1;
            anchor.Dx2 = picInfo.PicturesStyle.AnchorDx2;
            anchor.Dy1 = picInfo.PicturesStyle.AnchorDy1;
            anchor.Dy2 = picInfo.PicturesStyle.AnchorDy2;
            anchor.AnchorType = 2;
            var drawing = sheet.CreateDrawingPatriarch();
            var pic = drawing.CreatePicture(anchor, pictureIdx);
            if (sheet is HSSFSheet)
            {
                var shape = (HSSFShape)pic;
                shape.FillColor = picInfo.PicturesStyle.FillColor;
                shape.IsNoFill = picInfo.PicturesStyle.IsNoFill;
                shape.LineStyle = picInfo.PicturesStyle.LineStyle;
                shape.LineStyleColor = picInfo.PicturesStyle.LineStyleColor;
                shape.LineWidth = (int)picInfo.PicturesStyle.LineWidth;
            }
            else if (sheet is XSSFSheet)
            {
                var shape = (XSSFShape)pic;
                shape.FillColor = picInfo.PicturesStyle.FillColor;
                shape.IsNoFill = picInfo.PicturesStyle.IsNoFill;
                shape.LineStyle = picInfo.PicturesStyle.LineStyle;
                //shape.LineStyleColor = picInfo.PicturesStyle.LineStyleColor;
                shape.LineWidth = picInfo.PicturesStyle.LineWidth;
            }
        }

        public static List<PicturesInfo> GetAllPictureInfos(this ISheet sheet)
        {
            return sheet.GetAllPictureInfos(null, null, null, null);
        }

        public static List<PicturesInfo> GetAllPictureInfos(this ISheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol, bool onlyInternal = true)
        {
            if (sheet is HSSFSheet)
            {
                return GetAllPictureInfos((HSSFSheet)sheet, minRow, maxRow, minCol, maxCol, onlyInternal);
            }
            else if (sheet is XSSFSheet)
            {
                return GetAllPictureInfos((XSSFSheet)sheet, minRow, maxRow, minCol, maxCol, onlyInternal);
            }
            else
            {
                throw new Exception("未处理类型，没有为该类型添加：GetAllPicturesInfos()扩展方法！");
            }
        }

        private static List<PicturesInfo> GetAllPictureInfos(HSSFSheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol, bool onlyInternal)
        {
            List<PicturesInfo> picturesInfoList = new List<PicturesInfo>();

            var shapeContainer = sheet.DrawingPatriarch as HSSFShapeContainer;
            if (null != shapeContainer)
            {
                var shapeList = shapeContainer.Children;
                foreach (var shape in shapeList)
                {
                    if (shape is HSSFPicture && shape.Anchor is HSSFClientAnchor)
                    {
                        var picture = (HSSFPicture)shape;
                        var anchor = (HSSFClientAnchor)shape.Anchor;
                        if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, onlyInternal))
                        {
                            var picStyle = new PicturesStyle
                            {
                                AnchorDx1 = anchor.Dx1,
                                AnchorDx2 = anchor.Dx2,
                                AnchorDy1 = anchor.Dy1,
                                AnchorDy2 = anchor.Dy2,
                                IsNoFill = picture.IsNoFill,
                                LineStyle = picture.LineStyle,
                                LineStyleColor = picture.LineStyleColor,
                                LineWidth = picture.LineWidth,
                                FillColor = picture.FillColor
                            };
                            picturesInfoList.Add(new PicturesInfo(anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, picture.PictureData.Data, picStyle));
                        }
                    }
                }
            }

            return picturesInfoList;
        }

        private static List<PicturesInfo> GetAllPictureInfos(XSSFSheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol, bool onlyInternal)
        {
            List<PicturesInfo> picturesInfoList = new List<PicturesInfo>();

            var documentPartList = sheet.GetRelations();
            foreach (var documentPart in documentPartList)
            {
                if (documentPart is XSSFDrawing)
                {
                    var drawing = (XSSFDrawing)documentPart;
                    var shapeList = drawing.GetShapes();
                    foreach (var shape in shapeList)
                    {
                        if (shape is XSSFPicture)
                        {
                            var picture = (XSSFPicture)shape;
                            var anchor = picture.GetPreferredSize();

                            if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, onlyInternal))
                            {
                                var picStyle = new PicturesStyle
                                {
                                    AnchorDx1 = anchor.Dx1,
                                    AnchorDx2 = anchor.Dx2,
                                    AnchorDy1 = anchor.Dy1,
                                    AnchorDy2 = anchor.Dy2,
                                    //IsNoFill = picture.IsNoFill,
                                    //LineStyle = picture.LineStyle,
                                    //LineStyleColor = picture.LineStyleColor,
                                    //LineWidth = picture.LineWidth,
                                    //FillColor = picture.FillColor
                                };
                                picturesInfoList.Add(new PicturesInfo(anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, picture.PictureData.Data, picStyle));
                            }
                        }
                    }
                }
            }

            return picturesInfoList;
        }

        private static bool IsInternalOrIntersect(int? rangeMinRow, int? rangeMaxRow, int? rangeMinCol, int? rangeMaxCol,
            int pictureMinRow, int pictureMaxRow, int pictureMinCol, int pictureMaxCol, bool onlyInternal)
        {
            int _rangeMinRow = rangeMinRow ?? pictureMinRow;
            int _rangeMaxRow = rangeMaxRow ?? pictureMaxRow;
            int _rangeMinCol = rangeMinCol ?? pictureMinCol;
            int _rangeMaxCol = rangeMaxCol ?? pictureMaxCol;

            if (onlyInternal)
            {
                return (_rangeMinRow <= pictureMinRow && _rangeMaxRow >= pictureMaxRow &&
                        _rangeMinCol <= pictureMinCol && _rangeMaxCol >= pictureMaxCol);
            }
            else
            {
                return ((Math.Abs(_rangeMaxRow - _rangeMinRow) + Math.Abs(pictureMaxRow - pictureMinRow) >= Math.Abs(_rangeMaxRow + _rangeMinRow - pictureMaxRow - pictureMinRow)) &&
                (Math.Abs(_rangeMaxCol - _rangeMinCol) + Math.Abs(pictureMaxCol - pictureMinCol) >= Math.Abs(_rangeMaxCol + _rangeMinCol - pictureMaxCol - pictureMinCol)));
            }
        }
        #endregion
    }
}
