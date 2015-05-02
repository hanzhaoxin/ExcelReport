/*
 类：ISheetExtend
 描述：ISheet扩展方法
 编 码 人：韩兆新 日期：2015年04月11日
 修改记录：
    1: 修改人：韩兆新  日期：2015年05月01日
       修改内容：添加了获取合并区域信息列表的扩展方法GetAllMergedRegionInfos();
                 添加了添加合并区域的扩展方法AddMergedRegion();

*/
using System;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace ExcelReport
{
    internal static class ISheetExtend
    {
        #region 图片
        /// sheet中添加图片
        /// <param name="sheet"></param>
        /// <param name="picInfo"></param>
        public static void AddPicture(this ISheet sheet, PictureInfo picInfo)
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

        /// 获取sheet中包含图片的信息列表
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static List<PictureInfo> GetAllPictureInfos(this ISheet sheet)
        {
            return sheet.GetAllPictureInfos(null, null, null, null);
        }

        /// 获取sheet中指定区域包含图片的信息列表
        /// <param name="sheet"></param>
        /// <param name="minRow"></param>
        /// <param name="maxRow"></param>
        /// <param name="minCol"></param>
        /// <param name="maxCol"></param>
        /// <param name="onlyInternal"></param>
        /// <returns></returns>
        public static List<PictureInfo> GetAllPictureInfos(this ISheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol, bool onlyInternal = true)
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

        #region 实现
        private static List<PictureInfo> GetAllPictureInfos(HSSFSheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol, bool onlyInternal)
        {
            List<PictureInfo> picturesInfoList = new List<PictureInfo>();

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
                            var picStyle = new PictureStyle
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
                            picturesInfoList.Add(new PictureInfo(anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, picture.PictureData.Data, picStyle));
                        }
                    }
                }
            }

            return picturesInfoList;
        }

        private static List<PictureInfo> GetAllPictureInfos(XSSFSheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol, bool onlyInternal)
        {
            List<PictureInfo> picturesInfoList = new List<PictureInfo>();

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
                                var picStyle = new PictureStyle
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
                                picturesInfoList.Add(new PictureInfo(anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, picture.PictureData.Data, picStyle));
                            }
                        }
                    }
                }
            }

            return picturesInfoList;
        }
        
        #endregion
        #endregion

        #region 合并区域
        
        /// sheet中添加合并区域
        /// <param name="sheet"></param>
        /// <param name="regionInfo"></param>
        public static void AddMergedRegion(this ISheet sheet, MergedRegionInfo regionInfo)
        {
            var region = new CellRangeAddress(regionInfo.FirstRow,regionInfo.LastRow,regionInfo.FirstCol,regionInfo.LastCol);
            sheet.AddMergedRegion(region);
        }
        
        /// 获取sheet中包含合并区域的信息列表
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static List<MergedRegionInfo> GetAllMergedRegionInfos(this ISheet sheet)
        {
            return sheet.GetAllMergedRegionInfos(null, null, null, null);
        }

        /// 获取sheet中指定区域包含合并区域的信息列表
        /// <param name="sheet"></param>
        /// <param name="minRow"></param>
        /// <param name="maxRow"></param>
        /// <param name="minCol"></param>
        /// <param name="maxCol"></param>
        /// <param name="onlyInternal"></param>
        /// <returns></returns>
        public static List<MergedRegionInfo> GetAllMergedRegionInfos(this ISheet sheet, int? minRow, int? maxRow,
            int? minCol, int? maxCol, bool onlyInternal = true)
        {
            var regionInfoList = new List<MergedRegionInfo>();
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                var range = sheet.GetMergedRegion(i);
                if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, range.FirstRow, range.LastRow,
                    range.FirstColumn, range.LastColumn, onlyInternal))
                {
                    regionInfoList.Add(new MergedRegionInfo(range.FirstRow, range.LastRow, range.FirstColumn, range.LastColumn));
                }
            }
            return regionInfoList;
        }

        #endregion

        #region 辅助
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
