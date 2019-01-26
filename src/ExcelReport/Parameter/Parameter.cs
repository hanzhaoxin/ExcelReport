/*
 类：Parameter
 描述：参数信息
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：
 * 2015年08月01日  删除CellPoint，添加RowIndex、ColumnIndex。

*/

using ExcelReport.Exceptions;
using System.Collections.Generic;

namespace ExcelReport
{
    public class Parameter
    {
        private List<CellPoint> _cellPointList = new List<CellPoint>();

        public string Name { set; get; }
        
        public List<CellPoint> CellPointList
        {
            get
            {
                return _cellPointList;
            }
        }
        public int TagRowIndex
        {
            get
            {
                if (_cellPointList.IsNullOrEmpty())
                {
                    throw new ExcelReportTemplateException("parameter is not exists");
                }

                return _cellPointList[0].RowIndex;
            }
        }

    }
}