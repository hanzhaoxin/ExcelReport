/*
 类：WorkbookParameterContainer
 描述：工作薄参数容器
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：
 * 2015年08月01日（韩兆新）  修改自ExcelReport_v1.xx中的ParameterCollection。

*/
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;
using ExcelReport.Exceptions;

namespace ExcelReport
{
    public class WorkbookParameterContainer
    {
        private List<SheetParameterContainer> _sheetParameterContainerList = new List<SheetParameterContainer>();

        /// <summary>
        ///     Sheet参数容器
        /// </summary>
        /// <param name="sheetName">sheet名</param>
        /// <returns>Sheet参数容器</returns>
        public SheetParameterContainer this[string sheetName]
        {
            get
            {
                foreach (SheetParameterContainer sheetParameterContainer in _sheetParameterContainerList)
                {
                    if (sheetParameterContainer.SheetName.Equals(sheetName))
                    {
                        return sheetParameterContainer;
                    }
                }
                throw new ExcelReportTemplateException("sheet is not exists");
            }
            set
            {
                bool isExist = false;
                foreach (SheetParameterContainer sheetParameterContainer in _sheetParameterContainerList)
                {
                    if (sheetParameterContainer.SheetName.Equals(sheetName))
                    {
                        isExist = true;
                        sheetParameterContainer.ParameterList = value.ParameterList;
                    }
                }
                if (!isExist)
                {
                    _sheetParameterContainerList.Add(value);
                }
            }
        }

        #region 1.0 加载xml文件

        /// <summary>
        ///     加载xml文件
        /// </summary>
        /// <param name="xmlPath">xml文件路径</param>
        public void Load(string xmlPath)
        {
            string fileName = xmlPath;
            if (File.Exists(fileName))
            {
                var xmlSerializer = new XmlSerializer(_sheetParameterContainerList.GetType());
                using (var reader = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    _sheetParameterContainerList = xmlSerializer.Deserialize(reader) as List<SheetParameterContainer>;
                }
            }
            else
            {
                throw new ExcelReportTemplateException("template file is not exists");
            }
        }

        #endregion

        #region 1.1 保存xml文件

        /// <summary>
        ///     保存xml文件
        /// </summary>
        /// <param name="xmlPath">xml文件路径</param>
        public void Save(string xmlPath)
        {
            string fileName = xmlPath;
            var fileInfo = new FileInfo(fileName);
            DirectoryInfo directoryInfo = fileInfo.Directory;
            if (directoryInfo != null && !directoryInfo.Exists)
            {
                directoryInfo.Create();
            }
            var xmlSerializer = new XmlSerializer(_sheetParameterContainerList.GetType());
            using (var writer = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                xmlSerializer.Serialize(writer, _sheetParameterContainerList);
            }
        }

        #endregion
    }
}