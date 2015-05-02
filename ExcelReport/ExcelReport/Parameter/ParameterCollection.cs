/*
 类：ParameterCollection
 描述：模板中参数信息的集合
 编 码 人：韩兆新 日期：2015年01月17日
 修改记录：

*/

using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Xml.Serialization;

namespace ExcelReport
{
    public class ParameterCollection
    {
        #region 成员字段及属性
        protected List<Parameter> parameterList = new List<Parameter>();
        #endregion

        #region 索引器
        public Point this[string sheetName, string parameterName]
        {
            get
            {
                foreach (Parameter parameter in parameterList)
                {
                    if (parameter.SheetName.Equals(sheetName) && parameter.ParameterName.Equals(parameterName))
                    {
                        return parameter.CellPoint;
                    }
                }
                return new Point();
            }
            set
            {
                bool isExist = false;
                foreach (Parameter parameter in parameterList)
                {
                    if (parameter.SheetName.Equals(sheetName) && parameter.ParameterName.Equals(parameterName))
                    {
                        isExist = true;
                        parameter.CellPoint = value;
                    }
                }
                if (!isExist)
                {
                    parameterList.Add(new Parameter(sheetName, parameterName, value));
                }
            }
        }
        #endregion

        /// 加载xml文件
        /// <param name="xmlPath"></param>
        public void Load(string xmlPath)
        {
            string fileName = xmlPath;
            if (File.Exists(fileName))
            {
                XmlSerializer xmlSerializer = new XmlSerializer(parameterList.GetType());
                using (Stream reader = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    parameterList = xmlSerializer.Deserialize(reader) as List<Parameter>;
                }
            }
            else
            {
                parameterList = new List<Parameter>();
            }
        }

        /// 保存xml文件
        /// <param name="xmlPath"></param>
        public void Save(string xmlPath)
        {
            string fileName = xmlPath;
            FileInfo fileInfo = new FileInfo(fileName);
            DirectoryInfo directoryInfo = fileInfo.Directory;
            if (!directoryInfo.Exists)
            {
                directoryInfo.Create();
            }
            XmlSerializer xmlSerializer = new XmlSerializer(parameterList.GetType());
            using (Stream writer = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                xmlSerializer.Serialize(writer, parameterList);
            }
        }
    }
}