using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using ExcelReport;
using Extend;

namespace ImageExample
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDlg = new SaveFileDialog();
            saveFileDlg.Filter = "Excel 2003文件|*.xls|Excel 2007文件|*.xlsx";

            if (DialogResult.OK.Equals(saveFileDlg.ShowDialog()))
            {
                //实例化一个参数容器，并加载模板填充规则文件
                ParameterCollection collection = new ParameterCollection();
                collection.Load(@"Template\Template.xml");

                //实例化一个元素格式化器列表
                List<ElementFormatter> formatters = new List<ElementFormatter>();
                formatters.Add(new CellFormatter(collection["Sheet1", "BMP"], Image.FromFile("Image/BMP.bmp").ToBuffer()));
                formatters.Add(new CellFormatter(collection["Sheet1", "JPG"], Image.FromFile("Image/JPG.jpg").ToBuffer()));
                formatters.Add(new CellFormatter(collection["Sheet1", "PNG"], Image.FromFile("Image/PNG.png").ToBuffer()));

                //导出文件到本地
                Export.ExportToLocal(@"Template\Template.xls", saveFileDlg.FileName,
                    new SheetFormatterContainer("Sheet1", formatters));
            }
        }
    }
}
