using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelReport;

namespace _02_单元格格式化器
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDlg = new SaveFileDialog();
            saveFileDlg.Filter = "Excel 2003文件|*.xls|Excel 2007文件|*.xlsx";

            if (DialogResult.OK.Equals(saveFileDlg.ShowDialog()))
            {

                ParameterCollection collection = new ParameterCollection();
                collection.Load(@"Template\Template.xml");

                List<ElementFormatter> formatters = new List<ElementFormatter>();
                formatters.Add(new CellFormatter(collection["Sheet1","String"],"Hello World!"));
                formatters.Add(new CellFormatter(collection["Sheet1", "DateTime"], DateTime.Now));
                formatters.Add(new CellFormatter(collection["Sheet1", "Boolean"], true));
                formatters.Add(new CellFormatter(collection["Sheet1", "Double"], 3.14));
                formatters.Add(new CellFormatter(collection["Sheet1", "Image"], Image.FromFile("Image/C#高级编程.jpg").ToBuffer()));

                //导出文件到本地
                ExportHelper.ExportToLocal(@"Template\Template.xls", saveFileDlg.FileName,
                    new SheetFormatterContainer("Sheet1", formatters)
                    );

            }
        }
    }
}
