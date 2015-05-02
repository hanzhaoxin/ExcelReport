using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelReport;

namespace _01_多Sheet报表生成
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

                List<ElementFormatter> sheet1Formatters = new List<ElementFormatter>();
                sheet1Formatters.Add(new PartFormatter(collection["Sheet1", "Parameter"], "Parameter","World"));

                List<ElementFormatter> sheet2Formatters = new List<ElementFormatter>();
                sheet2Formatters.Add(new PartFormatter(collection["Sheet2", "Parameter"], "Parameter", "ExcelReport"));

                //导出文件到本地
                ExportHelper.ExportToLocal(@"Template\Template.xls", saveFileDlg.FileName,
                    new SheetFormatterContainer("Sheet1", sheet1Formatters),
                    new SheetFormatterContainer("Sheet2", sheet2Formatters)
                    );

            }
        }
    }
}
