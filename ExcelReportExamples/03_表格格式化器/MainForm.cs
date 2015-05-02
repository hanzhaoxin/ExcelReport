using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelReport;

namespace _03_表格格式化器
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
                int num = 0;
                List<ElementFormatter> formatters = new List<ElementFormatter>();
                formatters.Add(new TableFormatter<StudentInfo>(collection["Sheet1", "No"].X, StudentLogic.GetList(),
                    new TableColumnInfo<StudentInfo>(collection["Sheet1", "No"].Y, t => num++),
                    new TableColumnInfo<StudentInfo>(collection["Sheet1", "Name"].Y, t => t.Name),
                    new TableColumnInfo<StudentInfo>(collection["Sheet1", "Gender"].Y, t => t.Gender ? "男" : "女"),
                    new TableColumnInfo<StudentInfo>(collection["Sheet1", "Class"].Y, t => t.Class),
                    new TableColumnInfo<StudentInfo>(collection["Sheet1", "RecordNo"].Y, t => t.RecordNo),
                    new TableColumnInfo<StudentInfo>(collection["Sheet1", "Phone"].Y, t => t.Phone),
                    new TableColumnInfo<StudentInfo>(collection["Sheet1", "Email"].Y, t => t.Email) 
                    ));

                //导出文件到本地
                ExportHelper.ExportToLocal(@"Template\Template.xls", saveFileDlg.FileName,
                    new SheetFormatterContainer("Sheet1", formatters)
                    );

            }
        }
    }
}
