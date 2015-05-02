using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelReport;

namespace _04_重复单元格式化器
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
                formatters.Add(new RepeaterFormatter<StudentInfo>(collection["Sheet1", "rptStudentInfo_Start"],collection["Sheet1", "rptStudentInfo_End"], StudentLogic.GetList(),
                    new RepeaterCellInfo<StudentInfo>(collection["Sheet1", "Name"], t => t.Name),
                    new RepeaterCellInfo<StudentInfo>(collection["Sheet1", "Gender"], t => t.Gender ? "男" : "女"),
                    new RepeaterCellInfo<StudentInfo>(collection["Sheet1", "Class"], t => t.Class),
                    new RepeaterCellInfo<StudentInfo>(collection["Sheet1", "RecordNo"], t => t.RecordNo),
                    new RepeaterCellInfo<StudentInfo>(collection["Sheet1", "Phone"], t => t.Phone),
                    new RepeaterCellInfo<StudentInfo>(collection["Sheet1", "Email"], t => t.Email)
                    ));

                //导出文件到本地
                ExportHelper.ExportToLocal(@"Template\Template.xls", saveFileDlg.FileName,
                    new SheetFormatterContainer("Sheet1", formatters)
                    );

            }
        }
    }
}
