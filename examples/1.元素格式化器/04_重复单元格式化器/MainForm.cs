using System;
using System.Windows.Forms;
using ExcelReport;
using _04_重复单元格式化器.Properties;

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
            var saveFileDlg = new SaveFileDialog {Filter = Resources.SaveFileFilter};

            if (DialogResult.OK.Equals(saveFileDlg.ShowDialog()))
            {
                var workbookParameterContainer = new WorkbookParameterContainer();
                workbookParameterContainer.Load(@"Template\Template.xml");
                SheetParameterContainer sheetParameterContainer = workbookParameterContainer["重复单元格式化器"];

                ExportHelper.ExportToLocal(@"Template\Template.xls", saveFileDlg.FileName,
                    new SheetFormatter("重复单元格式化器",
                        new RepeaterFormatter<StudentInfo>(sheetParameterContainer["rptStudentInfo_Start"],
                            sheetParameterContainer["rptStudentInfo_End"], StudentLogic.GetList(),
                            new CellFormatter<StudentInfo>(sheetParameterContainer["Name"], t => t.Name),
                            new CellFormatter<StudentInfo>(sheetParameterContainer["Gender"], t => t.Gender ? "男" : "女"),
                            new CellFormatter<StudentInfo>(sheetParameterContainer["Class"], t => t.Class),
                            new CellFormatter<StudentInfo>(sheetParameterContainer["RecordNo"], t => t.RecordNo),
                            new CellFormatter<StudentInfo>(sheetParameterContainer["Phone"], t => t.Phone),
                            new CellFormatter<StudentInfo>(sheetParameterContainer["Email"], t => t.Email)
                            )
                        )
                    );
            }
        }
    }
}