using System;
using System.Windows.Forms;
using ExcelReport;
using _03_表格格式化器.Properties;

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
            var saveFileDlg = new SaveFileDialog {Filter = Resources.SaveFileFilter};

            if (DialogResult.OK.Equals(saveFileDlg.ShowDialog()))
            {
                var workbookParameterContainer = new WorkbookParameterContainer();
                workbookParameterContainer.Load(@"Template\Template.xml");
                SheetParameterContainer sheetParameterContainer = workbookParameterContainer["表格格式化器"];

                int num = 0;
                ExportHelper.ExportToLocal(@"Template\Template.xls", saveFileDlg.FileName,
                    new SheetFormatter("表格格式化器",
                        new PartFormatter(sheetParameterContainer["Author"],"hzx"),
                        new TableFormatter<StudentInfo>(sheetParameterContainer["No"], StudentLogic.GetList(),
                            new CellFormatter<StudentInfo>(sheetParameterContainer["No"], t => num++),
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