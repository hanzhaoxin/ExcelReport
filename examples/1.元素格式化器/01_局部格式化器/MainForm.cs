using System;
using System.Windows.Forms;
using ExcelReport;
using _01_局部格式化器.Properties;

namespace _01_局部格式化器
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
                var sheetParameterContainer = workbookParameterContainer["局部格式化器"];

                ExportHelper.ExportToLocal(@"Template\Template.xls", saveFileDlg.FileName,
                    new SheetFormatter("局部格式化器",
                        new PartFormatter(sheetParameterContainer["UserName"], "Jensen"),
                        new PartFormatter(sheetParameterContainer["GroupNo"], "116476496")
                        )
                    );
            }
        }
    }
}