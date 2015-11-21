using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelReport;
using _02_单元格格式化器.Properties;

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
            var saveFileDlg = new SaveFileDialog {Filter = Resources.SaveFileFilter};

            if (DialogResult.OK.Equals(saveFileDlg.ShowDialog()))
            {
                var workbookParameterContainer = new WorkbookParameterContainer();
                workbookParameterContainer.Load(@"Template\Template.xml");
                var sheetParameterContainer = workbookParameterContainer["单元格格式化器"];

                ExportHelper.ExportToLocal(@"Template\Template.xls", saveFileDlg.FileName,
                    new SheetFormatter("单元格格式化器",
                        new CellFormatter(sheetParameterContainer["String"], "Hello World!"),
                        new CellFormatter(sheetParameterContainer["Boolean"], true),
                        new CellFormatter(sheetParameterContainer["DateTime"], DateTime.Now),
                        new CellFormatter(sheetParameterContainer["Double"], 3.14),
                        new CellFormatter(sheetParameterContainer["Image"],Image.FromFile("Image/C#高级编程.jpg").ToBuffer())
                        )
                    );
            }
        }
    }
}