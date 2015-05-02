using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelReport;

namespace _01_生成工资表_工资条报表
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
                collection.Load(@"Template\演示：工资条模板.xml");

                List<ElementFormatter> payTableFormatters = new List<ElementFormatter>();
                payTableFormatters.Add(new PartFormatter(collection["工资表", "Dept"], "Dept", "生产A部"));
                payTableFormatters.Add(new CellFormatter(collection["工资表", "Month"], "2015年4月"));
                payTableFormatters.Add(new TableFormatter<SalaryInfo>(collection["工资表", "JobNo"].X, SalaryInfoLogic.GetList(),
                    new TableColumnInfo<SalaryInfo>(collection["工资表", "JobNo"].Y, t => t.JobNo),
                    new TableColumnInfo<SalaryInfo>(collection["工资表", "Name"].Y, t => t.Name),
                    new TableColumnInfo<SalaryInfo>(collection["工资表", "HourlyWage"].Y, t => t.HourlyWage),
                    new TableColumnInfo<SalaryInfo>(collection["工资表", "ManHour"].Y, t => t.ManHour),
                    new TableColumnInfo<SalaryInfo>(collection["工资表", "UsualOvertime"].Y, t => t.UsualOvertime),
                    new TableColumnInfo<SalaryInfo>(collection["工资表", "WeekendOvertime"].Y, t => t.WeekendOvertime),
                    new TableColumnInfo<SalaryInfo>(collection["工资表", "LegalOvertime"].Y, t => t.LegalOvertime),
                    new TableColumnInfo<SalaryInfo>(collection["工资表", "BasePay"].Y, t => t.BasePay),
                    new TableColumnInfo<SalaryInfo>(collection["工资表", "OvertimePay"].Y, t => t.OvertimePay),
                    new TableColumnInfo<SalaryInfo>(collection["工资表", "FoodAllowance"].Y, t => t.FoodAllowance),
                    new TableColumnInfo<SalaryInfo>(collection["工资表", "OtherAllowance"].Y, t => t.OtherAllowance),
                    new TableColumnInfo<SalaryInfo>(collection["工资表", "GrossPay"].Y, t => t.GrossPay),
                    new TableColumnInfo<SalaryInfo>(collection["工资表", "BoardWages"].Y, t => t.BoardWages),
                    new TableColumnInfo<SalaryInfo>(collection["工资表", "Utilities"].Y, t => t.Utilities),
                    new TableColumnInfo<SalaryInfo>(collection["工资表", "BadDebt"].Y, t => t.BadDebt),
                    new TableColumnInfo<SalaryInfo>(collection["工资表", "RealWages"].Y, t => t.RealWages),
                    new TableColumnInfo<SalaryInfo>(collection["工资表", "Remark"].Y, t => t.Remark)
                    ));

                List<ElementFormatter> payStripFormatters = new List<ElementFormatter>();
                payStripFormatters.Add(new RepeaterFormatter<SalaryInfo>(collection["工资条", "rptPaySlipStart"], collection["工资条", "rptPaySlipEnd"], SalaryInfoLogic.GetList(),
                    new RepeaterCellInfo<SalaryInfo>(collection["工资条", "Title"], t => "某工厂生产A部工资条"),
                    new RepeaterCellInfo<SalaryInfo>(collection["工资条", "Month"], t => "2015年4月"),
                    new RepeaterCellInfo<SalaryInfo>(collection["工资条", "JobNo"], t => t.JobNo),
                    new RepeaterCellInfo<SalaryInfo>(collection["工资条", "Name"], t => t.Name),
                    new RepeaterCellInfo<SalaryInfo>(collection["工资条", "HourlyWage"], t => t.HourlyWage),
                    new RepeaterCellInfo<SalaryInfo>(collection["工资条", "ManHour"], t => t.ManHour),
                    new RepeaterCellInfo<SalaryInfo>(collection["工资条", "UsualOvertime"], t => t.UsualOvertime),
                    new RepeaterCellInfo<SalaryInfo>(collection["工资条", "WeekendOvertime"], t => t.WeekendOvertime),
                    new RepeaterCellInfo<SalaryInfo>(collection["工资条", "LegalOvertime"], t => t.LegalOvertime),
                    new RepeaterCellInfo<SalaryInfo>(collection["工资条", "BasePay"], t => t.BasePay),
                    new RepeaterCellInfo<SalaryInfo>(collection["工资条", "OvertimePay"], t => t.OvertimePay),
                    new RepeaterCellInfo<SalaryInfo>(collection["工资条", "FoodAllowance"], t => t.FoodAllowance),
                    new RepeaterCellInfo<SalaryInfo>(collection["工资条", "OtherAllowance"], t => t.OtherAllowance),
                    new RepeaterCellInfo<SalaryInfo>(collection["工资条", "GrossPay"], t => t.GrossPay),
                    new RepeaterCellInfo<SalaryInfo>(collection["工资条", "BoardWages"], t => t.BoardWages),
                    new RepeaterCellInfo<SalaryInfo>(collection["工资条", "Utilities"], t => t.Utilities),
                    new RepeaterCellInfo<SalaryInfo>(collection["工资条", "BadDebt"], t => t.BadDebt),
                    new RepeaterCellInfo<SalaryInfo>(collection["工资条", "RealWages"], t => t.RealWages),
                    new RepeaterCellInfo<SalaryInfo>(collection["工资条", "Remark"], t => t.Remark)
                    ));

                //导出文件到本地
                ExportHelper.ExportToLocal(@"Template\演示：工资条模板.xls", saveFileDlg.FileName,
                    new SheetFormatterContainer("工资表", payTableFormatters),
                    new SheetFormatterContainer("工资条", payStripFormatters)
                    );
            }
        }
    }
}
