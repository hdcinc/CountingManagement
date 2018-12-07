using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using unvell.ReoGrid;
using unvell.ReoGrid.CellTypes;

namespace CountingManagement
{
    public partial class FormPLTop : Form
    {

        private readonly Color CELL_BACK_COLOR_EDIT_ONLY = Color.FromArgb(255,255,255);
        private readonly Color CELL_BACK_COLOR_EDIT_ABLE = Color.FromArgb(189, 215, 238);
        private readonly Color CELL_BACK_COLOR_READ_ONLY = Color.FromArgb(217, 217, 217);


        public FormPLTop()
        {
            InitializeComponent();
        }

        private void FormPLTop_Load(object sender, EventArgs e)
        {
            reoGridControl1.Load(@"..\CM_BIN\ReoGridPLTop.xlsx");

            Worksheet sheet = reoGridControl1.CurrentWorksheet;

            sheet.BeforeCellEdit += Sheet_BeforeCellEdit;
            sheet.AfterCellEdit += Sheet_AfterCellEdit;
            sheet.CellEditTextChanging += Sheet_CellEditTextChanging;
            sheet.CellEditCharInputed += Sheet_CellEditCharInputed;

            sheet.CellDataChanged += Sheet_CellDataChanged;
            sheet.CellKeyUp += Sheet_CellKeyUp;

            sheet.FocusPosChanged += Sheet_FocusPosChanged;


            reoGridControl1.ActionPerformed += ReoGridControl1_ActionPerformed;




            sheet["B2"] = new ButtonCell("Hello");

            WorksheetRangeStyle wrs = new WorksheetRangeStyle();

            wrs.Flag = PlainStyleFlag.BackColor;
            wrs.BackColor = Color.Red;

            sheet.SetRangeStyles("B3", wrs);

        }

        private void ReoGridControl1_ActionPerformed(object sender, unvell.ReoGrid.Events.WorkbookActionEventArgs e)
        {
        }

        private void Sheet_FocusPosChanged(object sender, unvell.ReoGrid.Events.CellPosEventArgs e)
        {
        }

        private void Sheet_CellKeyUp(object sender, unvell.ReoGrid.Events.CellKeyDownEventArgs e)
        {
            Console.WriteLine("Sheet_CellKeyUp");

            if (e.KeyCode == unvell.ReoGrid.Interaction.KeyCode.Delete)
            {
                Worksheet worksheet = reoGridControl1.CurrentWorksheet;

                unvell.ReoGrid.Events.CellEventArgs eventArgs = new unvell.ReoGrid.Events.CellEventArgs(worksheet.GetCell(worksheet.FocusPos));
                Sheet_CellDataChanged(sender, eventArgs);
            }
        }

        private void Sheet_CellDataChanged(object sender, unvell.ReoGrid.Events.CellEventArgs e)
        {
            Console.WriteLine("Sheet_CellDataChanged");

            Worksheet worksheet = reoGridControl1.CurrentWorksheet;

            switch (e.Cell.Address)
            {
                case "N7":
                    if (worksheet.GetCell("N13").Style.BackColor != CELL_BACK_COLOR_EDIT_ONLY)
                    {
                        worksheet.GetCell("N13").Data = ConvertCellData(worksheet.GetCellData("N7")) - ConvertCellData(worksheet.GetCellData("N10"));
                    }



                    break;
                case "N13":

                    Decimal calValue = ConvertCellData(worksheet.GetCellData("N7")) - ConvertCellData(worksheet.GetCellData("N10"));
                    Decimal newValue = ConvertCellData(e.Cell.Data);

                    WorksheetRangeStyle style = new WorksheetRangeStyle();
                    style.Flag = PlainStyleFlag.BackColor;
                    if (calValue == newValue)
                    {
                        style.BackColor = CELL_BACK_COLOR_EDIT_ABLE;
                    }
                    else
                    {
                        style.BackColor = CELL_BACK_COLOR_EDIT_ONLY;
                    }
                    worksheet.SetRangeStyles("N13", style);

                    break;
                default:
                    break;
            }

        }

        private void Sheet_BeforeCellEdit(object sender, unvell.ReoGrid.Events.CellBeforeEditEventArgs e)
        {
            Console.WriteLine("Sheet_BeforeCellEdit");
        }

        private void Sheet_AfterCellEdit(object sender, unvell.ReoGrid.Events.CellAfterEditEventArgs e)
        {
            Console.WriteLine("Sheet_AfterCellEdit");

            //Worksheet worksheet = reoGridControl1.CurrentWorksheet;

            //if (e.Cell.Address.Equals("N13"))
            //{
            //    Decimal calValue = ConvertCellData(worksheet.GetCellData("N7")) - ConvertCellData(worksheet.GetCellData("N10"));
            //    Decimal newValue = ConvertCellData(e.NewData);

            //    WorksheetRangeStyle style = new WorksheetRangeStyle();
            //    style.Flag = PlainStyleFlag.BackColor;
            //    if (calValue == newValue)
            //    {
            //        style.BackColor = Color.Red;
            //    }
            //    else
            //    {
            //        style.BackColor = Color.White;
            //    }
            //    worksheet.SetRangeStyles("N13", style);
            //}
        }

        private Decimal ConvertCellData(object cellData)
        {
            Decimal value = 0;

            Decimal.TryParse(Convert.ToString(cellData), out value);

            return value;
        }

        private void Sheet_CellEditTextChanging(object sender, unvell.ReoGrid.Events.CellEditTextChangingEventArgs e)
        {
            Console.WriteLine("Sheet_CellEditTextChanging");
        }

        private void Sheet_CellEditCharInputed(object sender, unvell.ReoGrid.Events.CellEditCharInputEventArgs e)
        {
            Console.WriteLine("Sheet_CellEditCharInputed");
        }


        private void button1_Click(object sender, EventArgs e)
        {
            Worksheet sheet = reoGridControl1.CurrentWorksheet;

            MessageBox.Show(sheet.GetCellData("N7").ToString());
        }
    }
}
