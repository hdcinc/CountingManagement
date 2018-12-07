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

        private Worksheet worksheet;

        public FormPLTop()
        {
            InitializeComponent();
        }

        private void FormPLTop_Load(object sender, EventArgs e)
        {
            reoGridControl1.Load(@"..\CM_BIN\ReoGridPLTop.xlsx");

            worksheet = reoGridControl1.CurrentWorksheet;

            worksheet.CellDataChanged += Sheet_CellDataChanged;
            worksheet.CellKeyUp += Sheet_CellKeyUp;
        }

        private void Sheet_CellKeyUp(object sender, unvell.ReoGrid.Events.CellKeyDownEventArgs e)
        {
            Console.WriteLine("Sheet_CellKeyUp");

            if (e.KeyCode == unvell.ReoGrid.Interaction.KeyCode.Delete)
            {
                unvell.ReoGrid.Events.CellEventArgs eventArgs = new unvell.ReoGrid.Events.CellEventArgs(worksheet.GetCell(worksheet.FocusPos));
                Sheet_CellDataChanged(sender, eventArgs);
            }
        }

        private void Sheet_CellDataChanged(object sender, unvell.ReoGrid.Events.CellEventArgs e)
        {
            Console.WriteLine("Sheet_CellDataChanged");

            switch (e.Cell.Address)
            {
                case "N7":
                    if (worksheet.GetCell("N13").Style.BackColor != CELL_BACK_COLOR_EDIT_ONLY)
                    {
                        worksheet.GetCell("N13").Data = ConvertCellData(worksheet.GetCellData("N7")) - ConvertCellData(worksheet.GetCellData("N10"));
                    }
                    if (worksheet.GetCell("N19").Style.BackColor != CELL_BACK_COLOR_EDIT_ONLY)
                    {
                        worksheet.GetCell("N19").Data = ConvertCellData(worksheet.GetCellData("N13")) - ConvertCellData(worksheet.GetCellData("N16"));
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

        private Decimal ConvertCellData(object cellData)
        {
            Decimal value = 0;

            Decimal.TryParse(Convert.ToString(cellData), out value);

            return value;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            worksheet["売上総利益_計画_2015"] = 111;
        }

        private void ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            worksheet.InsertRows(worksheet.FocusPos.Row, 1);
        }

        private void ToolStripMenuItem2_Click(object sender, EventArgs e)
        {

        }

        private void ToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.ShowDialog();
        }

        private void contextMenuStrip_cell_Opening(object sender, CancelEventArgs e)
        {
            if (worksheet.FocusPos.ToAddress() == "K13") {
                // contextMenuStrip_cell.Items[0].Visible = false;
                e.Cancel = true;
            } else
            {
                contextMenuStrip_cell.Items[0].Visible = true;
            }
        }
    }
}
