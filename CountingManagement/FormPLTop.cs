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
            if (e.KeyCode == unvell.ReoGrid.Interaction.KeyCode.Delete)
            {
                // DeleteキーでCellDataChangedが発生しないので
                unvell.ReoGrid.Events.CellEventArgs eventArgs = new unvell.ReoGrid.Events.CellEventArgs(worksheet.GetCell(worksheet.FocusPos));
                Sheet_CellDataChanged(sender, eventArgs);
            }
        }

        private void Sheet_CellDataChanged(object sender, unvell.ReoGrid.Events.CellEventArgs e)
        {
            switch (worksheet.GetNameByRange(e.Cell.Address))
            {
                case "売上_計画_1":

                    worksheet.CellDataChanged -= Sheet_CellDataChanged;

                    if (worksheet.GetCell("売上総利益_計画_1").Style.BackColor != CELL_BACK_COLOR_EDIT_ONLY)
                    {
                        worksheet["売上総利益_計画_1"] = ConvertCellData(worksheet["売上_計画_1"]) - ConvertCellData(worksheet["売上原価_計画_1"]);
                    }
                    if (worksheet.GetCell("営業利益_計画_1").Style.BackColor != CELL_BACK_COLOR_EDIT_ONLY)
                    {
                        worksheet["営業利益_計画_1"] = ConvertCellData(worksheet["売上総利益_計画_1"]) - ConvertCellData(worksheet["販管費_計画_1"]);
                    }
                    if (worksheet.GetCell("経常利益_計画_1").Style.BackColor != CELL_BACK_COLOR_EDIT_ONLY)
                    {
                        worksheet["経常利益_計画_1"] = ConvertCellData(worksheet["営業利益_計画_1"]) + ConvertCellData(worksheet["営業外収益_計画_1"]) - ConvertCellData(worksheet["営業外費用_計画_1"]);
                    }
                    if (worksheet.GetCell("税引前当期純利益_計画_1").Style.BackColor != CELL_BACK_COLOR_EDIT_ONLY)
                    {
                        worksheet["税引前当期純利益_計画_1"] = ConvertCellData(worksheet["経常利益_計画_1"]) + ConvertCellData(worksheet["特別利益_計画_1"]) - ConvertCellData(worksheet["特別損失_計画_1"]);
                    }
                    if (worksheet.GetCell("当期純利益_計画_1").Style.BackColor != CELL_BACK_COLOR_EDIT_ONLY)
                    {
                        worksheet["当期純利益_計画_1"] = ConvertCellData(worksheet["税引前当期純利益_計画_1"]) - ConvertCellData(worksheet["法人税等_計画_1"]);
                    }

                    worksheet.CellDataChanged += Sheet_CellDataChanged;


                    break;
                case "売上総利益_計画_1":

                    Decimal calValue = ConvertCellData(worksheet["売上_計画_1"]) - ConvertCellData(worksheet["売上原価_計画_1"]);
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
                    worksheet.SetRangeStyles("売上総利益_計画_1", style);

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
