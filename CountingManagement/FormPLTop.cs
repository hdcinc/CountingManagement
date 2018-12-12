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

        private bool CanCellInput(Cell cell)
        {
            switch (worksheet.GetNameByRange(cell.Address))
            {
                case "売上_計画_1":
                case "売上原価_計画_1":
                case "売上総利益_計画_1":
                case "販管費_計画_1":
                case "営業利益_計画_1":
                case "営業外収益_計画_1":
                case "営業外費用_計画_1":
                case "経常利益_計画_1":
                case "特別利益_計画_1":
                case "特別損失_計画_1":
                case "税引前当期純利益_計画_1":
                case "法人税等_計画_1":
                case "当期純利益_計画_1":

                    return true;
            }

            return false;
        }

        private void Sheet_CellDataChanged(object sender, unvell.ReoGrid.Events.CellEventArgs e)
        {

            // ===================================================================
            // 入力可能なセルか判断する
            // ===================================================================

            if (!CanCellInput(e.Cell))
            {
                return;
            }

            // ===================================================================
            // セル名を取得する
            // ===================================================================

            string[] names = worksheet.GetNameByRange(e.Cell.Address).Split('_');

            // ===================================================================
            // セルの入力値によって入力されたセルの色を変える
            // ===================================================================

            //Decimal calValue = 0;
            //Decimal newValue = ConvertCellData(e.Cell.Data);

            //switch (worksheet.GetNameByRange(e.Cell.Address))
            //{
            //    case "売上総利益_計画_1":

            //        calValue = ConvertCellData(worksheet.GetCell("売上_計画_1").Data) - ConvertCellData(worksheet.GetCell("売上原価_計画_1").Data);

            //        break;

            //    case "営業利益_計画_1":

            //        calValue = ConvertCellData(worksheet.GetCell("売上総利益_計画_1").Data) - ConvertCellData(worksheet.GetCell("販管費_計画_1").Data);

            //        break;

            //    case "経常利益_計画_1":

            //        calValue = ConvertCellData(worksheet.GetCell("営業利益_計画_1").Data) + ConvertCellData(worksheet.GetCell("営業外収益_計画_1").Data) - ConvertCellData(worksheet.GetCell("営業外費用_計画_1").Data);

            //        break;

            //}

            //e.Cell.Style.BackColor = (calValue == newValue) ? CELL_BACK_COLOR_EDIT_ABLE : CELL_BACK_COLOR_EDIT_ONLY;



            switch (worksheet.GetNameByRange(e.Cell.Address))
            {
                case "売上総利益_計画_1":
                case "営業利益_計画_1":
                case "経常利益_計画_1":

                    Decimal calValue = 0;
                    Decimal newValue = ConvertCellData(e.Cell.Data);

                    switch (worksheet.GetNameByRange(e.Cell.Address))
                    {
                        case "売上総利益_計画_1":

                            calValue = ConvertCellData(worksheet.GetCell("売上_計画_1").Data) - ConvertCellData(worksheet.GetCell("売上原価_計画_1").Data);

                            break;

                        case "営業利益_計画_1":

                            calValue = ConvertCellData(worksheet.GetCell("売上総利益_計画_1").Data) - ConvertCellData(worksheet.GetCell("販管費_計画_1").Data);

                            break;

                        case "経常利益_計画_1":

                            calValue = ConvertCellData(worksheet.GetCell("営業利益_計画_1").Data) + ConvertCellData(worksheet.GetCell("営業外収益_計画_1").Data) - ConvertCellData(worksheet.GetCell("営業外費用_計画_1").Data);

                            break;

                    }

                    e.Cell.Style.BackColor = (calValue == newValue) ? CELL_BACK_COLOR_EDIT_ABLE : CELL_BACK_COLOR_EDIT_ONLY;

                    break;

            }



            // ===================================================================
            // 他のセルの値を再計算する
            // ===================================================================

            switch (worksheet.GetNameByRange(e.Cell.Address))
            {
                case "売上_計画_1":
                case "売上原価_計画_1":
                case "売上総利益_計画_1":
                case "販管費_計画_1":
                case "営業利益_計画_1":
                case "営業外収益_計画_1":
                case "営業外費用_計画_1":
                case "経常利益_計画_1":
                case "特別利益_計画_1":
                case "特別損失_計画_1":
                case "税引前当期純利益_計画_1":
                case "法人税等_計画_1":
                case "当期純利益_計画_1":

                    // セルに値を設定する間、本イベントを発生しないようにする（セルに値を設定するたびに本イベントが発生する）
                    worksheet.CellDataChanged -= Sheet_CellDataChanged;

                    if (worksheet.GetCell("売上総利益_計画_1").Style.BackColor != CELL_BACK_COLOR_EDIT_ONLY)
                    {
                        worksheet["売上総利益_計画_1"] = ConvertCellData(worksheet.GetCell("売上_計画_1").Data) - ConvertCellData(worksheet.GetCell("売上原価_計画_1").Data);
                    }
                    if (worksheet.GetCell("営業利益_計画_1").Style.BackColor != CELL_BACK_COLOR_EDIT_ONLY)
                    {
                        worksheet["営業利益_計画_1"] = ConvertCellData(worksheet.GetCell("売上総利益_計画_1").Data) - ConvertCellData(worksheet.GetCell("販管費_計画_1").Data);
                    }
                    if (worksheet.GetCell("経常利益_計画_1").Style.BackColor != CELL_BACK_COLOR_EDIT_ONLY)
                    {
                        worksheet["経常利益_計画_1"] = ConvertCellData(worksheet.GetCell("営業利益_計画_1").Data) + ConvertCellData(worksheet.GetCell("営業外収益_計画_1").Data) - ConvertCellData(worksheet.GetCell("営業外費用_計画_1").Data);
                    }
                    if (worksheet.GetCell("税引前当期純利益_計画_1").Style.BackColor != CELL_BACK_COLOR_EDIT_ONLY)
                    {
                        worksheet["税引前当期純利益_計画_1"] = ConvertCellData(worksheet.GetCell("経常利益_計画_1").Data) + ConvertCellData(worksheet.GetCell("特別利益_計画_1").Data) - ConvertCellData(worksheet.GetCell("特別損失_計画_1").Data);
                    }
                    if (worksheet.GetCell("当期純利益_計画_1").Style.BackColor != CELL_BACK_COLOR_EDIT_ONLY)
                    {
                        worksheet["当期純利益_計画_1"] = ConvertCellData(worksheet.GetCell("税引前当期純利益_計画_1").Data) - ConvertCellData(worksheet.GetCell("法人税等_計画_1").Data);
                    }

                    // セルに値を設定し終わったので、本イベントを発生するようにする
                    worksheet.CellDataChanged += Sheet_CellDataChanged;

                    break;
                default:
                    break;
            }

        }

        private Decimal Calculate売上総利益(string planResult, int businessYearColumn)
        {
            return ConvertCellData(worksheet.GetCell("売上_計画_1").Data) - ConvertCellData(worksheet.GetCell("売上原価_計画_1").Data);
        }

        private Decimal ConvertCellData(object cellData)
        {
            Decimal value = 0;

            Decimal.TryParse(Convert.ToString(cellData), out value);

            return value;
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
