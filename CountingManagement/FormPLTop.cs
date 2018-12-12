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

        public enum NameInTheCell
        {
            /// <summary>
            /// 勘定科目
            /// </summary>
            AccountTitle,
            /// <summary>
            /// 予実
            /// </summary>
            PlanResult,
            /// <summary>
            /// 年度列番号
            /// </summary>
            BusinessYearColumn
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

            string[] splitedNames = worksheet.GetNameByRange(e.Cell.Address).Split('_');

            Dictionary<NameInTheCell, string> names = new Dictionary<NameInTheCell, string>();
            names.Add(NameInTheCell.AccountTitle, splitedNames[0]);
            names.Add(NameInTheCell.PlanResult, splitedNames[1]);
            names.Add(NameInTheCell.BusinessYearColumn, splitedNames[2]);

            // ===================================================================
            // セルの入力値によって入力されたセルの色を変える
            // ===================================================================

            switch (names[NameInTheCell.AccountTitle])
            {
                case "売上総利益":
                case "営業利益":
                case "経常利益":

                    Decimal calValue = 0;
                    Decimal newValue = ConvertCell(e.Cell.Data);

                    switch (names[NameInTheCell.AccountTitle])
                    {
                        case "売上総利益":

                            calValue = Calculate売上総利益(names[NameInTheCell.PlanResult], names[NameInTheCell.BusinessYearColumn]);
                            break;

                        case "営業利益":

                            calValue = Calculate営業利益(names[NameInTheCell.PlanResult], names[NameInTheCell.BusinessYearColumn]);
                            break;

                        case "経常利益":

                            calValue = Calculate経常利益(names[NameInTheCell.PlanResult], names[NameInTheCell.BusinessYearColumn]);
                            break;

                    }

                    e.Cell.Style.BackColor = (calValue == newValue) ? CELL_BACK_COLOR_EDIT_ABLE : CELL_BACK_COLOR_EDIT_ONLY;

                    break;

            }

            // ===================================================================
            // 他のセルの値を再計算する
            // ===================================================================

            switch (names[NameInTheCell.AccountTitle])
            {
                case "売上":
                case "売上原価":
                case "売上総利益":
                case "販管費":
                case "営業利益":
                case "営業外収益":
                case "営業外費用":
                case "経常利益":
                case "特別利益":
                case "特別損失":
                case "税引前当期純利益":
                case "法人税等":
                case "当期純利益":

                    // セルに値を設定する間、本イベントを発生しないようにする（セルに値を設定するたびに本イベントが発生する）
                    worksheet.CellDataChanged -= Sheet_CellDataChanged;

                    string tail = "_" + names[NameInTheCell.PlanResult] + "_" + names[NameInTheCell.BusinessYearColumn];

                    if (worksheet.GetCell("売上総利益"+tail).Style.BackColor != CELL_BACK_COLOR_EDIT_ONLY)
                    {
                        worksheet["売上総利益"+tail] = Calculate売上総利益(names[NameInTheCell.PlanResult], names[NameInTheCell.BusinessYearColumn]);
                    }
                    if (worksheet.GetCell("営業利益"+tail).Style.BackColor != CELL_BACK_COLOR_EDIT_ONLY)
                    {
                        worksheet["営業利益"+tail] = Calculate営業利益(names[NameInTheCell.PlanResult], names[NameInTheCell.BusinessYearColumn]);
                    }
                    if (worksheet.GetCell("経常利益"+tail).Style.BackColor != CELL_BACK_COLOR_EDIT_ONLY)
                    {
                        worksheet["経常利益"+tail] = Calculate経常利益(names[NameInTheCell.PlanResult], names[NameInTheCell.BusinessYearColumn]);
                    }
                    if (worksheet.GetCell("税引前当期純利益"+tail).Style.BackColor != CELL_BACK_COLOR_EDIT_ONLY)
                    {
                        worksheet["税引前当期純利益" + tail] = Calculate税引前当期純利益(names[NameInTheCell.PlanResult], names[NameInTheCell.BusinessYearColumn]);
                    }
                    if (worksheet.GetCell("当期純利益"+tail).Style.BackColor != CELL_BACK_COLOR_EDIT_ONLY)
                    {
                        worksheet["当期純利益" + tail] = Calculate当期純利益(names[NameInTheCell.PlanResult], names[NameInTheCell.BusinessYearColumn]);
                    }

                    // セルに値を設定し終わったので、本イベントを発生するようにする
                    worksheet.CellDataChanged += Sheet_CellDataChanged;

                    break;
                default:
                    break;
            }

        }

        private Decimal Calculate売上総利益(string planResult, string businessYearColumn)
        {

            string tail = "_" + planResult + "_" + businessYearColumn;

            return ConvertCell(worksheet.GetCell("売上" + tail).Data) - ConvertCell(worksheet.GetCell("売上原価" + tail).Data);
        }

        private Decimal Calculate営業利益(string planResult, string businessYearColumn)
        {

            string tail = "_" + planResult + "_" + businessYearColumn;

            return ConvertCell(worksheet.GetCell("売上総利益" + tail).Data) - ConvertCell(worksheet.GetCell("販管費" + tail).Data);
        }

        private Decimal Calculate経常利益(string planResult, string businessYearColumn)
        {

            string tail = "_" + planResult + "_" + businessYearColumn;

            return ConvertCell(worksheet.GetCell("営業利益" + tail).Data) + ConvertCell(worksheet.GetCell("営業外収益" + tail).Data) - ConvertCell(worksheet.GetCell("営業外費用" + tail).Data);
        }

        private Decimal Calculate税引前当期純利益(string planResult, string businessYearColumn)
        {

            string tail = "_" + planResult + "_" + businessYearColumn;

            return ConvertCell(worksheet.GetCell("経常利益"+tail).Data) + ConvertCell(worksheet.GetCell("特別利益"+tail).Data) - ConvertCell(worksheet.GetCell("特別損失"+tail).Data);
        }

        private Decimal Calculate当期純利益(string planResult, string businessYearColumn)
        {

            string tail = "_" + planResult + "_" + businessYearColumn;

            return ConvertCell(worksheet.GetCell("税引前当期純利益"+tail).Data) - ConvertCell(worksheet.GetCell("法人税等"+tail).Data);
        }

        private Decimal ConvertCell(object cellData)
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
