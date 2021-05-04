using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = WinFormsApp_ExcelRun_59112.Excel_Run;

namespace WinFormsApp_ExcelRun_59112
{
    public partial class form_Main : Form
    {
        private Excel _excel = null;
        public form_Main()
        {
            InitializeComponent();
            _excel = new Excel();
        }

        private void button_RunExcel_Click(object sender, EventArgs e)
        {
            _excel.Excel_New();
        }

        private void button_WriteColumn_Click(object sender, EventArgs e)
        {
            int count_cells = listBox_InputCell.Items.Count;
            int count_data = listBox_InputData.Items.Count;
            if (count_cells != count_data)
            {
                MessageBox.Show(null, "Число ячеек и число элементов GUI не совпадают!", "Ошибка ввода данных", 0);
                return;
            }
            for (int i = 0; i < count_cells; i++)
            {
                string str_cell = listBox_InputCell.Items[i].ToString();
                string str_write = listBox_InputData.Items[i].ToString();
                bool is_set = _excel.Excel_Write_Cell(1, str_cell, str_write);
                if (!is_set)
                {
                    break;
                }
            }
        }

        private void button_WriteFormula_Click(object sender, EventArgs e)
        {
            string str_cell = textBox_CellFormula.Text;
            string str_data = textBox_InputFormula.Text;
            _excel.Excel_Write_Formula(1, str_cell, str_data);
            str_data = "";
            _excel.Excel_Read_Cell(1, str_cell, ref str_data);
            textBox_ResultFormula.Text = str_data;
        }

        private void button_ReadCell_Click(object sender, EventArgs e)
        {
            string str_cell = textBox_CellData.Text;
            string str_data = "";
            _excel.Excel_Read_Cell(1, str_cell, ref str_data);
            textBox_ResultData.Text = str_data;
        }

        private void button_ShowExcel_Click(object sender, EventArgs e)
        {
            _excel.Excel_Visible(true);
        }

        private void button_Exit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button_Close_Excel_Click(object sender, EventArgs e)
        {
            _excel.Excel_Close();
        }

        private void button_HideExcel_Click(object sender, EventArgs e)
        {
            _excel.Excel_Visible(false);
        }
    }
}
