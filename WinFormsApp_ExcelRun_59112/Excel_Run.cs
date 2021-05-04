using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WinFormsApp_ExcelRun_59112
{
    public class Excel_Run
    {
        private Excel.Application _xls_App = null;
        private Excel.Workbook _xls_Book = null;
        private Excel.Sheets _xls_SetOfSheets = null;
        private Excel.Worksheet _xls_Sheet = null;
        private Excel.Range _xls_Range = null;
        private int _numOfSheets = 0;
        private bool _access = false;

        public void Excel_Close()
        {
            if (_xls_App != null)
            {
                _access = false;
                _xls_Book.Close(0);
                _xls_App.Quit();
                _xls_Book = null;
                _xls_App = null;
                GC.Collect();
            }
            else
            {
                MessageBox.Show(null, "Файл уже был закрыт или не был создан!",
                    "Информация о приложении Excel", 0);
            }
        }

        public bool Excel_New()
        {
            if (_access && _xls_App != null)
            {
                Excel_Close();
            }
            _xls_App = new Excel.Application();
            if (_xls_App != null)
            {
                _xls_Book = _xls_App.Workbooks.Add();
                if (_xls_Book != null)
                {
                    _xls_SetOfSheets = _xls_Book.Worksheets;
                    _numOfSheets = _xls_SetOfSheets.Count;
                    _access = true;
                }
                else
                {
                    MessageBox.Show(null, "Книга Excel не была инициализирована!", "Ошибка инициализации", 0);
                }
            }
            return _access;
        }

        public bool Excel_Visible(bool flag_visible)
        {
            bool is_show = false;
            if (_access == true && _xls_App != null)
            {
                _xls_App.Visible = flag_visible;
                is_show = true;
            }
            else
            {
                MessageBox.Show(null, "Приложение Excel не было запущено!",
        "Ошибка доступа к свойству видимости Excel", 0);
            }
            return is_show;
        }

        public bool Excel_Write_Formula(int num_of_sheet, string str_cell, string str_set)
        {
            bool is_write = false;
            if (_access && _xls_App != null && _xls_Book != null && _xls_SetOfSheets
        != null && num_of_sheet > 0 && num_of_sheet <= _numOfSheets)
            {
                _xls_Sheet = _xls_SetOfSheets.get_Item(num_of_sheet);
                _xls_Range = _xls_Sheet.get_Range(str_cell, str_cell);
                if (_xls_Range != null)
                {
                    _xls_Range.Select();
                    _xls_Range.Formula = str_set;
                    is_write = true;
                }
                else
                {
                    MessageBox.Show(null, "Запись данных невозможно! Нет доступа к указанной ячейке", "Ошибка записи формулы", 0);
                }
            }
            else
            {
                MessageBox.Show(null, "Запись данных невозможна! Указанный лист недоступен", "Ошибка записи формулы", 0);
            }
            return is_write;
        }

        public bool Excel_Write_Cell(int num_of_sheet, string str_cell, string str_set)
        {
            bool is_write = false;
            if (_access && _xls_App != null && _xls_Book != null && _xls_SetOfSheets != null
            && num_of_sheet > 0 && num_of_sheet <= _numOfSheets)

            {
                _xls_Sheet = _xls_SetOfSheets.get_Item(num_of_sheet);
                _xls_Range = _xls_Sheet.get_Range(str_cell, str_cell);

                if (_xls_Range != null)
                {
                    _xls_Range.Value2 = str_set;
                    is_write = true;
                }
                else
                {
                    MessageBox.Show(null, "Запись данных невозможна! Нет доступа к указанной ячейке", "Ошибка записи ячейки", 0);
                }
            }
            else
            {
                MessageBox.Show(null, "Запись данных невозможна! Указанный лись недоступен", "Ошибка записи ячейки", 0);
            }
            return is_write;
        }

        public bool Excel_Read_Cell(int num_of_sheet, string str_cell, ref string str_get)
        {
            bool is_read = false;
            if (_access && _xls_App != null && _xls_Book != null &&
      _xls_SetOfSheets != null && num_of_sheet > 0 && num_of_sheet <= _numOfSheets)
            {
                _xls_Sheet = _xls_SetOfSheets.get_Item(num_of_sheet);
                _xls_Range = _xls_Sheet.get_Range(str_cell, str_cell);
                if (_xls_Range != null)
                {
                    _xls_Range.Select();
                    dynamic value = _xls_Range.Value2;
                    if (value == null)
                    {
                        str_get = "";
                    }
                    else
                    {
                        str_get = value.ToString();
                    }
                    is_read = true;
                }
                else
                {
                    MessageBox.Show(null, "Чтение данных невозможно! Нет доступа к указанной ячейки", "Ошибка чтения ячейки", 0);
                }
            }
            else
            {
                MessageBox.Show(null, " Чтение данных невозможно! Указанный лист недоступен", "Ошибка чтения ячейки", 0);
            }
            return is_read;
        }
    }
}
