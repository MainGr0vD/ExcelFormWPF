using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WpfApp1.Models;

namespace WpfApp1.ViewModels
{

   

    public class PersonDataViewModel : INotifyPropertyChanged
    {

        private PersonDataModel _personData;

        public PersonDataModel PersonData
        {
            get { return _personData; }
            set
            {
                if (_personData != value)
                {
                    _personData = value;
                    OnPropertyChanged(nameof(PersonData));
                }
            }
        }

       

        private Worksheet CellBorders(Worksheet worksheet, int rowIndex, int columnIndex, string styleName)
        {
            // Получение ячейки, у которой нужно удалить границы
            Cell cell = worksheet.Cells[rowIndex, columnIndex];
            // Удаление границ ячейки
            Aspose.Cells.Style style = cell.GetStyle();
            switch (styleName)
            {
                case "RemoveCellBorders":
                    style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.None;
                    style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.None;
                    style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.None;
                    style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.None;
                    break;
                case "ButtomCellBorder":
                    style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.None;
                    style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.None;
                    style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.None;
                    break;
                case "ThinCellBorder":
                    style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                    style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                    break;
            }
            cell.SetStyle(style);
            return worksheet;
        }

        private Worksheet CellsBorder(Worksheet worksheet, int firstRow, int firstColumn, int lastRow, int lastColumn, string style)
        {
            for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++)
            {
                for (int colIndex = firstColumn; colIndex <= lastColumn; colIndex++)
                {
                    CellBorders(worksheet, rowIndex, colIndex, style);
                }
            }

            return worksheet;
        }

        private Worksheet MergeCells(Worksheet worksheet, int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
            string firstCellAddress = CellsHelper.CellIndexToName(firstRow, firstColumn);
            string lastCellAddress = CellsHelper.CellIndexToName(lastRow, lastColumn);
            Range range = worksheet.Cells.CreateRange(firstCellAddress, lastCellAddress);
            range.Merge();
            return worksheet;
        }


        private Worksheet SetBoldText(Worksheet worksheet, int rowIndex, int columnIndex)
        {
            // Получение ячейки, для которой нужно установить полужирный текст
            Cell cell = worksheet.Cells[rowIndex, columnIndex];
            // Установка полужирного шрифта для ячейки
            Aspose.Cells.Style style = cell.GetStyle();
            style.Font.IsBold = true;
            cell.SetStyle(style);
            return worksheet;
        }

        private Worksheet SetStylesCells(Worksheet worksheet, int firstRow, int firstColumn, int lastRow, int lastColumn, string[] style, string putValue)
        {
            for (int indexElement = 0; indexElement < style.Length; indexElement++)
            {
                switch (style[indexElement])
                {
                    case "RemoveBorder":
                        worksheet = CellsBorder(worksheet, firstRow, firstColumn, lastRow, lastColumn, "RemoveBorder");
                        break;
                    case "BottomThin":
                        worksheet = CellsBorder(worksheet, firstRow, firstColumn, lastRow, lastColumn, "RemoveBorder");
                        break;
                    case "ThinBorder":
                        worksheet = CellsBorder(worksheet, firstRow, firstColumn, lastRow, lastColumn, "ThinBorder");
                        break;
                    case "MergeCells":
                        worksheet = MergeCells(worksheet, firstRow, firstColumn, lastRow, lastColumn);
                        break;
                    case "SetBoldText":
                        worksheet = SetBoldText(worksheet, firstRow, firstColumn);
                        break;
                }
            }
            Cell cell = worksheet.Cells[firstRow, firstColumn];
            cell.PutValue(putValue);
            return worksheet;
        }

        public void GenerateExcelTemplate(string name, string address, string start, string finish, string credit, string price, string summ, string finalSumm, string monthStr, string year, DatePicker dataTimePicker)
        {

            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet = CellsBorder(worksheet, 1, 1, 10, 10, "RemoveBorder");
            worksheet = SetStylesCells(worksheet, 1, 1, 1, 3, new string[] { "RemoveBorder", "MergeCells", "SetBoldText" }, "СЧЕТ - ИЗВЕЩЕНИЯ за ");
            worksheet = SetStylesCells(worksheet, 1, 4, 1, 4, new string[] { "BottomThin", "SetBoldText" }, monthStr);
            worksheet = SetStylesCells(worksheet, 1, 5, 1, 5, new string[] { "BottomThin", "SetBoldText" }, year + " г.");
            worksheet = SetStylesCells(worksheet, 2, 1, 2, 1, new string[] { }, "Получатель: ");
            worksheet = SetStylesCells(worksheet, 2, 2, 2, 3, new string[] { "SetBoldText", "MergeCells" }, "МУП \"Архиповка\"");
            worksheet = SetStylesCells(worksheet, 3, 1, 3, 1, new string[] { "SetBoldText" }, "Абонент: ");
            worksheet = SetStylesCells(worksheet, 3, 2, 3, 5, new string[] { "MergeCells", "BottomThin", "SetBoldText" }, name);
            worksheet = SetStylesCells(worksheet, 4, 1, 4, 1, new string[] { }, "Адрес: ");
            worksheet = SetStylesCells(worksheet, 4, 2, 4, 5, new string[] { }, address);
            worksheet = CellsBorder(worksheet, 5, 1, 8, 6, "ThinBorder");
            worksheet = SetStylesCells(worksheet, 5, 1, 5, 1, new string[] { }, "Показ. Счетч.");
            worksheet = SetStylesCells(worksheet, 5, 2, 5, 2, new string[] { }, "Начальное");
            worksheet = SetStylesCells(worksheet, 5, 3, 5, 3, new string[] { }, "Конечное");
            worksheet = SetStylesCells(worksheet, 5, 4, 5, 4, new string[] { }, "Расход, куб.м");
            worksheet = SetStylesCells(worksheet, 5, 5, 5, 5, new string[] { }, "Цена");
            worksheet = SetStylesCells(worksheet, 5, 6, 6, 6, new string[] { }, "Сумма");

            worksheet = SetStylesCells(worksheet, 6, 1, 6, 1, new string[] { }, "ХВС");
            worksheet = SetStylesCells(worksheet, 6, 2, 6, 2, new string[] { }, start);
            worksheet = SetStylesCells(worksheet, 6, 3, 6, 3, new string[] { }, finish);
            worksheet = SetStylesCells(worksheet, 6, 4, 6, 4, new string[] { }, "33,76");
            worksheet = SetStylesCells(worksheet, 6, 5, 6, 5, new string[] { }, price);
            worksheet = SetStylesCells(worksheet, 6, 6, 6, 6, new string[] { }, summ);

            worksheet = SetStylesCells(worksheet, 7, 1, 7, 1, new string[] { }, "ГВС");
            worksheet = SetStylesCells(worksheet, 7, 2, 7, 2, new string[] { }, "");
            worksheet = SetStylesCells(worksheet, 7, 3, 7, 3, new string[] { }, "");
            worksheet = SetStylesCells(worksheet, 7, 4, 7, 4, new string[] { }, "");
            worksheet = SetStylesCells(worksheet, 7, 5, 7, 5, new string[] { }, "");
            worksheet = SetStylesCells(worksheet, 7, 6, 7, 6, new string[] { }, "");

            worksheet = SetStylesCells(worksheet, 8, 1, 8, 1, new string[] { }, "Водоотвед.");
            worksheet = SetStylesCells(worksheet, 8, 2, 8, 2, new string[] { }, "");
            worksheet = SetStylesCells(worksheet, 8, 3, 8, 3, new string[] { }, "");
            worksheet = SetStylesCells(worksheet, 8, 4, 8, 4, new string[] { }, "");
            worksheet = SetStylesCells(worksheet, 8, 5, 8, 5, new string[] { }, "");
            worksheet = SetStylesCells(worksheet, 8, 6, 8, 6, new string[] { }, "");

            worksheet = SetStylesCells(worksheet, 9, 1, 9, 2, new string[] { "MergeCells" }, "Задолж-ть (переплата)");
            worksheet = SetStylesCells(worksheet, 9, 3, 9, 3, new string[] { "BottomThin", "SetBoldText" }, credit);

            worksheet = SetStylesCells(worksheet, 9, 4, 9, 5, new string[] { "MergeCells", "SetBoldText" }, "Всего к оплате");
            worksheet = SetStylesCells(worksheet, 9, 6, 9, 6, new string[] { "BottomThin", "SetBoldText" }, finalSumm);
            // Сохранение книги Excel в файл
            workbook.Save("Template.xlsx");
        }
        private string GetOperation(string startStr, string finishStr, string metod)
        {
            double startInt;
            double finishInt;
            string price = "0";

            if (double.TryParse(startStr, out startInt) && double.TryParse(finishStr, out finishInt))
            {
                switch (metod)
                {
                    case "GetFinal":
                        price = (finishInt + startInt).ToString();
                        break;
                    case "GetPrice":
                        price = (finishInt - startInt).ToString();
                        break;
                    case "GetSumm":
                        price = (finishInt * startInt).ToString();
                        break;
                }

            }
            else
            {
                if (metod == "GetFinal")
                {
                    price = startStr;
                }
                else
                {
                    MessageBox.Show("Введите корректные данные!");
                }
            }
            return price;
        }

        private string GetFinal(string startStr, string finishStr) => GetOperation(startStr, finishStr, "GetFinal");

        private string GetPrice(string startStr, string finishStr) => GetOperation(startStr, finishStr, "GetPrice");

        private string GetSumm(string startStr, string finishStr) => GetOperation(startStr, finishStr, "GetSumm");

        private string GetMonth(int month)
        {
            switch (month)
            {
                case 1:
                    return "Январь";
                case 2:
                    return "Февраль";
                case 3:
                    return "Март";
                case 4:
                    return "Апрель";
                case 5:
                    return "Май";
                case 6:
                    return "Июнь";
                case 7:
                    return "Июль";
                case 8:
                    return "Август";
                case 9:
                    return "Сентябрь";
                case 10:
                    return "Октябрь";
                case 11:
                    return "Ноябрь";
                case 12:
                    return "Декабрь";
                default:
                    return "Некорректный номер месяца";
            }
        }

        private void ExportToExcel()
        {
            string price = GetPrice(PersonData.startPerson, PersonData.finishPerson);
            string summ = GetSumm(price, "33,76");
            string finalSumm = GetFinal(summ, PersonData.creditPerson);

            string monthStr = "";
            string year = "";
            DateTime ? selectedDate = PersonData.datePerson.SelectedDate;
            if (selectedDate.HasValue)
            {
                int monthInt = selectedDate.Value.Month;
                monthStr = GetMonth(monthInt);
                year = selectedDate.Value.Year.ToString();
            }
            GenerateExcelTemplate(PersonData.namePerson, PersonData.addressPerson, PersonData.startPerson, PersonData.finishPerson, PersonData.creditPerson, price, summ, finalSumm, monthStr, year, PersonData.datePerson);
            
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

    }
}
