using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using WpfApp1.Models;

namespace WpfApp1.ViewModels
{
    internal class MainViewModels
    {
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

        public void CreateExcelTemplate(string name, string address, string start, string finish, string credit, DatePicker dataTimePicker)
        {
            string price = GetPrice(start, finish);
            string summ = GetSumm(price, "33,76");
            string finalSumm = GetFinal(summ, credit);

            string monthStr = "";
            string year = "";
            DateTime ? selectedDate = dataTimePicker.SelectedDate;
            if (selectedDate.HasValue)
            {
                int monthInt = selectedDate.Value.Month;
                monthStr = GetMonth(monthInt);
                year = selectedDate.Value.Year.ToString();
            }
            Models.Models models = new Models.Models();
            models.GenerateExcelTemplate(name, address, start, finish, credit, price, summ, finalSumm, monthStr, year, dataTimePicker);
            
        }
    }
}
