﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Drawing;
using System.Diagnostics;

namespace WpfApp1.Models
{
    internal class Models
    {
        private Worksheet RemoveCellBorders(Worksheet worksheet, int rowIndex, int columnIndex)
        {
            // Получение ячейки, у которой нужно удалить границы
            Cell cell = worksheet.Cells[rowIndex, columnIndex];
            // Удаление границ ячейки
            Aspose.Cells.Style style = cell.GetStyle();
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.None;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.None;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.None;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.None;
            cell.SetStyle(style);
            return worksheet;
        }

        private Worksheet ButtomCellBorder(Worksheet worksheet, int rowIndex, int columnIndex)
        {
            // Получение ячейки, у которой нужно удалить границы
            Cell cell = worksheet.Cells[rowIndex, columnIndex];
            // Удаление границ ячейки
            Aspose.Cells.Style style = cell.GetStyle();
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.None;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.None;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.None;
            cell.SetStyle(style);
            return worksheet;
        }
        private Worksheet ThinCellBorder(Worksheet worksheet, int rowIndex, int columnIndex)
        {
            // Получение ячейки, у которой нужно удалить границы
            Cell cell = worksheet.Cells[rowIndex, columnIndex];
            // Удаление границ ячейки
            Aspose.Cells.Style style = cell.GetStyle();
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            cell.SetStyle(style);
            return worksheet;
        }

        private Worksheet CellsBorder(Worksheet worksheet, int firstRow, int firstColumn, int lastRow, int lastColumn, string style)
        {
            for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++)
            {
                for (int colIndex = firstColumn; colIndex <= lastColumn; colIndex++)
                {
                    switch (style)
                    {
                        case "BottomThin":
                            worksheet = ButtomCellBorder(worksheet, rowIndex, colIndex);
                            break;
                        case "RemoveBorder":
                            worksheet = RemoveCellBorders(worksheet, rowIndex, colIndex);
                            break;
                        case "ThinBorder":
                            worksheet = ThinCellBorder(worksheet, rowIndex, colIndex);
                            break;
                    }
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

        public void GenerateExcelTemplate(string name,string address, string start, string finish, string credit, string price, string summ, string finalSumm, string monthStr, string year, DatePicker dataTimePicker)
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
    }
}
