using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;

namespace DataTableToXml
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //foreach (var excProc in Process.GetProcessesByName("Excel"))
            //{
            //    try
            //    {
            //        excProc.Kill(); // Завершаем процесс
            //        excProc.WaitForExit(); // Ждем полного закрытия
            //    }
            //    catch
            //    {
            //        // Игнорируем ошибки, например, если процесс нельзя закрыть
            //    }
            //}
            DataTable myTable = new DataTable();

            myTable.Columns.Add("Имя", typeof(string));
            myTable.Columns.Add("Возраст", typeof(int));
            myTable.Columns.Add("Город", typeof(string));
            myTable.Columns.Add("Дата рождения", typeof(DateTime));
            myTable.Columns.Add("Студент", typeof(bool));
            myTable.Columns.Add("Баллы", typeof(double));

            myTable.Rows.Add("Алексей", 30, "Минск", new DateTime(1994, 5, 10), true, 87.5);
            myTable.Rows.Add("Мария", 25, "Гомель", new DateTime(1999, 11, 23), false, 92.3);
            myTable.Rows.Add("Иван", 28, "Брест", new DateTime(1996, 2, 3), true, 78.0);
            myTable.Rows.Add("Ольга", 22, "Витебск", new DateTime(2002, 8, 17), false, 95.8);
            myTable.Rows.Add("Никита", 35, "Гродно", new DateTime(1989, 1, 30), true, 65.4);
            myTable.Rows.Add("Елена", 29, "Могилёв", new DateTime(1995, 3, 12), false, 88.1);
            myTable.Rows.Add("Дмитрий", 40, "Гомель", new DateTime(1983, 7, 5), true, 72.9);
            myTable.Rows.Add("Светлана", 27, "Минск", new DateTime(1996, 12, 1), false, 94.7);
            myTable.Rows.Add("Владимир", 31, "Брест", new DateTime(1992, 9, 14), true, 81.3);
            myTable.Rows.Add("Наталья", 24, "Витебск", new DateTime(2000, 6, 25), false, 90.5);
            myTable.Rows.Add("Андрей", 33, "Гродно", new DateTime(1990, 4, 18), true, 68.7);
            myTable.Rows.Add("Ирина", 26, "Минск", new DateTime(1997, 11, 30), true, 77.4);
            myTable.Rows.Add("Олег", 38, "Могилёв", new DateTime(1985, 1, 20), false, 83.6);

            ExcelExporter.ExportDataTableToExcel(myTable, @"C:\\Users\\zimnitskyaa\\Desktop\\test.xlsx");

            //Открытие CSV в Excel
            Process.Start(new ProcessStartInfo()
            {
                FileName = @"C:\\Users\\zimnitskyaa\\Desktop\\test.xlsx",
                UseShellExecute = true
            });
        }
    }

    public class ExcelExporter
    {
        public static void ExportDataTableToExcel(DataTable table, string filePath)
        {
            // Создаём новый Excel-документ
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                // Добавляем часть книги (WorkbookPart) и инициализируем Workbook
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // Добавляем часть листа (WorksheetPart) и создаём структуру данных листа (SheetData)
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                // Добавляем коллекцию листов и создаём лист с именем "Лист1"
                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart), // связываем лист с частью
                    SheetId = 1,                                  // идентификатор листа
                    Name = "Лист1"                                // имя листа в Excel
                };
                sheets.Append(sheet);

                // Добавляем стили к книге (шрифты, заливки, границы, форматы ячеек)
                AddStyles(workbookPart);

                // Формируем строку заголовков таблицы
                Row headerRow = new Row();

                for (int i = 0; i < table.Columns.Count; i++)
                {
                    uint styleIndex;

                    if (i == 0)
                    {
                        // Левая ячейка заголовка — стиль с толстой левой границей
                        styleIndex = 6;
                    }
                    else if (i == table.Columns.Count - 1)
                    {
                        // Правая ячейка заголовка — стиль с толстой правой границей
                        styleIndex = 7;
                    }
                    else
                    {
                        // Внутренние ячейки — стиль с тонкими верхней и нижней границами
                        styleIndex = 5;
                    }

                    Cell cell = new Cell
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(table.Columns[i].ColumnName),
                        StyleIndex = styleIndex
                    };
                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                Columns columns = new Columns();
                columns.Append(new Column()
                {
                    Min = 4,
                    Max = 4,
                    Width = 15,
                    CustomWidth = true
                });

                worksheetPart.Worksheet.InsertAt(columns, 0);

                // Заполняем строки данными из DataTable
                foreach (DataRow dtRow in table.Rows)
                {
                    Row newRow = new Row();

                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        object item = dtRow[i];
                        Cell cell = new Cell();

                        // В зависимости от типа данных устанавливаем значение и стиль ячейки
                        if (item is int || item is float || item is double || item is decimal)
                        {
                            // Числа: записываем как строку с форматом числа
                            cell.CellValue = new CellValue(Convert.ToString(item, CultureInfo.InvariantCulture));
                            cell.StyleIndex = (int)CellTypes.Number; // стиль для чисел
                        }
                        else if (item is DateTime dateTime)
                        {
                            // Даты в Excel хранятся как числа (OADate)
                            cell.CellValue = new CellValue(dateTime.ToOADate().ToString(CultureInfo.InvariantCulture));
                            cell.DataType = CellValues.Number;
                            cell.StyleIndex = (uint)CellTypes.Date;
                        }
                        else if (item is bool b)
                        {
                            // Булевы значения: тип Boolean и стиль 
                            cell.DataType = CellValues.Boolean;
                            cell.CellValue = new CellValue(b ? "1" : "0");
                            cell.StyleIndex = (int)CellTypes.Boolean;
                        }
                        else
                        {
                            // Текст и прочее: строковый тип, стиль для текста 
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(item?.ToString() ?? string.Empty);
                            cell.StyleIndex = (int)CellTypes.Text;
                        }

                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }

                // Сохраняем книгу
                workbookPart.Workbook.Save();
            }
        }

        private static void AddStyles(WorkbookPart workbookPart)
        {
            // Добавляем новую часть стилей (WorkbookStylesPart) в книгу
            WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();

            // Определяем стили: шрифты, заливки, границы и форматы ячеек
            stylesPart.Stylesheet = new Stylesheet(
                // Определяем шрифты
                new Fonts(
                    new DocumentFormat.OpenXml.Spreadsheet.Font(), // 0 - обычный шрифт
                    new DocumentFormat.OpenXml.Spreadsheet.Font(new Bold())           // 1 - жирный шрифт
                ),

                // Определяем заливки (фоны ячеек)
                new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }),     // 0 - без заливки
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 })  // 1 - стандартная сетка Excel
                ),

                // Определяем границы ячеек
                new Borders(
                    new Border(), // 0 - без границ
                    new Border(   // 1 - тонкие границы со всех сторон
                        new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new TopBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new BottomBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } }
                    ),
                    new Border(   // 2 - толстые границы сверху и снизу
                        new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new TopBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                        new BottomBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } }
                    ),
                    new Border(   // 3 - толстые границы сверху снизу и слева
                        new LeftBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                        new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new TopBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                        new BottomBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } }
                    ),
                    new Border(   // 4 - толстые границы сверху снизу и справа
                        new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new RightBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                        new TopBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                        new BottomBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } }
                    )
                ),

                // Определяем форматы ячеек
                new CellFormats(
                    new CellFormat(),                 // 0 - по умолчанию (без стиля)
                    new CellFormat { FontId = 0, BorderId = 1, ApplyBorder = true },    // 1 - текст: обычный шрифт, зелёный фон, границы
                    new CellFormat { FontId = 0, NumberFormatId = 4, BorderId = 1, ApplyBorder = true, ApplyNumberFormat = true }, // 2 - число с форматом чисел (NumberFormatId 4 — число с 2 знаками после запятой)
                    new CellFormat { FontId = 0, NumberFormatId = 14, BorderId = 1, ApplyBorder = true, ApplyNumberFormat = true }, // 3 - дата (формат даты)
                    new CellFormat { FontId = 1, ApplyFont = true,  BorderId = 1, ApplyBorder = true }, // 4 - булевы значения 
                    new CellFormat { BorderId = 2, ApplyBorder = true }, // 5 - сверху снизу толстая
                    new CellFormat { BorderId = 3, ApplyBorder = true }, // 6 - левая толстая
                    new CellFormat { BorderId = 4, ApplyBorder = true } // 7 - правая толстая
                )
            );

            // Сохраняем стили
            stylesPart.Stylesheet.Save();
        }

        // Перечисление для удобства использования индексов стилей
        public enum CellTypes
        {
            Default = 0,
            Text,
            Number,
            Date,
            Boolean
        }
        public enum BorderTypes
        {
            Default = 0,
            Header = 1,
            Text = 2,
            Number = 3,
            Date = 4,
            Boolean = 5
        }
    }
}