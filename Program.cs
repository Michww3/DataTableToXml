using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;

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
                    Name = "Отчёт"                                // имя листа в Excel
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
                        styleIndex = (int)HeaderPosition.Left;
                    }
                    else if (i == table.Columns.Count - 1)
                    {
                        // Правая ячейка заголовка — стиль с толстой правой границей
                        styleIndex = (int)HeaderPosition.Right;
                    }
                    else
                    {
                        // Внутренние ячейки — стиль с тонкими верхней и нижней границами
                        styleIndex = (int)HeaderPosition.Inner;
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
                    Min = 1,
                    Max = (uint)table.Columns.Count,
                    Width = 15,
                    CustomWidth = true
                });

                worksheetPart.Worksheet.InsertAt(columns, 0);

                int columnsCount = table.Columns.Count; // количество столбцов
                int headerRowIndex = 1; // строка заголовков, обычно 1

                // Формируем адрес диапазона с заголовками
                string startColumn = "A";
                string endColumn = GetExcelColumnName(columnsCount);
                string filterRange = $"{startColumn}{headerRowIndex}:{endColumn}{headerRowIndex}";

                // Добавляем или обновляем элемент AutoFilter
                var worksheet = worksheetPart.Worksheet;

                var autoFilter = worksheet.Elements<AutoFilter>().FirstOrDefault();
                if (autoFilter != null)
                {
                    autoFilter.Reference = filterRange;
                }
                else
                {
                    autoFilter = new AutoFilter() { Reference = filterRange };
                    worksheet.InsertAfter(autoFilter, worksheet.Elements<SheetData>().First());
                }

                // Сохраняем изменения
                worksheet.Save();

                // Вспомогательная функция для преобразования номера столбца в букву Excel
                string GetExcelColumnName(int columnNumber)
                {
                    int dividend = columnNumber;
                    string columnName = string.Empty;
                    int modulo;

                    while (dividend > 0)
                    {
                        modulo = (dividend - 1) % 26;
                        columnName = Convert.ToChar(65 + modulo) + columnName;
                        dividend = (dividend - modulo) / 26;
                    }

                    return columnName;
                }

                // Создаём или получаем SheetViews
                SheetViews sheetViews = worksheet.Elements<SheetViews>().FirstOrDefault();
                if (sheetViews == null)
                {
                    sheetViews = new SheetViews();
                    worksheet.InsertAt(sheetViews, 0);
                }

                // Создаём SheetView
                SheetView sheetView = sheetViews.Elements<SheetView>().FirstOrDefault();
                if (sheetView == null)
                {
                    sheetView = new SheetView() { WorkbookViewId = 0U };
                    sheetViews.Append(sheetView);
                }

                // Создаём панель для закрепления (закрепляем первую строку — значит панель начинается со второй строки, Index 1)
                Pane pane = new Pane()
                {
                    VerticalSplit = 1D,       // число строк, которые будут закреплены сверху (1 — первая строка)
                    TopLeftCell = "A2",       // Ячейка, которая будет в верхнем левом углу видимой области после закрепления
                    ActivePane = PaneValues.BottomLeft, // Активная панель после закрепления
                    State = PaneStateValues.Frozen     // Устанавливаем состояние "закреплено"
                };

                // Удаляем старую панель, если есть
                var oldPane = sheetView.Elements<Pane>().FirstOrDefault();
                if (oldPane != null)
                    oldPane.Remove();

                sheetView.Append(pane);

                // Сохраняем изменения
                worksheet.Save();

                for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
                {
                    var dtRow = table.Rows[rowIndex];
                    Row newRow = new Row();

                    for (int colIndex = 0; colIndex < table.Columns.Count; colIndex++)
                    {
                        object item = dtRow[colIndex];
                        Cell cell = new Cell();

                        // Определяем тип значения и его строковое представление
                        string valueText = item?.ToString() ?? string.Empty;
                        cell.CellValue = new CellValue(valueText);

                        Type dataType = item?.GetType();
                        bool isNumeric = dataType == typeof(int) || dataType == typeof(double) || dataType == typeof(float) || dataType == typeof(decimal);
                        bool isDate = dataType == typeof(DateTime);

                        if (isNumeric)
                        {
                            cell.CellValue = new CellValue(Convert.ToString(item, CultureInfo.InvariantCulture));
                            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }
                        else if (isDate)
                        {
                            DateTime dateValue = (DateTime)item;
                            cell.CellValue = new CellValue(dateValue.ToOADate().ToString(CultureInfo.InvariantCulture));
                            cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }
                        else
                        {
                            // Строка
                            cell.DataType = new EnumValue<CellValues>(CellValues.String);
                            cell.CellValue = new CellValue(valueText);
                        }

                        // Позиция ячейки в таблице
                        CellPosition position = GetCellPosition(rowIndex, table.Rows.Count, colIndex, table.Columns.Count);

                        // Выбор стиля в зависимости от позиции и типа данных
                        cell.StyleIndex = GetStyleIndex(position, isNumeric, isDate);

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
                    new Border(   // 1 - толстые границы сверху и снизу
                        new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new TopBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                        new BottomBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } }
                    ),
                    new Border(   // 2 - толстые границы сверху снизу и слева
                        new LeftBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                        new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new TopBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                        new BottomBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } }
                    ),
                    new Border(   // 3 - толстые границы сверху снизу и справа
                        new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new RightBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                        new TopBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                        new BottomBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } }
                    ),
                    new Border(   // 4 - тонкие границы со всех сторон
                        new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new TopBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new BottomBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } }
                    ),
                    // 5 - левая верхняя ячейка
                    new Border(
                        new LeftBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                        new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new TopBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                        new BottomBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } }
                    ),

                    // 6 - левая средняя ячейка
                    new Border(
                        new LeftBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                        new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new TopBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new BottomBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } }
                    ),

                    // 7 - левая нижняя ячейка
                    new Border(
                        new LeftBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                        new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new TopBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new BottomBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } }
                    ),

                    // 8 - нижняя ячейка
                    new Border(
                        new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new TopBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new BottomBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } }
                    ),

                    // 9 - правая верхняя ячейка
                    new Border(
                        new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new RightBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                        new TopBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                        new BottomBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } }
                    ),

                    // 10 - правая средняя ячейка
                    new Border(
                        new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new RightBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                        new TopBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new BottomBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } }
                    ),

                    // 11 - правая нижняя ячейка
                    new Border(
                        new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new RightBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                        new TopBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new BottomBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } }
                    ),

                    // 12 - верхняя ячейка
                    new Border(
                        new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } },
                        new TopBorder { Style = BorderStyleValues.Thick, Color = new Color { Auto = true } },
                        new BottomBorder { Style = BorderStyleValues.Thin, Color = new Color { Auto = true } }
                    )
                ),

                //TO DO header all border thick
                new CellFormats(
                    new CellFormat(),                                         // 0 - по умолчанию (без стиля)

                    //header
                    new CellFormat { BorderId = 1, ApplyBorder = true },      // 1 - заголовок (сверху и снизу толстая)
                    new CellFormat { BorderId = 2, ApplyBorder = true },      // 2 - заголовок: левая ячейка
                    new CellFormat { BorderId = 3, ApplyBorder = true },      // 3 - заголовок: правая ячейка

                    //default
                    new CellFormat { BorderId = 4, ApplyBorder = true },      // 4 - ячейка внутри таблицы (тонкие границы)
                    new CellFormat { BorderId = 5, ApplyBorder = true },      // 5 - левая верхняя ячейка
                    new CellFormat { BorderId = 6, ApplyBorder = true },      // 6 - левая средняя ячейка
                    new CellFormat { BorderId = 7, ApplyBorder = true },      // 7 - левая нижняя ячейка
                    new CellFormat { BorderId = 8, ApplyBorder = true },      // 8 - нижняя внутренняя 
                    new CellFormat { BorderId = 9, ApplyBorder = true },      // 9 - правая верхняя ячейка
                    new CellFormat { BorderId = 10, ApplyBorder = true },     // 10 - правая средняя ячейка
                    new CellFormat { BorderId = 11, ApplyBorder = true },     // 11 - правая нижняя ячейка
                    new CellFormat { BorderId = 12, ApplyBorder = true },      // 12 - верхняя внутренняя ячейка

                    // --- 13–18: Number
                    new CellFormat { BorderId = 4, ApplyBorder = true, NumberFormatId = 4, ApplyNumberFormat = true },  // 13 - Middle Number
                    new CellFormat { BorderId = 5, ApplyBorder = true, NumberFormatId = 4, ApplyNumberFormat = true },  // 14 - LeftTop Number
                    new CellFormat { BorderId = 6, ApplyBorder = true, NumberFormatId = 4, ApplyNumberFormat = true },  // 15 - LeftMiddle Number
                    new CellFormat { BorderId = 7, ApplyBorder = true, NumberFormatId = 4, ApplyNumberFormat = true },  // 16 - LeftBottom Number
                    new CellFormat { BorderId = 8, ApplyBorder = true, NumberFormatId = 4, ApplyNumberFormat = true },  // 17 - Bottom Number
                    new CellFormat { BorderId = 9, ApplyBorder = true, NumberFormatId = 4, ApplyNumberFormat = true },  // 18 - RightTop Number
                    new CellFormat { BorderId = 10, ApplyBorder = true, NumberFormatId = 4, ApplyNumberFormat = true }, // 19 - RightMiddle Number
                    new CellFormat { BorderId = 11, ApplyBorder = true, NumberFormatId = 4, ApplyNumberFormat = true }, // 20 - RightBottom Number
                    new CellFormat { BorderId = 12, ApplyBorder = true, NumberFormatId = 4, ApplyNumberFormat = true }, // 21 - Top Number

                    // --- 22–30: Date
                    new CellFormat { BorderId = 4, ApplyBorder = true, NumberFormatId = 14, ApplyNumberFormat = true },  // 22 - Middle Date
                    new CellFormat { BorderId = 5, ApplyBorder = true, NumberFormatId = 14, ApplyNumberFormat = true },  // 23 - LeftTop Date
                    new CellFormat { BorderId = 6, ApplyBorder = true, NumberFormatId = 14, ApplyNumberFormat = true },  // 24 - LeftMiddle Date
                    new CellFormat { BorderId = 7, ApplyBorder = true, NumberFormatId = 14, ApplyNumberFormat = true },  // 25 - LeftBottom Date
                    new CellFormat { BorderId = 8, ApplyBorder = true, NumberFormatId = 14, ApplyNumberFormat = true },  // 26 - Bottom Date
                    new CellFormat { BorderId = 9, ApplyBorder = true, NumberFormatId = 14, ApplyNumberFormat = true },  // 27 - RightTop Date
                    new CellFormat { BorderId = 10, ApplyBorder = true, NumberFormatId = 14, ApplyNumberFormat = true }, // 28 - RightMiddle Date
                    new CellFormat { BorderId = 11, ApplyBorder = true, NumberFormatId = 14, ApplyNumberFormat = true }, // 29 - RightBottom Date
                    new CellFormat { BorderId = 12, ApplyBorder = true, NumberFormatId = 14, ApplyNumberFormat = true } // 30 - Top Date
                )
            );

            // Сохраняем стили
            stylesPart.Stylesheet.Save();
        }

        private static CellPosition GetCellPosition(int rowIndex, int totalRows, int colIndex, int totalCols)
        {
            bool isTop = rowIndex == 0;
            bool isBottom = rowIndex == totalRows - 1;
            bool isLeft = colIndex == 0;
            bool isRight = colIndex == totalCols - 1;

            if (isTop && isLeft) return CellPosition.LeftTop;
            if (isTop && isRight) return CellPosition.RightTop;
            if (isBottom && isLeft) return CellPosition.LeftBottom;
            if (isBottom && isRight) return CellPosition.RightBottom;
            if (isTop) return CellPosition.Top;
            if (isBottom) return CellPosition.Bottom;
            if (isLeft) return CellPosition.LeftMiddle;
            if (isRight) return CellPosition.RightMiddle;

            return CellPosition.Inner;
        }

        private static uint GetStyleIndex(CellPosition position, bool isNumeric, bool isDate)
        {
            int baseIndex = 0;

            switch (position)
            {
                case CellPosition.Inner:
                    baseIndex = 0;
                    break;
                case CellPosition.LeftTop:
                    baseIndex = 1;
                    break;
                case CellPosition.LeftMiddle:
                    baseIndex = 2;
                    break;
                case CellPosition.LeftBottom:
                    baseIndex = 3;
                    break;
                case CellPosition.Bottom:
                    baseIndex = 4;
                    break;
                case CellPosition.RightTop:
                    baseIndex = 5;
                    break;
                case CellPosition.RightMiddle:
                    baseIndex = 6;
                    break;
                case CellPosition.RightBottom:
                    baseIndex = 7;
                    break;
                case CellPosition.Top:
                    baseIndex = 8;
                    break;
                default:
                    baseIndex = 0;
                    break;
            }

            if (isNumeric)
                return (uint)(13 + baseIndex);  // Number block
            if (isDate)
                return (uint)(22 + baseIndex);  // Date block

            return (uint)(4 + baseIndex);       // Text/default
        }

    }

    public enum HeaderPosition
    {
        Inner = 1,
        Left,
        Right
    }

    public enum CellPosition
    {
        Inner = 4,
        LeftTop,
        LeftMiddle,
        LeftBottom,
        Bottom,
        RightTop,
        RightMiddle,
        RightBottom,
        Top
    }
}