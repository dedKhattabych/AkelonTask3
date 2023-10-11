using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using AkelonTask3;

namespace AkelonTask3
{
    public static class ExcelManager
    {
        private const int CLIENT_TABLE_INDEX = 1;

        public static bool StartParseProduct(string excelFilePath, out List<Product> products)
        {
            products = new List<Product>();
            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelFilePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();


                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    Console.WriteLine("Лист: " + sheet.Name);

                    foreach (Row row in sheetData.Elements<Row>().Skip(1))
                    {
                        List<string> rowData = new List<string>();

                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            string cellValue = cell.InnerText;

                            if (cell.DataType != null)
                            {
                                int sharedStringIndex = int.Parse(cellValue);
                                SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
                                SharedStringTable sharedStringTable = sharedStringTablePart.SharedStringTable;
                                cellValue = sharedStringTable.ChildElements[sharedStringIndex].InnerText;
                                rowData.Add(cellValue);
                                continue;
                            }

                            rowData.Add(cellValue);

                        }
                        products.Add(new Product(Int32.Parse(rowData[0]), rowData[1], rowData[2], Decimal.Parse(rowData[3])));
                        Console.WriteLine(string.Join("\t", rowData));
                    }

                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex}");
                return false;
            }
            return true;
        }

        public static bool StartParseClients(string excelFilePath, out List<Client> clients)
        {
            clients = new List<Client>();
            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelFilePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    var sheets = workbookPart.Workbook.Descendants<Sheet>();
                    var sheet = sheets.ToList()[1];

                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    Console.WriteLine("Лист: " + sheet.Name);

                    foreach (Row row in sheetData.Elements<Row>().Skip(1))
                    {
                        List<string> rowData = new List<string>();

                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            string cellValue = cell.InnerText;

                            if (cell.DataType != null) 
                            {
                                int sharedStringIndex = int.Parse(cellValue);
                                SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
                                SharedStringTable sharedStringTable = sharedStringTablePart.SharedStringTable;
                                cellValue = sharedStringTable.ChildElements[sharedStringIndex].InnerText;
                                rowData.Add(cellValue);
                                continue;
                            }

                            rowData.Add(cellValue);

                        }
                        try
                        {
                            clients.Add(new Client(Int32.Parse(rowData[0]), rowData[1], rowData[2], rowData[3]));
                            Console.WriteLine(string.Join("\t", rowData));
                        }
                        catch (Exception)
                        {

                        }

                    }
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex}");
                return false;
            }
            return true;
        }

        public static bool StartParseRequests(string excelFilePath, out List<Request> requests)
        {
            requests = new List<Request>();
            int currentRow = 1;
            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelFilePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    var sheets = workbookPart.Workbook.Descendants<Sheet>();
                    var sheet = sheets.ToList()[2];

                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    Console.WriteLine("Лист: " + sheet.Name);

                    foreach (Row row in sheetData.Elements<Row>().Skip(1))
                    {
                        currentRow += 1;
                        List<string> rowData = new List<string>();

                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            string cellValue = cell.InnerText;

                            if (cell.DataType != null)
                            {
                                int sharedStringIndex = int.Parse(cellValue);
                                SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
                                SharedStringTable sharedStringTable = sharedStringTablePart.SharedStringTable;
                                cellValue = sharedStringTable.ChildElements[sharedStringIndex].InnerText;
                                rowData.Add(cellValue);
                                continue;
                            }

                            rowData.Add(cellValue);

                        }
                        try
                        {
                            var temprequest = new Request(
                                Int32.Parse(rowData[0]),
                                Int32.Parse(rowData[1]),
                                Int32.Parse(rowData[2]),
                                Int32.Parse(rowData[3]),
                                Int32.Parse(rowData[4]),
                                rowData[5]);
                            requests.Add(temprequest);
                            Console.WriteLine(temprequest.ToString());
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("___________________________________");
                            Console.WriteLine($"Не удалось считать значения в таблице заявки для строки {currentRow}");
                            continue;
                        }
                    }
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex}");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Обновить контактное лицо по организации.
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <param name="orgName"></param>
        /// <param name="newFIO"></param>
        /// <returns></returns>
        public static bool UpdateClientContactNameByOrganization(string excelFilePath, string orgName, string newFIO)
        {
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(excelFilePath, true))
                {
                    WorkbookPart workbookPart = document.WorkbookPart;
                    var sheets = workbookPart.Workbook.Descendants<Sheet>();
                    var sheet = sheets.ToList()[CLIENT_TABLE_INDEX];

                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    var rows = sheetData.Elements<Row>().Skip(1).ToList();

                    for (int i = 0; i < rows.Count; i++)
                    {
                        List<string> rowData = new List<string>();
                        var cell = rows[i].Elements<Cell>().Where(c => c.CellReference == $"B{i + 2}").FirstOrDefault();
                        if (cell == null)
                        {
                            continue;
                        }
                        string cellValue = cell.InnerText;

                        if (cell.DataType != null)
                        {
                            int sharedStringIndex = int.Parse(cellValue);
                            SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
                            SharedStringTable sharedStringTable = sharedStringTablePart.SharedStringTable;
                            cellValue = sharedStringTable.ChildElements[sharedStringIndex].InnerText;

                            if (!cellValue.ToLower().Equals(orgName.ToLower()))
                            {
                                continue;
                            }
                            int index = InsertSharedStringItem(newFIO, sharedStringTablePart);
                            Cell cell1 = InsertCellInWorksheet("D", (uint)i + 2, worksheetPart);
                            cell1.CellValue = new CellValue(index.ToString());
                            cell1.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                            worksheetPart.Worksheet.Save();

                            Console.WriteLine("Контактное лицо изменено.");
                            return true;
                        }
                    }
                }
                Console.WriteLine("Не найдена организация.");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Произошла ошибка во время работы с таблицей Клиенты:");
                Console.WriteLine(ex);
                return false;
            }
        }
        /// <summary>
        /// Вставляем я таблицу строк значение
        /// </summary>
        /// <param name="text">текст который нужно вставить</param>
        /// <param name="shareStringPart">Таблица строковых значений.</param>
        /// <returns></returns>
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        /// <summary>
        /// Вставить ячейку в табличку.
        /// </summary>
        /// <param name="columnName">Имя колонки</param>
        /// <param name="rowIndex">Индекс строки</param>
        /// <param name="worksheetPart">Таблица</param>
        /// <returns>Если ячейка существует, возвращаем.</returns>
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // Если в табличке нету строки с нужным индексом, вставляем.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }
            // Если нет ячейки по колонке, вставляем.
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

    }
}
