using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using PracticTask3.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;
using DocumentFormat.OpenXml.Office2016.Excel;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace PracticTask3
{
    static class Program
    {
        private static string filename = string.Empty;
        private static bool isLoaded = false;
        private static List<Product> products = new List<Product>();
        private static List<Client> clients = new List<Client>();
        private static List<Request> requests = new List<Request>();
        public static void Main()
        {
            Console.WriteLine("Добро пожаловать!");
            while (true)
            {
                Console.WriteLine("Для выбора пункта меню напишите номер пункта");

                if (string.IsNullOrEmpty(filename))
                {
                    Console.WriteLine("1. Загрузить документ формата XLSX;\n2. Выйти из приложения.");
                    Console.Write("Опция: ");
                    var answer = Console.ReadLine();
                    char option = answer.FirstOrDefault(c => char.IsDigit(c));


                    if (option == '1')
                    {
                        filename = GetFilepath();
                        Console.Clear();
                        if (!string.IsNullOrEmpty(filename))
                            Console.WriteLine("Документ загружен");

                        continue;
                    }
                    else if (option == '2')
                    {
                        break;
                    }
                    else
                    {
                        GetError("Не был выбран ни один из пунктов меню!");
                        continue;
                    }
                }
                else
                {
                    Console.WriteLine(ViewAllDataFromDocument());
                    string options = "1. Вывести информацию о клиентах по наименованию товара;\n" +
                                     "2. Изменить контактное лицо клиента;\n" +
                                     "3. Вывести \"золотого\" клиента;\n" +
                                     "4. Вернуться в главное меню.\n";
                    Console.WriteLine(options);
                    Console.Write("Опция: ");
                    var answer = Console.ReadLine();
                    char option = answer.FirstOrDefault(c => char.IsDigit(c));


                    if (option == '1')
                    {
                        Console.Write("Введите наименование товара:");
                        string productName = Console.ReadLine();

                        if (string.IsNullOrEmpty(productName))
                        {
                            GetError("Некорректное значение!");
                            continue;
                        }
                        var product = products.Where(p => p.Name == productName).FirstOrDefault();
                        if (product == null)
                        {
                            GetError("Товар не найден");
                            continue;
                        }
                        var clientsId = requests.Where(r => r.ProductId == product.Id).Select(r => r.ClientId);
                        if (clientsId.Count() < 1)
                        {
                            GetError("Клиентов по данному товару не найдено");
                            continue;
                        }
                        var productClients = clients.Where(c => clientsId.Contains(c.Id)).ToList();
                        Console.Clear();
                        Console.WriteLine("Клиенты по товару " + productName + ":");
                        foreach (var client in productClients)
                        {
                            Console.WriteLine(client);
                        }
                        Console.Write("Нажмите любую кнопку, чтобы вернуться в меню...");
                        Console.ReadKey();
                        continue;
                    }
                    else if (option == '2')
                    {
                        Console.WriteLine("Выберите номер строки, которую хотите изменить");
                        for (int i = 0; i < clients.Count(); i++)
                        {
                            Console.WriteLine((i+1) + ". |" + clients[i]);
                        }
                        
                        int index = 0;
                        Console.Write("Введите номер строки:");
                        if (!int.TryParse(Console.ReadLine(), out index) && (index >= clients.Count || index < 0))
                        {
                            GetError("Некорректное значение!");
                            continue;
                        }
                        index--;

                        var client = clients[index];
                        Console.Clear();
                        Console.WriteLine("Текущие данные клиента:");
                        Console.WriteLine(client);
                        Console.WriteLine();
                        Console.WriteLine("Введите поочерёдно значения для каждой из столбцов.\n" +
                                          "(Если вы не планируете менять значение, то просто нажмите Enter)");
                        string organizationName, address, contactPerson = "";
                        Console.Write("Наименование организации: ");
                        organizationName = Console.ReadLine();
                        Console.Write("Адрес: ");
                        address = Console.ReadLine();
                        Console.Write("Контактное лицо (ФИО): ");
                        contactPerson = Console.ReadLine();

                        client.ChangeClientData(organizationName, contactPerson, address);
                        bool result = false;
                        try
                        {
                            // index+2 потому что мы пропускаем строку с заголовками столбцов
                            result = TryRewriteRowInSheet(filename, "Клиенты", index+2, client.ToString());
                        }
                        catch (Exception ex)
                        {
                            GetError("Ошибка записи изменений в документ");
                            continue;
                        }
                        if (result)
                        {
                            Console.Clear();
                            Console.WriteLine("Новые данные клиента:\n" + client);
                            Console.Write("Нажмите любую кнопку, чтобы вернуться в меню...");
                            Console.ReadKey();
                            continue;
                        }
                        else
                        {
                            GetError("Не удалось внести изменения в документ");
                            continue;
                        }
                    }
                    else if (option == '3')
                    {
                        Console.WriteLine("Выберите режим выборки из заявок:\n1. За год\n2. За месяц");
                        Console.Write("Введите номер опции: ");
                        int index = 0;
                        if (!int.TryParse(Console.ReadLine(), out index) && (index > 2 || index < 1))
                        {
                            GetError("Некорректное значение!");
                            continue;
                        }
                        ReportType reportType;
                        DateTime date;
                        switch (index)
                        {
                            case 1:
                                reportType = ReportType.Yearly; 
                                break;
                            case 2:
                                reportType = ReportType.Monthly; 
                                break;
                            default: 
                                reportType = ReportType.Yearly;
                                break;
                        }
                        if (reportType == ReportType.Yearly)
                        {
                            Console.Write("Введите год по которому вести выборку: ");
                            int year = 0;
                            if (!int.TryParse(Console.ReadLine(), out year) && ( year < 1 && year > DateTime.Today.Year))
                            {
                                GetError("Некорректное значение!");
                                continue;
                            }
                            date = new DateTime(year, 1,1);
                        }
                        else
                        {
                            Console.Write("Введите год по которому вести выборку: ");
                            int year = 0;
                            if (!int.TryParse(Console.ReadLine(), out year) && (year < 1 && year > DateTime.Today.Year))
                            {
                                GetError("Некорректное значение!");
                                continue;
                            }
                            Console.Write("Введите месяц по которому вести выборку: ");
                            int month = 0;
                            if (!int.TryParse(Console.ReadLine(), out month) && (month < 1 && month > 12))
                            {
                                GetError("Некорректное значение!");
                                continue;
                            }
                            date = new DateTime(year, month, 1);
                        }
                        
                        var client = GetGoldenClientByDate(clients, requests, reportType, date);
                        if (client == null)
                        {
                            GetError("Error");
                            continue;
                        }
                        Console.Clear();
                        Console.WriteLine("Золотой клиент по вашему запросу:\n" + client);
                        Console.Write("Нажмите любую кнопку, чтобы вернуться в меню...");
                        Console.ReadKey();
                        continue;

                    }
                    else if (option == '4')
                    {
                        filename = null;
                        products.Clear();
                        clients.Clear();
                        requests.Clear();
                        isLoaded = false;
                        Console.Clear();
                        continue;
                    }
                    else
                    {
                        GetError("Не был выбран ни один из пунктов меню!");
                        continue;
                    }
                }

            }
        }

        enum ReportType
        {
            Yearly,
            Monthly
        }

        private static Client GetGoldenClientByDate(IEnumerable<Client> clients, IEnumerable<Request> requests, ReportType reportType, DateTime date)
        {
            Client goldenClient = null;
            if (clients == null || requests == null) { return goldenClient; }
            requests = requests.Where(r =>
            {
                bool correct = false;
                switch (reportType)
                {
                    case ReportType.Yearly:
                        correct = r.PlacementDate.Year == date.Year;
                        break;
                    case ReportType.Monthly:
                        correct = r.PlacementDate.Year == date.Year && 
                        r.PlacementDate.Month == date.Month;
                        break;
                }
                return correct;
            }).ToList();
            int maxCountRequests = 0;
            foreach (Client client in clients)
            {
                int count = requests.Where(r => r.ClientId == client.Id).Count();
                if (count > maxCountRequests) 
                { 
                    maxCountRequests = count;
                    goldenClient = client;
                }
            }
            return goldenClient;
        }
        private static bool LoadDataFromDocument(string filename)
        {
            using (var doc = SpreadsheetDocument.Open(filename, false))
            {
                var workbook = doc.WorkbookPart;
                var sharedStringTable = workbook.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>();
                string currentTable = "";
                foreach (var worksheet in workbook.Workbook.Descendants<Sheet>())
                {
                    currentTable = worksheet.Name;
                    var sheet = (WorksheetPart)workbook.GetPartById(worksheet.Id);
                    foreach (var row in sheet.Worksheet.Descendants<Row>())
                    {
                        if (row.RowIndex == 1) { continue; }
                        switch (currentTable)
                        {
                            case "Товары":
                                {
                                    var values = new string[4];
                                    foreach (var cell in row.Elements<Cell>())
                                    {
                                        if (cell.CellReference.InnerText.Contains("A"))
                                        {
                                            values[0] = cell.CellValue?.InnerText ?? string.Empty;
                                        }
                                        else if (cell.CellReference.InnerText.Contains("B"))
                                        {
                                            values[1] = GetTextFromCell(cell, sharedStringTable);
                                        }
                                        else if (cell.CellReference.InnerText.Contains("C"))
                                        {
                                            values[2] = GetTextFromCell(cell, sharedStringTable);
                                        }
                                        else if (cell.CellReference.InnerText.Contains("D"))
                                        {
                                            values[3] = cell.CellValue?.InnerText ?? string.Empty;
                                        }
                                    }
                                    if (string.IsNullOrEmpty(values[0])) break;
                                    products.Add(new Product(values));
                                }
                                break;
                            case "Клиенты":
                                {
                                    var values = new string[6];
                                    foreach (var cell in row.Elements<Cell>())
                                    {
                                        if (cell.CellReference.InnerText.Contains("A"))
                                        {
                                            values[0] = cell.CellValue?.InnerText ?? string.Empty;
                                        }
                                        else if (cell.CellReference.InnerText.Contains("B"))
                                        {
                                            values[1] = GetTextFromCell(cell, sharedStringTable);
                                        }
                                        else if (cell.CellReference.InnerText.Contains("C"))
                                        {
                                            values[2] = GetTextFromCell(cell, sharedStringTable);
                                        }
                                        else if (cell.CellReference.InnerText.Contains("D"))
                                        {
                                            values[3] = GetTextFromCell(cell, sharedStringTable);
                                        }
                                    }
                                    if (string.IsNullOrEmpty(values[0])) break;
                                    clients.Add(new Client(values));
                                }
                                break;
                            case "Заявки":
                                {
                                    var values = new string[6];
                                    foreach (var cell in row.Elements<Cell>())
                                    {
                                        if (cell.CellReference.InnerText.Contains("A"))
                                        {
                                            values[0] = cell.CellValue?.InnerText ?? string.Empty;
                                        }
                                        else if (cell.CellReference.InnerText.Contains("B"))
                                        {
                                            values[1] = cell.CellValue?.InnerText ?? string.Empty;
                                        }
                                        else if (cell.CellReference.InnerText.Contains("C"))
                                        {
                                            values[2] = cell.CellValue?.InnerText ?? string.Empty;
                                        }
                                        else if (cell.CellReference.InnerText.Contains("D"))
                                        {
                                            values[3] = cell.CellValue?.InnerText ?? string.Empty;
                                        }
                                        else if (cell.CellReference.InnerText.Contains("E"))
                                        {
                                            values[4] = cell.CellValue?.InnerText ?? string.Empty;
                                        }
                                        else if (cell.CellReference.InnerText.Contains("F"))
                                        {
                                            values[5] = cell.CellValue?.InnerText ?? string.Empty;
                                        }
                                    }
                                    if (string.IsNullOrEmpty(values[0])) break;
                                    var request = new Request(values);
                                    request.Product = products.FirstOrDefault(p => p.Id == request.ProductId);
                                    request.Client = clients.FirstOrDefault(c => c.Id == request.ClientId);
                                    requests.Add(request);
                                }
                                break;
                            default: break;
                        }

                    }
                }
                isLoaded = true;
            }
            return isLoaded;
        }
        private static bool TryRewriteRowInSheet(string filename, string sheetName, int rowIndex, string data)
        {
            if (string.IsNullOrEmpty(data)) { return false; }
            string[] values = data.Split('|');
            try
            {
                using (var doc = SpreadsheetDocument.Open(filename, true))
                {
                    var workbook = doc.WorkbookPart;
                    var sharedStringTable = workbook.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>();
                    var worksheetParts = workbook.WorksheetParts;
                    var worksheet = workbook.Workbook.Descendants<Sheet>().FirstOrDefault(sh => sh.Name == sheetName);
                    if (worksheet == null) { return false; }

                    var sheet = (WorksheetPart)workbook.GetPartById(worksheet.Id);
                    var row = sheet.Worksheet.Descendants<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
                    if (row == null) { return false;}

                    var cells = row.Elements<Cell>().ToList();
                    for (int i= 0; i < values.Length; i++)
                    {
                        if (!string.IsNullOrEmpty(values[i]) || cells.Count <= i+1)
                            break;
                        cells[i+1].CellValue = new CellValue(values[i]);
                    }
                }
            }
            catch
            {
                return false;
            }
            return true;
        }
        private static void GetError(string errorText)
        {
            Console.Clear();
            Console.WriteLine(errorText);
        }

        private static string ViewAllDataFromDocument()
        {
            try
            {
                if (!isLoaded)
                {
                    LoadDataFromDocument(filename);
                    if (!isLoaded) return string.Empty;
                }

                StringBuilder sb = new StringBuilder();
                sb.AppendLine(Product.View(products));
                sb.AppendLine(Client.View(clients));
                sb.AppendLine(Request.View(requests));
                return sb.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine("При открытии/чтении файла возникла ошибка!");
            }
            return string.Empty;
        }
        private static string GetTextFromCell(Cell cell, IEnumerable<SharedStringItem> sharedStrings)
        {
            int id = 0;
            if (int.TryParse(cell.CellValue?.InnerText, out id))
            {
                var value = sharedStrings.ElementAt(id);
                return value.Text.Text ?? value.InnerText ?? string.Empty;
            }
            return string.Empty;
        }

        private static string GetFilepath()
        {
            Console.Write("\nВведите путь к файлу:");
            string filename = Console.ReadLine();

            if (!string.IsNullOrEmpty(filename) && File.Exists(filename) && filename.EndsWith(".xlsx"))
            {
                return filename;
            }
            return string.Empty;
        }
    }
}