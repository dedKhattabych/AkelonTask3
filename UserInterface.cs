using AkelonTask3;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AkelonTask3
{
    public static class UserInterface
    {
        public static void Start()
        {
            string excelFilePath = string.Empty;
            List<Product> products = null;
            List<Client> clients = null;
            List<Request> requests = null;

            while (true)
            {
                Console.WriteLine("Введите путь файла");
                excelFilePath = Console.ReadLine();// C:\\Users\\user\\Documents\\4.xlsx
                if (!File.Exists(excelFilePath))
                {
                    Console.WriteLine("Файл не найден.");
                    continue;
                }
                if (!Path.GetExtension(excelFilePath).Equals(".xlsx"))
                {
                    Console.WriteLine("Некорректное расширение файла, пожалуйста введите путь до файла xlsx");
                    continue;
                }
                if (!ExcelManager.StartParseProduct(excelFilePath, out products))
                {
                    return;
                }
                if (!ExcelManager.StartParseClients(excelFilePath, out clients))
                {
                    return;
                }
                if (!ExcelManager.StartParseRequests(excelFilePath, out requests))
                {
                    return;
                }


                break;
            }

            while (true)
            {
                Console.WriteLine("Введите число для дальнейшего действия:");
                Console.WriteLine("1 - Вывести информацию о клиентах, заказавших конкретный товар.");
                Console.WriteLine("2 - Изменить контактное лицо организации.");
                Console.WriteLine("3 - Определить золотого клиента.");
                Console.WriteLine("4 - Выход из приложения.");

                var inputIsNumber = Int32.TryParse(Console.ReadLine(), out int commandNumber);
                if (!inputIsNumber)
                {
                    Console.WriteLine("Введите корректное число");
                    continue;
                }
                switch (commandNumber)
                {
                    case 1:
                        GetClientsFromProduct(products, clients, requests);
                        break;
                    case 2:
                        SetClientNameByOrg(excelFilePath);
                        break;
                    case 3:
                        GetGoldClientFromYearAndMonth(clients, requests);
                        break;
                    case 4:
                        return;
                    default:
                        Console.WriteLine("Введите корректное число");
                        continue;
                }
            }

        }
        /// <summary>
        /// Изменить контактное лицо организации.
        /// </summary>
        /// <param name="excelFilePath">Путь к файлу xlsx</param>
        private static void SetClientNameByOrg(string excelFilePath)
        {
            string orgInput;
            string FIOInput;
            while (true)
            {
                Console.WriteLine("Введите организацию");
                orgInput = Console.ReadLine();
                if (String.IsNullOrEmpty(orgInput))
                {
                    continue;
                }
                break;
            }
            while (true)
            {
                Console.WriteLine("Введите ФИО нового контактного лица.");
                FIOInput = Console.ReadLine();
                if (String.IsNullOrEmpty(FIOInput))
                {
                    continue;
                }
                break;
            }
            if (ExcelManager.UpdateClientContactNameByOrganization(excelFilePath, orgInput, FIOInput))
            {
                Console.WriteLine("Изменения внесены успешно.");
            }
            return;
        }
        /// <summary>
        /// Определить золотого клиента.
        /// </summary>
        /// <param name="clients">Список клиентов.</param>
        /// <param name="requests">Список заявок.</param>
        private static void GetGoldClientFromYearAndMonth(List<Client> clients, List<Request> requests)
        {
            Console.WriteLine("Введите год формата ХХХХ");
            int year;
            while (true)
            {
                var YearInput = Console.ReadLine();
                if (YearInput.Length == 4 && Int32.TryParse(YearInput, out year))
                {
                    break;
                }
            }

            Console.WriteLine("Введите месяц формата 1,2...12");
            int month;
            while (true)
            {
                var MonthInput = Console.ReadLine();
                if (Int32.TryParse(MonthInput, out month))
                {
                    if (month > 0 && month < 13)
                    {
                        break;
                    }
                }
            }
            var requestsByDate = requests.Where(x => x.PostingDate.Year == year && x.PostingDate.Month == month).ToList();

            var clientsBydate = new Dictionary<int, int>();

            foreach (var client in requestsByDate)
            {
                if (clientsBydate.ContainsKey(client.ClientCode))
                {
                    clientsBydate[client.ClientCode]++;
                    continue;
                }
                clientsBydate.Add(client.ClientCode, 1);
            }
            var maxKeyValue = clientsBydate.OrderByDescending(kv => kv.Value)
                                .First()
                                .Key;
            Console.WriteLine($"Клиент с наибольшим количество заказов за {year}.{month} это {clients.Where(x => x.ClientCode == maxKeyValue).ToList().First().OrgName}");
        }
        private static void GetClientsFromProduct(List<Product> products, List<Client> clients, List<Request> requests)
        {
            var productName = products.Select(x => x.Name.ToLower()).ToList();
            Console.WriteLine("Напишите наименование товара, чтобы получить информацию о клиентах, заказавших этот товар.");
            for (int i = 0; i < productName.Count; i++)
            {
                Console.WriteLine($"{productName[i]}");
            }

            while (true)
            {
                var productNameInput = Console.ReadLine().ToLower();

                if (!productName.Contains(productNameInput))
                {
                    Console.WriteLine("Товар не найден, напишите еще раз");
                    continue;
                }

                var product = products.Where(x => x.Name.ToLower() == productNameInput).ToList().FirstOrDefault();
                var requestsFromProduct = requests.Where(x => x.ProductCode == product.Code).ToList();
                for (int i = 0; i < requestsFromProduct.Count; i++)
                {
                    var clientFromRequest = clients.Where(x => x.ClientCode == requestsFromProduct[i].ClientCode).First();
                    Console.WriteLine("____________________________");
                    Console.WriteLine($"Клиент номер:{i}");
                    Console.WriteLine($"{clientFromRequest}");
                    Console.WriteLine($"{requestsFromProduct[i]}" + $" Цена: {product.UnitPrice}");
                    Console.WriteLine("____________________________");
                }
                break;
            }
        }
    }
}
