using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AkelonTask3
{
    public class Request
    {
        public int RequestCode { get; set; }
        public int ProductCode { get; set; }
        public int ClientCode { get; set; }
        public int NumberRequest { get; set; }
        public int CountOfProducts { get; set; }

        public DateTime PostingDate = DateTime.Now;
        public Request(int requestCode, int productCode, int clientCode, int numberRequest, int countOfProducts, string postingDate)
        {
            RequestCode = requestCode;
            ProductCode = productCode;
            ClientCode = clientCode;
            NumberRequest = numberRequest;
            CountOfProducts = countOfProducts;
            double serialDate = double.Parse(postingDate);
            DateTime dateTime = DateTime.FromOADate(serialDate);
            PostingDate = DateTime.Parse(dateTime.ToString("yyyy-MM-dd"));
        }
        public override string ToString()
        {
            return $"Код заявки:{this.RequestCode}" +
                "\t" + $"Код товара:{this.ProductCode}" +
                "\t" + $"Код клиента:{this.ClientCode}" +
                "\t" + $"Номер заявки:{this.NumberRequest}" +
                "\t" + $"Требуемое количество:{this.CountOfProducts}" +
                "\t" + $"Дата размещения:{this.PostingDate.ToString("d")}";
        }
    }
}
