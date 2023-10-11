using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AkelonTask3
{
    public class Product
    {
        public int Code { get; set; }
        public string Name { get; set; }
        public string Unit { get; set; }
        public decimal UnitPrice { get; set; }
        public Product(int code, string name, string unit, decimal unitPrice)
        {
            Code = code;
            Name = name;
            Unit = unit;
            UnitPrice = unitPrice;
        }



    }

}
