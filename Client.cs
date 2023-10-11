using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AkelonTask3
{
    public class Client
    {
        public int ClientCode { get; set; }
        public string OrgName { get; set; }
        public string Address { get; set; }
        public string ContactFullName { get; set; }

        public Client(int clientCode, string orgName, string address, string contactFullName)
        {
            ClientCode = clientCode;
            OrgName = orgName;
            Address = address;
            ContactFullName = contactFullName;
        }

        public override string ToString()
        {
            return $"Код клиента:{this.ClientCode} | Имя организации:{this.OrgName} | Адрес:{this.Address} | Контактное лицо:{this.ContactFullName}";
        }
    }
}
