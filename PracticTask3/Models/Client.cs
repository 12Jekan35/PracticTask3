using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PracticTask3.Models
{
    public class Client : IViewItem
    {
        public int Id { get; private set; }
        public string OrganizationName { get; private set; }
        public string ContactPerson { get; private set; }
        public string Address { get; private set; }
        public static string Header { get => " Организация | Адрес | Контактное лицо (ФИО) |"; }
        virtual public IEnumerable<Request> Requests { get; set; }

        public Client(int id, string organizationName, string contactPerson, string address)
        {
            Id = id;
            OrganizationName = organizationName;
            ContactPerson = contactPerson;
            Address = address;
        }

        public Client(string[] values)
        {
            Id = int.Parse(values[0]);
            OrganizationName = values[1];
            ContactPerson = values[2];
            Address = values[3];
        }

        public override string ToString()
        {
            return OrganizationName + "|" + Address + "|" + ContactPerson + "|";
        }


        public Client ChangeClientData(string organizationName = "", string contactPerson = "", string address = "")
        {
            OrganizationName = !string.IsNullOrEmpty(organizationName)? organizationName: OrganizationName;
            ContactPerson = !string.IsNullOrEmpty(contactPerson) ? contactPerson : ContactPerson;
            Address = !string.IsNullOrEmpty(address) ? address : Address;

            return this;
        }

        public static string View(IEnumerable<object> list)
        {
            if (list is IEnumerable<Client> clientList)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("                 Клиенты");
                sb.AppendLine(Header);
                foreach (var client in clientList)
                {
                    sb.AppendLine(client.ToString());
                }
                return sb.ToString();
            }
            return string.Empty;
        }
    }
}
