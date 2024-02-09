using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PracticTask3.Models
{
    public class Request : IViewItem
    {
        public int Id { get; private set; }
        public int RequestNumber { get; private set; }
        public int ProductId { get; private set; }
        public int ClientId { get; private set; }
        public int RequiredAmount { get; private set; }
        public DateTime PlacementDate { get; private set; }

        virtual public Client Client { get; set; }
        virtual public Product Product { get; set; }
        public static string Header { get => " Товар | Клиент | Номер заявки | Требуемое количество | Дата размещения"; }
        public Request(int id, int requestNumber, int productId, int clientId, int requiredAmount, DateTime placementDate) 
        {
            Id = id;
            RequestNumber = requestNumber;
            ProductId = productId;
            ClientId = clientId;
            RequiredAmount = requiredAmount;
            PlacementDate = placementDate;
        }
        public Request(string[] values)
        {
            Id = int.Parse(values[0]);
            ProductId = int.Parse(values[1]);
            ClientId = int.Parse(values[2]);
            RequestNumber = int.Parse(values[3]);
            RequiredAmount = int.Parse(values[4]);
            PlacementDate = DateTime.FromOADate(double.Parse(values[5]));
        }
        public Request(int id, int requestNumber, int productId, int clientId, int requiredAmount, DateTime placementDate, IEnumerable<Product> products, IEnumerable<Client> clients)
        {
            Id = id;
            RequestNumber = requestNumber;
            ProductId = productId;
            ClientId = clientId;
            RequiredAmount = requiredAmount;
            PlacementDate = placementDate;
            Client = clients.FirstOrDefault(client => Client.Id == client.Id);
            Product = products.FirstOrDefault(product => Product.Id == product.Id);
        }

        public override string ToString()
        {
            return Product.Name + "|" + Client.OrganizationName + "|" + RequestNumber + "|" + RequiredAmount + "|" + PlacementDate.ToString("d") + "|";
        }

        public static string View(IEnumerable<object> list)
        {
            if (list is IEnumerable<Request> requestList)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("                     Заявки");
                sb.AppendLine(Header);
                foreach (var request in requestList)
                {
                    sb.AppendLine(request.ToString());
                }
                return sb.ToString();
            }
            return string.Empty;
        }
    }
}
