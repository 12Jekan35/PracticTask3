using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PracticTask3.Models
{
    public class Product : IViewItem
    {
        public int Id { get; private set; }
        public string Name { get; private set; }
        public string MeasureUnit { get; private set; }
        public double PriceForUnit { get; private set; }
        public static string Header { get => " Наименование | Ед. измерения | Цена товара за единицу |"; }
        public Product(int id, string name, string measureUnit, double priceForUnit)
        {
            Id = id;
            Name = name;
            MeasureUnit = measureUnit;
            PriceForUnit = priceForUnit;

        }
        public Product(string[] values)
        {
            Id = int.Parse(values[0]);
            Name = values[1];
            MeasureUnit = values[2];
            PriceForUnit = double.Parse(values[3]);
        }

        public override string ToString()
        {
            return Name + "|" + MeasureUnit + "|" + PriceForUnit + "|";
        }

        public static string View(IEnumerable<object> list)
        {
            if (list is IEnumerable<Product> productList)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("                     Продукты");
                sb.AppendLine(Header);
                foreach (var product in productList)
                {
                    sb.AppendLine(product.ToString());
                }
                return sb.ToString();
            }
            return string.Empty;
        }
    }
}
