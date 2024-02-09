using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PracticTask3.Models
{
    public interface IViewItem
    {
        public abstract static string View(IEnumerable<object> list);
    }
}
