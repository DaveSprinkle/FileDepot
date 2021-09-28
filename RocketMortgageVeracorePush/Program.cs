using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RocketMortgageVeracorePush
{
    class Program
    {
        static void Main(string[] args)
        {
            Helpers h = new Helpers();
            FileOperations f = new FileOperations();

            h.DisplayTitle();
            h.DisplayMenu();

            List<Order> ords = new List<Order>();
            ords = f.Orders();

            f.CreateTabDelimitedOrderFile(ords);

            f.CreateXML(ords);

            Console.ReadKey();
        }
    }
}
