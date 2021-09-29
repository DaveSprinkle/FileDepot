using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RocketMortgageVeracorePush
{
    public class UI
    {
        public void RunUI()
        {
            Helpers h = new Helpers();
            FileOperations f = new FileOperations();

            h.DisplayTitle();
            h.DisplayMenu();

            bool exit = false;

            while (exit == false)
            {
                Console.WriteLine("Please make a selection from the menu:");
                char selection = Console.ReadKey(true).KeyChar;

                switch (selection)
                {
                    case '1':
                        f.CreateDateFolders();
                        f.DownloadArt();
                        //GETS THE ORDERS FROM THE EXCEL FILE
                        List<Order> ords = f.Orders();
                        //CREATES A JSON FILE FROM THE ORDERS
                        f.CreateJSONOrders(ords);
                        //CREATES A TAB DELIMITED TEXT FILE FROM THE ORDERS
                        f.CreateTabDelimitedOrderFile(ords);
                        //CREATES AN XML FILE FROM THE ORDERS
                        f.CreateXML(ords);
                        break;
                    case '2':
                        f.OpenExcelOrder();
                        break;
                    case '3':
                        exit = true;
                        break;
                    default:
                        Console.WriteLine("That's not a valid selection. Please try again:");
                        break;
                }
            }
        }
    }
}
