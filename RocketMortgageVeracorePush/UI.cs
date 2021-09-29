using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RocketMortgageVeracorePush
{
    public class UI
    {
            //INSTANTIATE NEEDED CLASSES
            FileOperations f = new FileOperations();
            Helpers h = new Helpers();        
        
        public void RunUI()
        {
            //DISPLAY TITLE AND MENU
            h.DisplayTitle();
            h.DisplayMenu();
            Console.WriteLine();

            //RUN MENU UNTIL USER EXITS
            bool exit = false;
            while (exit == false)
            {
                Console.WriteLine("Please make a selection from the menu:");
                char selection = Console.ReadKey(true).KeyChar;

                switch (selection)
                {
                    case '1':
                        Console.WriteLine("Processing orders...");
                        ProcessFiles();
                        Console.WriteLine("File processing completed.");
                        List<Order> ords = f.Orders();
                        foreach(Order o in ords)
                        {
                            f.CreateOrderSheet(o);
                        }




                        break;
                    case '2':
                        Console.WriteLine("View order sheet...");
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


        public void ProcessFiles()
        {
            f.CreateDateFolders();
            f.DownloadArt();
            List<Order> ords = f.Orders();  //GETS THE ORDERS FROM THE EXCEL FILE
            f.CreateJSONOrders(ords);  //CREATES A JSON FILE FROM THE ORDERS
            f.CreateTabDelimitedOrderFile(ords);  //CREATES A TAB DELIMITED TEXT FILE FROM THE ORDERS
            f.CreateXML(ords);  //CREATES AN XML FILE FROM THE ORDERS
        }



    }
}
