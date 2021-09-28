using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RocketMortgageVeracorePush
{
    public class Helpers
    {
        static string[] Title = {"ROCKET MORTGAGE ORDER PROCESSING", "v 1.0" };
        
        static string WorkingDirectory = @"\\hw30106-01\HP_Serv02_Y\RocketMortgage";

        static string[] MenuItems = { "Process Today's Orders", "View Today's Orders", "Exit" };

        public void DisplayTitle()
        {
            int lenTitle = Title[0].Length;
            if (lenTitle % 2 != 0) lenTitle++;
            string barLine = new string('=', lenTitle + 20);
            int littleBump = (lenTitle + 20 - Title[1].Length) / 2;
            Console.WriteLine(barLine);
            Console.WriteLine(new string(' ',10)  + Title[0]);
            Console.WriteLine();
            Console.WriteLine(new string(' ', littleBump) + Title[1]);
            Console.WriteLine(barLine);
        }

        public void DisplayMenu()
        {
            int i = 1;
            Console.WriteLine();
            foreach(string m in MenuItems)
            {
                Console.WriteLine(i + ".) " + m);
                i++;
            }

            Console.WriteLine();
        }










    }
}
