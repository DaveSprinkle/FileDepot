using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Excel = Microsoft.Office.Interop.Excel;

namespace RocketMortgageVeracorePush
{
    public class FileOperations
    {
        static string WorkingDirectory = @"\\hw30106-01\HP_Serv02_Y\RocketMortgage\";
        static string fakeExcelName = "Quicken_ProdOrderFile_2021-09-15.xlsx";

        public void OpenExcelOrder()
        {
            Process.Start(WorkingDirectory + fakeExcelName);
        }

        public List<Order> Orders()
        {
            List<Order> orders = new List<Order>();

            Excel.Application xL = new Excel.Application();
            Excel.Workbook wB = xL.Workbooks.Open(WorkingDirectory + fakeExcelName);
            Excel.Worksheet wS = wB.Sheets[1];
            Excel.Range range = wS.UsedRange;

            for (int r = 3; r <= range.Rows.Count; r++)
            {
                Order newOrder = new Order();
                newOrder.OrderDate = DateTime.FromOADate(range.Cells[r, 1].Value2);
                newOrder.JobNumber = (int)range.Cells[r, 2].Value;
                newOrder.SKU = range.Cells[r, 3].Value;
                newOrder.FileName = range.Cells[r, 4].Value2;
                newOrder.FileURL = range.Cells[r, 5].Value2;
                newOrder.Quantity = (int)range.Cells[r, 6].Value;
                newOrder.OrderShipQuantity = (int)range.Cells[r, 7].Value;
                newOrder.FirstName = range.Cells[r, 8].Value2;
                newOrder.LastName = range.Cells[r, 9].Value2;
                newOrder.Address = range.Cells[r, 10].Value2;
                newOrder.Address2 = range.Cells[r, 11].Value2;
                newOrder.City = range.Cells[r, 12].Value2;
                newOrder.State = range.Cells[r, 13].Value2;
                newOrder.Zip = range.Cells[r, 14].Value2;
                newOrder.Email = range.Cells[r, 15].Value2;
                newOrder.Phone = range.Cells[r, 16].Value2;

                orders.Add(newOrder);
            }

            wB.Close();
            xL.Quit();

            return orders;
        }


        public void CreateDateFolders()
        {
            Excel.Application xL = new Excel.Application();
            Excel.Workbook wB = xL.Workbooks.Open(WorkingDirectory + fakeExcelName);
            Excel.Worksheet wS = wB.Sheets[1];
            Excel.Range range = wS.UsedRange;

            HashSet<string> folders = new HashSet<string>();
            HashSet<string> orders = new HashSet<string>();

            for (int r = 3; r <= range.Rows.Count; r++)
            {
                DateTime orderDate = DateTime.FromOADate(range.Cells[r, 1].Value2);
                string orderNum = range.Cells[r, 1].Value2.ToString();

                string getDateString = orderDate.ToString("yyyyMMdd");

                folders.Add(getDateString);
                
            }

            foreach (string f in folders)
            {
                if(!Directory.Exists(WorkingDirectory + f))
                {
                    Directory.CreateDirectory(WorkingDirectory + f);
                }
            }

            wB.Close();
            xL.Quit();
        }





        //NOT FINISHED
        public void DownloadArt()
        {
            Excel.Application xL = new Excel.Application();
            Excel.Workbook wB = xL.Workbooks.Open(WorkingDirectory + fakeExcelName);
            Excel.Worksheet wS = wB.Sheets[1];
            Excel.Range range = wS.UsedRange;

            for (int r = 3; r <= range.Rows.Count; r++)
            {
                int JobNumber = (int)range.Cells[r, 2].Value;
                string FileURL = range.Cells[r, 5].Value2;
                if(FileURL != null & FileURL != "")
                {
                    using (WebClient webClient = new WebClient())
                    {
                        webClient.DownloadFile(FileURL, WorkingDirectory + JobNumber + ".pdf");
                        Thread.Sleep(1000);
                    }
                }
            }




            wB.Close();
            xL.Quit();
        }



        public void CreateTabDelimitedOrderFile(List<Order> orders)
        {
            StreamWriter sw = new StreamWriter(WorkingDirectory + "test.txt");
            string orderInfo;

            string header = "OrderDate\tJobNumber\tSKU\tFileName\tFileURL\tQuantity\tOrderShipQuantity\tFirstName\tLastName\tAddress\tAddress2\tCity\tState\tZip\tEmail\tPhone";

            sw.WriteLine(header);

            foreach (Order o in orders)
            {
                orderInfo = o.OrderDate + "\t" + o.JobNumber + "\t" + o.SKU + "\t" + o.FileName + "\t" + o.FileURL + "\t" + o.Quantity + "\t" + o.OrderShipQuantity + "\t" + o.FirstName + "\t" + o.LastName + "\t" + o.Address + "\t" + o.Address2 + o.City + "\t" + o.State + "\t" + o.Zip + "\t" + o.Email + "\t" + o.Phone;

                sw.WriteLine(orderInfo);
            }
            sw.Close();
        }


        public void CreateXML(List<Order> orders)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(List<Order>));
            using (TextWriter writer = new StreamWriter(WorkingDirectory + "test.xml"))
            {
                serializer.Serialize(writer, orders);
            }
        }


        public void CreateJSONOrder(Order order, int marker)
        {
            JsonSerializer serializer = new JsonSerializer();
            serializer.NullValueHandling = NullValueHandling.Ignore;

            using (StreamWriter sw = new StreamWriter(WorkingDirectory + marker.ToString() + "json.txt"))

            using (JsonWriter writer = new JsonTextWriter(sw))
            {
                serializer.Serialize(writer, order);
            }
        }


        public void CreateJSONOrders(List<Order> orders)
        {
            JsonSerializer serializer = new JsonSerializer();
            serializer.NullValueHandling = NullValueHandling.Ignore;

            StreamWriter sw = new StreamWriter(WorkingDirectory + "json.txt");
            JsonWriter writer = new JsonTextWriter(sw);

            foreach (Order o in orders)
            {
                serializer.Serialize(writer, o);
            }
            sw.Close();
        }









    }
}
