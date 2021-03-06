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
using Word = Microsoft.Office.Interop.Word;


namespace RocketMortgageVeracorePush
{
    public class FileOperations
    {
        static string WorkingDirectory = @"\\hw30106-01\HP_Serv02_Y\RocketMortgage\";
        static string DownloadDirectory = @"\\hw30106-01\HP_Serv02_Y\RocketMortgage\Download";
        static string LoggingDirectory = @"\\hw30106-01\HP_Serv02_Y\RocketMortgage\";

        static string fakeExcelName = "Quicken_ProdOrderFile_2021-09-15.xlsx";


        public string GetYesterFileName()
        {
            DateTime yesterday = DateTime.Now.AddDays(-1);
            string yesterString = yesterday.ToString("yyyy-MM-dd");
            return "Quicken_ProdOrderFile_" + yesterString + ".xlsx";
        }


        public string[] GetNewFile()
        {
            string[] files = Directory.GetFiles(WorkingDirectory, "*.xlsx");

            return files;
        }

        
        public List<string> GetFiles()
        {
            List<string> files = new List<string>();
            string[] fileNames = Directory.GetFiles(WorkingDirectory, "*.xlsx");

            files = fileNames.ToList();

            return files;
        }
      

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

            Dictionary<int, string> orderNumDate = new Dictionary<int, string>();

            for (int r = 3; r <= range.Rows.Count; r++)
            {
                DateTime orderDate = DateTime.FromOADate(range.Cells[r, 1].Value2);
                int orderNum = (int)range.Cells[r, 2].Value;
                string getDateString = orderDate.ToString("yyyyMMdd");

                orderNumDate.Add(orderNum, getDateString);
            }

            foreach(KeyValuePair<int, string> o in orderNumDate)
            {
                if(!Directory.Exists(WorkingDirectory + o.Value))
                {
                    Directory.CreateDirectory(WorkingDirectory + o.Value);
                    string dateDir = WorkingDirectory + o.Value + @"\";
                    if (!Directory.Exists(dateDir + o.Key))
                    {
                        Directory.CreateDirectory(dateDir + o.Key);
                    }
                }
                else
                {
                    string dateDir = WorkingDirectory + o.Value + @"\";
                    if (!Directory.Exists(dateDir + o.Key))
                    {
                        Directory.CreateDirectory(dateDir + o.Key);
                    }
                }

                Console.WriteLine(o.Key);
            }

            wB.Close();
            xL.Quit();
        }


        public List<string> AllDirectories()
        {
            List<string> dirs = new List<string>();
            string[] directories = Directory.GetDirectories(WorkingDirectory, "*", SearchOption.AllDirectories);
            dirs = directories.ToList<string>();

            return dirs;
        }


        public string GetOrderDirectory(int orderId)
        {
            string dir = "";
            List<string> dirs = AllDirectories();
            foreach(string d in dirs)
            {
                if (Path.GetFileName(Path.GetDirectoryName(d)) == orderId.ToString())
                {
                    dir = Path.GetFullPath(d);
                }
            }
            return dir;
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
                DateTime orderDate = DateTime.FromOADate(range.Cells[r, 1].Value2);
                string getDateString = orderDate.ToString("yyyyMMdd");
                int JobNumber = (int)range.Cells[r, 2].Value;
                string FileURL = range.Cells[r, 5].Value2;
                if(FileURL != null & FileURL != "")
                {
                    using (WebClient webClient = new WebClient())
                    {
                        webClient.DownloadFile(FileURL, WorkingDirectory + getDateString + @"\" + JobNumber + @"\" + JobNumber+ ".pdf");
                        Thread.Sleep(500);
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


        public void CreateOrderSheet(Order order)
        {
            //USE THE WORD APPLICATION 
            Word.Application word = new Word.Application();
            word.ShowAnimation = false;
            word.Visible = false;

            //START A NEW DOCUMENT
            object missing = System.Reflection.Missing.Value;
            Word.Document doc = word.Documents.Add(ref missing, ref missing, ref missing, ref missing);

            //ADD A PARAGRAPH AND A TABLE
            var par1 = doc.Paragraphs.Add();
            var tableRange = doc.Range(1, 2);
            var table1 = doc.Tables.Add(tableRange, 16, 2);

            //PUT PROPERTY NAMES DOWN THE FIRST COLUMN
            table1.Cell(1, 1).Range.Text = "Job Number";
            table1.Cell(2, 1).Range.Text = "Order Date";
            table1.Cell(3, 1).Range.Text = "SKU";
            table1.Cell(4, 1).Range.Text = "File Name";
            table1.Cell(5, 1).Range.Text = "File URL";
            table1.Cell(6, 1).Range.Text = "Quantity";
            table1.Cell(7, 1).Range.Text = "Order Ship Quantity";
            table1.Cell(8, 1).Range.Text = "First Name";
            table1.Cell(9, 1).Range.Text = "Last Name";
            table1.Cell(10, 1).Range.Text = "Address";
            table1.Cell(11, 1).Range.Text = "Address Line 2";
            table1.Cell(12, 1).Range.Text = "City";
            table1.Cell(13, 1).Range.Text = "State";
            table1.Cell(14, 1).Range.Text = "Zip";
            table1.Cell(15, 1).Range.Text = "Email";
            table1.Cell(16, 1).Range.Text = "Phone";

            //ALIGN THE FIRST COLUMN TO THE RIGHT
            for(int i = 1; i < 17; i++)
            {
                table1.Cell(i, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            }

            //PUT THE PROPERTY VALUES DOWN THE SECOND COLUMN 
            table1.Cell(1, 2).Range.Text = order.JobNumber.ToString();
            table1.Cell(2, 2).Range.Text = order.OrderDate.ToString();
            table1.Cell(3, 2).Range.Text = order.SKU.ToString();
            if (order.FileName != null & order.FileName != "") table1.Cell(4, 2).Range.Text = order.FileName.ToString();
            if (order.FileURL != null & order.FileURL != "") table1.Cell(5, 2).Range.Text = order.FileURL.ToString();
            table1.Cell(6, 2).Range.Text = order.Quantity.ToString();
            table1.Cell(7, 2).Range.Text = order.OrderShipQuantity.ToString();
            table1.Cell(8, 2).Range.Text = order.FirstName.ToString();
            table1.Cell(9, 2).Range.Text = order.LastName.ToString();
            table1.Cell(10, 2).Range.Text = order.Address.ToString();
            if(order.Address2 != null & order.Address2 != "") table1.Cell(11, 2).Range.Text = order.Address2.ToString();
            table1.Cell(12, 2).Range.Text = order.City.ToString();
            table1.Cell(13, 2).Range.Text = order.State.ToString();
            table1.Cell(14, 2).Range.Text = order.Zip.ToString();
            table1.Cell(15, 2).Range.Text = order.Email.ToString();
            table1.Cell(16, 2).Range.Text = order.Phone.ToString();

            //CHANGE THE JOB NUMBER ROW TO LARGE FONT
            table1.Cell(1, 1).Range.Font.Size = 24;
            table1.Cell(1, 2).Range.Font.Size = 24;

            //SHRINK THE FIRST COLUMN TO MINIMAL SIZE
            table1.AllowAutoFit = true;
            Word.Column col1 = table1.Columns[1];
            col1.AutoFit();

            //SAVE THE DOCUMENT TO THE PROPER FOLDER AND QUIT WORD
            doc.SaveAs(WorkingDirectory + order.OrderDate + @"\" + order.JobNumber + @"\" + order.JobNumber + ".doc");
            word.Quit();
        }


        public string WriteData(string input)
        {
            if(input != null & input != "")
            {
                return input;
            }
            return "";
        }



    }
}
