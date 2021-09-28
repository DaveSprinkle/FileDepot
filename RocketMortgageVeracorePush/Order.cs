using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace RocketMortgageVeracorePush
{
    [Serializable]
    public class Order
    {
        public int OrderID { get; set; }
        public DateTime OrderDate { get; set; }
        public int JobNumber { get; set; }
        public string SKU { get; set; }
        public string FileName { get; set; }
        public string FileURL { get; set; }
        public int Quantity { get; set; }
        public int OrderShipQuantity { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Address { get; set; }
        public string Address2 { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Zip { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        
        
        public Order() { }


        public Order(int orderID, DateTime orderDate, int jobNumber, string sku, string fileName, string fileURL, int quantity, int orderShipQuantity, string firstName, string lastName, string address, string address2, string city, string state, string zip, string email, string phone)
        {
            OrderID = orderID;
            OrderDate = orderDate;
            JobNumber = jobNumber;
            SKU = sku;
            FileName = fileName;
            FileURL = fileURL;
            Quantity = quantity;
            OrderShipQuantity = orderShipQuantity;
            FirstName = firstName;
            LastName = lastName;
            Address = address;
            Address2 = address2;
            City = city;
            State = state;
            Zip = zip;
            Email = email;
            Phone = phone;
        }


        public void ShowOrderDetails(Order order)
        {
            Console.WriteLine(order.OrderID);
            Console.WriteLine(order.OrderDate);
            Console.WriteLine(order.JobNumber);
            Console.WriteLine(order.SKU);
            Console.WriteLine(order.FileName);
            Console.WriteLine(order.FileURL);
            Console.WriteLine(order.Quantity);
            Console.WriteLine(order.OrderShipQuantity);
            Console.WriteLine(order.FirstName);
            Console.WriteLine(order.LastName);
            Console.WriteLine(order.Address);
            Console.WriteLine(order.Address2);
            Console.WriteLine(order.City);
            Console.WriteLine(order.State);
            Console.WriteLine(order.Zip);
            Console.WriteLine(order.Email);
            Console.WriteLine(order.Phone);
        }


        public void ShowThisOrderDetails()
        {
            Console.WriteLine();
            Console.WriteLine("CUSTOMER INFORMATION:");
            Console.WriteLine();
            Console.WriteLine("\t" + this.OrderDate.ToString("MM/dd/yyyy"));
            Console.WriteLine("\t" + this.JobNumber);
            Console.WriteLine("\t" + this.SKU);
            Console.WriteLine("\t" + this.FileName);
            Console.WriteLine("\t" + this.FileURL);
            Console.WriteLine("\t" + this.Quantity);
            Console.WriteLine("\t" + this.OrderShipQuantity);
            Console.WriteLine("\t" + this.FirstName);
            Console.WriteLine("\t" + this.LastName);
            Console.WriteLine("\t" + this.Address);
            Console.WriteLine("\t" + this.Address2);
            Console.WriteLine("\t" + this.City);
            Console.WriteLine("\t" + this.State);
            Console.WriteLine("\t" + this.Zip);
            Console.WriteLine("\t" + this.Email);
            Console.WriteLine("\t" + this.Phone);
        }
        






    }
}
