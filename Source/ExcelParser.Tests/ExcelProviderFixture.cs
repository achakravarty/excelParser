using System.Collections.Generic;
using ExcelParser.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelParser.Tests
{
    [TestClass]
    public class ExcelProviderFixture
    {
        public TestContext TestContext { get; set; }

        [TestMethod]
        public void ShouldParseExact()
        {
            using (var excelProvider = new ExcelProvider(@"D:\ExcelParser.xlsx"))
            {
                var customers = excelProvider.ParseExact<Customer>(x => x.Cells["Id"].Value.Equals("1"));
                foreach (var customer in customers)
                {
                    //Do Something
                }
            }
        }
    }

    public class Customer
    {
        public string Id { get; set; }

        public string Name { get; set; }

        public string Phone { get; set; }

        public string City { get; set; }

        public string Country { get; set; }

        public List<Order> Orders { get; set; }

        public Address Address { get; set; }
    }

    public class Order
    {
        public string Name { get; set; }
    }

    public class Address
    {
        public string Name { get; set; }
    }
}
