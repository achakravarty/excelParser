using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
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
            TestContext.BeginTimer("Timer1");
            using (var excelProvider = new ExcelProvider(@"D:\ExcelParser.xlsx"))
            {
                var a = excelProvider.ParseExact<Customer>(x => x.Cells["Id"].Value.Equals("1"));
                foreach (var customer in a)
                {
                    TestContext.EndTimer("Timer1");
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
