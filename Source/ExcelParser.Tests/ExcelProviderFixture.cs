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
        [TestMethod]
        public void ShouldParseExact()
        {
            using (var excelProvider = new ExcelProvider(@"D:\ExcelParser.xlsx"))
            {
                var a = excelProvider.ParseExact<Customer>();
                foreach (var customer in a)
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
    }
}
