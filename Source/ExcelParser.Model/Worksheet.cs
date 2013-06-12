using System.Collections.Generic;
namespace ExcelParser.Model
{
    public class Worksheet
    {
        public string Name { get; set; }

        public IEnumerable<Row> Rows { get; set; }
    }
}