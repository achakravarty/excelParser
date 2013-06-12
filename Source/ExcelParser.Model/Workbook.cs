using System.Collections.Generic;

namespace ExcelParser.Model
{
    public class Workbook
    {
        public IEnumerable<Worksheet> Sheets { get; set; }
    }
}