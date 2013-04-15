using System.Collections.Generic;

namespace ExcelParser.Model
{
    public class Worksheet
    {
        public string Name { get; set; }

        public RowCollection Rows { get; set; }

        public List<string> Columns { get; set; }
    }
}