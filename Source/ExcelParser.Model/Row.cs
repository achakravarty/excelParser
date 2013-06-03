using System.Collections.Generic;

namespace ExcelParser.Model
{
    public class Row
    {
        public int Index { get; set; }

        public IEnumerable<Cell> Cells { get; set; }
    }
}