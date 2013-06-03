using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelParser.Model
{
    public class CellIndexer : CustomGetIndexer<string, Cell>
    {
        private readonly IEnumerable<Cell> _cells;

        public CellIndexer(IEnumerable<Cell> cells)
        {
            _cells = cells;
        }

        protected override Func<string, Cell> Getter
        {
            get
            {
                return x => _cells.FirstOrDefault(y => y.ColumnHeader.Equals(x));
            }
        }

        public IEnumerable<Cell> Where(Predicate<Cell> predicate)
        {
            return _cells.Where(x => predicate(x));
        }
    }
}