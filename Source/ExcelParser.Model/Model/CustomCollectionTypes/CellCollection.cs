using System.Collections;
using System.Collections.Generic;

namespace ExcelParser.Model
{
    public class CellCollection : IEnumerable<Cell>
    {
        private readonly List<Cell> _cells = new List<Cell>();

        /// <summary>
        /// Custom indexer on a collection of cells
        /// </summary>
        /// <param name="columnHeader">The respective column header for the cell to find</param>
        /// <returns>The cell with the specified column header</returns>
        public Cell this[string columnHeader]
        {
            get { return _cells.Find(x => x.ColumnHeader == columnHeader); }
        }

        public Cell this[int index]
        {
            get { return _cells[index]; }
        }

        public void Add(Cell cell)
        {
            if (cell == null || string.IsNullOrWhiteSpace(cell.ColumnHeader))
            {
                return;
            }
            _cells.Add(cell);
        }

        public int Count
        {
            get { return _cells.Count; }
        }

        public List<string> Values
        {
            get
            {
                var cellValues = new List<string>();
                _cells.ForEach(x => cellValues.Add(x.Value));
                return cellValues;
            }
        }

        public List<string> ColumnHeaders
        {
            get
            {
                var columnHeaders = new List<string>();
                _cells.ForEach(x => columnHeaders.Add(x.ColumnHeader));
                return columnHeaders;
            }
        }

        public IEnumerator GetEnumerator()
        {
            return _cells.GetEnumerator();
        }

        IEnumerator<Cell> IEnumerable<Cell>.GetEnumerator()
        {
            return _cells.GetEnumerator();
        }
    }
}