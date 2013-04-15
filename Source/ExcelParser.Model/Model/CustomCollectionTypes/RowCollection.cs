using System.Collections;
using System.Collections.Generic;

namespace ExcelParser.Model
{
    public class RowCollection : IEnumerable<Row>
    {
        private readonly List<Row> _rows = new List<Row>();

        public Row this[int rowIndex]
        {
            get { return _rows[rowIndex]; }
        }

        public void Add(Row row)
        {
            if (row == null || row.Cells == null || row.Cells.Count == 0)
            {
                return;
            }
            _rows.Add(row);
        }

        public int IndexOf(Row testRow)
        {
            return _rows.IndexOf(testRow);
        }

        public int CurrentRowIndex(Row row)
        {
            return _rows.IndexOf(row);
        }

        public int Count
        {
            get { return _rows.Count; }
        }

        public IEnumerator GetEnumerator()
        {
            return _rows.GetEnumerator();
        }

        IEnumerator<Row> IEnumerable<Row>.GetEnumerator()
        {
            return _rows.GetEnumerator();
        }
    }
}