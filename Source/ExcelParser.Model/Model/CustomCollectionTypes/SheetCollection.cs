using System.Collections;
using System.Collections.Generic;

namespace ExcelParser.Model
{
    public class SheetCollection : IEnumerable<Worksheet>
    {
        private readonly List<Worksheet> _sheets = new List<Worksheet>();

        public Worksheet this[string sheetName]
        {
            get { return _sheets.Find(x => x.Name == sheetName); }
        }

        public Worksheet this[int sheetIndex]
        {
            get
            {
                if (_sheets != null && _sheets.Count > 0)
                {
                    return _sheets[0];
                }
                return null;
            }
        }

        public void Add(Worksheet sheet)
        {
            if (sheet == null || sheet.Rows == null || sheet.Rows.Count == 0)
            {
                return;
            }
            _sheets.Add(sheet);
        }

        public int Count
        {
            get { return _sheets.Count; }
        }

        public IEnumerator GetEnumerator()
        {
            foreach (var worksheet in _sheets)
            {
                yield return worksheet;
            }
        }

        IEnumerator<Worksheet> IEnumerable<Worksheet>.GetEnumerator()
        {
            foreach (var worksheet in _sheets)
            {
                yield return worksheet;
            }
        }
    }
}