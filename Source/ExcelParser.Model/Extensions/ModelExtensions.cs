using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace ExcelParser.Model.Extensions
{
    public static class ModelExtensions
    {
        public static IEnumerable<Row> ToRows(this Range usedRange)
        {
            var columnsCount = usedRange.Columns.Count;
            var rowsCount = usedRange.Rows.Count;
            for (var i = 2; i <= rowsCount; i++)
            {
                yield return new Row { Index = i - 2, Cells = new CellIndexer(GetCells(usedRange, i, columnsCount)) };
            }
        }

        public static IEnumerable<Cell> GetCells(Range usedRange, int rowIndex, int columnsCount)
        {
            for (var i = 1; i <= columnsCount; i++)
            {
                yield return
                    new Cell
                    {
                        ColumnHeader = usedRange[1, i].Value.ToString(),
                        Value = usedRange.Cells[rowIndex, i].Value.ToString()
                    };
            }
        }
    }
}