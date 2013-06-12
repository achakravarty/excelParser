using System.Collections.Generic;
using System.Linq;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace ExcelParser.Model.Extensions
{
    public static class ModelExtensions
    {
        public static IEnumerable<Worksheet> ToModel(this MsExcel.Sheets sheets)
        {
            return from MsExcel.Worksheet sheet in sheets select new Worksheet { Name = sheet.Name, Rows = sheet.UsedRange.ToModel() };
        }

        public static IEnumerable<Row> ToModel(this MsExcel.Range usedRange)
        {
            var columnsCount = usedRange.Columns.Count;
            var rowsCount = usedRange.Rows.Count;
            for (var i = 2; i <= rowsCount; i++)
            {
                yield return new Row { Index = i - 2, Cells = new CellIndexer(GetCells(usedRange, i, columnsCount)) };
            }
        }

        internal static IEnumerable<Cell> GetCells(MsExcel.Range usedRange, int rowIndex, int columnsCount)
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