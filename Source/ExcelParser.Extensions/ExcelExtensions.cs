using System;
using System.Linq;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace ExcelParser.Core.Extensions
{
    public static class ExcelExtensions
    {
        public static MsExcel.Worksheet Find(this MsExcel.Sheets worksheets, Predicate<MsExcel.Worksheet> predicate)
        {
            return worksheets.Cast<MsExcel.Worksheet>().FirstOrDefault(worksheet => predicate(worksheet));
        }
    }
}