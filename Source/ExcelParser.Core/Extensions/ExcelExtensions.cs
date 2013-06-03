using System;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace ExcelParser.Core.Extensions
{
    public static class ExcelExtensions
    {
        public static Worksheet Find(this Sheets worksheets, Predicate<Worksheet> predicate)
        {
            return worksheets.Cast<Worksheet>().FirstOrDefault(worksheet => predicate(worksheet));
        }
    }
}