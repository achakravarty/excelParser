using System;

namespace ExcelParser.Attributes
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    public class ExcelSheetAttribute : Attribute
    {
        public string Name { get; set; }
    }
}