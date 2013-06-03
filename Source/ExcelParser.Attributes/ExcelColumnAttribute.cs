using System;

namespace ExcelParser.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class ExcelColumnAttribute : Attribute
    {
        public string Name { get; set; }
    }
}