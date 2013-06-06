using System;
using System.Reflection;
using ExcelParser.PropertyBinders.Interfaces;

namespace ExcelParser.PropertyBinders
{
    public class ArrayTypeBinder : IPropertyBinder
    {
        public void Bind(object obj, PropertyInfo propertyInfo, object value)
        {
            throw new NotImplementedException();
        }
    }
}