using System;
using System.Reflection;
using ExcelParser.PropertyBinders.Interfaces;

namespace ExcelParser.PropertyBinders
{
    public class EnumTypeBinder : IPropertyBinder
    {
        public void Bind(object target, PropertyInfo propertyInfo, object value)
        {
            var enumValue = Enum.Parse(propertyInfo.PropertyType, value.ToString());
            propertyInfo.SetValue(target, enumValue, null);
        }
    }
}