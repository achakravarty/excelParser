using System;
using System.Reflection;

namespace ExcelParser.PropertyBinders.Interfaces
{
    public interface IPropertyBinder
    {
        void Bind(object target, PropertyInfo propertyInfo, object value);
    }
}