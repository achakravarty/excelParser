using System.Reflection;

namespace ExcelParser.PropertyBinders.Interfaces
{
    public interface IPropertyBinder
    {
        void Bind(object obj, PropertyInfo propertyInfo, object value);
    }
}