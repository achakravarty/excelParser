using System.Reflection;
using ExcelParser.PropertyBinders.Interfaces;

namespace ExcelParser.PropertyBinders
{
    public class ValueTypeBinder : IPropertyBinder
    {
        public void Bind(object obj, PropertyInfo propertyInfo, object value)
        {
            propertyInfo.SetValue(obj, value, null);
        }
    }
}