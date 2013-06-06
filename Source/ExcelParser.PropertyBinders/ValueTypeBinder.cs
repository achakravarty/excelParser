using System.Reflection;
using ExcelParser.PropertyBinders.Interfaces;

namespace ExcelParser.PropertyBinders
{
    public class ValueTypeBinder : IPropertyBinder
    {
        public void Bind(object target, PropertyInfo propertyInfo, object value)
        {
            propertyInfo.SetValue(target, value, null);
        }
    }
}