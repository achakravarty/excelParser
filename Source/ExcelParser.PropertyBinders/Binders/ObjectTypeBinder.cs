using System;
using System.Linq;
using System.Reflection;
using ExcelParser.Model;
using ExcelParser.PropertyBinders.Interfaces;

namespace ExcelParser.PropertyBinders
{
    public class ObjectTypeBinder : IPropertyBinder
    {
        public void Bind(object target, PropertyInfo propertyInfo, object value)
        {
            var type = propertyInfo.PropertyType;
            var obj = Activator.CreateInstance(type);
            var propertyBinder = new PropertyBinder();
            propertyBinder.Bind(obj, x => x.Cells["Id"].Value.Equals(value.ToString()));
            propertyInfo.SetValue(target, obj, null);
        }
    }
}