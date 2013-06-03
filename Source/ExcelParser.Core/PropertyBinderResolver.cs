using System;
using ExcelParser.PropertyBinders;
using ExcelParser.PropertyBinders.Interfaces;

namespace ExcelParser.Core
{
    public static class PropertyBinderResolver
    {
        public static IPropertyBinder Resolve(Type type)
        {
            if (type.IsValueType || type == typeof(string) || type == typeof(DateTime) || type.IsPrimitive)
            {
                return new ValueTypeBinder();
            }
            return new ObjectTypeBinder();
        }
    }
}