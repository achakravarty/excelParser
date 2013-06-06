using System;
using System.Collections.Generic;
using ExcelParser.PropertyBinders;
using ExcelParser.PropertyBinders.Interfaces;

namespace ExcelParser.PropertyBinders
{
    public static class PropertyBinderResolver
    {
        public static IPropertyBinder Resolve(Type type)
        {
            if (type.IsValueType || type == typeof(string) || type == typeof(DateTime) || type.IsPrimitive)
            {
                return new ValueTypeBinder();
            }
            if (type.IsEnum)
            {
                return new EnumTypeBinder();
            }
            if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(List<>))
            {
                return new ListTypeBinder();
            }
            return new ObjectTypeBinder();
        }
    }
}