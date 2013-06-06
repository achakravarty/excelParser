using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelParser.Model;
using ExcelParser.PropertyBinders.Interfaces;
using System.Collections;

namespace ExcelParser.PropertyBinders
{
    public class ListTypeBinder : IPropertyBinder
    {
        public void Bind(object target, PropertyInfo propertyInfo, object value)
        {
            var type = propertyInfo.PropertyType;
            var genericType = type.GetGenericArguments()[0];
            var valueString = value.ToString();
            var data = new List<string>();
            if (valueString.Contains(","))
            {
                Array.ForEach(valueString.Split(','), data.Add);
            }
            var listObj = (IList)Activator.CreateInstance(type);
            foreach (var item in data)
            {
                var obj = Activator.CreateInstance(genericType);
                var propertyBinder = new PropertyBinder();
                var tempValue = item;
                propertyBinder.Bind(obj, x => x.Cells["Id"].Value.Equals(tempValue));
                listObj.Add(obj);
            }
            propertyInfo.SetValue(target, listObj, null);
        }
    }
}