using System;
using ExcelParser.Extensions;
using ExcelParser.Model;

namespace ExcelParser.PropertyBinders
{
    public class PropertyBinder
    {
        public void Bind(object target, Predicate<Row> predicate)
        {
            var type = target.GetType();
            var values = ExcelSheetManager.GetValues(type, predicate);
            values.ForEach(value =>
                {
                    var binder = PropertyBinderResolver.Resolve(value.Item1.PropertyType);
                    binder.Bind(target, value.Item1, value.Item2);
                });
        }
    }
}