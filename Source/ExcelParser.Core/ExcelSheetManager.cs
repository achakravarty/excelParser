﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelParser.Attributes;
using ExcelParser.Model;
using ExcelParser.Model.Extensions;

namespace ExcelParser.Core
{
    public static class ExcelSheetManager
    {
        public static List<Tuple<PropertyInfo, object>> GetValues(Type type, Predicate<Row> predicate, Microsoft.Office.Interop.Excel.Range rows)
        {
            var values = new List<Tuple<PropertyInfo, object>>();
            var properties = type.GetProperties();
            var row = rows.ToRows().FirstOrDefault(r => predicate(r));
            if (row != null)
            {
                properties.ToList().ForEach(property =>
                    {

                        var columnHeader = GetColumnHeader(property);
                        var cells = from cell in row.Cells
                                    where cell.ColumnHeader.Equals(columnHeader)
                                    select cell.Value;
                        var value = cells.First();
                        values.Add(new Tuple<PropertyInfo, object>(property, value));
                    });
            }
            return values;

        }

        public static string GetSheetName(Type type)
        {
            var customAttributes = type.GetCustomAttributes(typeof(ExcelSheetAttribute), false);
            if (customAttributes.Length > 0)
            {
                var excelSheetAttribute = customAttributes[0] as ExcelSheetAttribute;
                if (excelSheetAttribute != null)
                {
                    return excelSheetAttribute.Name;
                }
            }
            return type.Name;
        }

        public static string GetColumnHeader(PropertyInfo propertyInfo)
        {
            var customAttributes = propertyInfo.GetCustomAttributes(typeof(ExcelColumnAttribute), false);
            if (customAttributes.Length > 0)
            {
                var excelColumnAttribute = customAttributes[0] as ExcelColumnAttribute;
                if (excelColumnAttribute != null)
                {
                    return excelColumnAttribute.Name;
                }
            }
            return propertyInfo.Name;
        }
    }
}