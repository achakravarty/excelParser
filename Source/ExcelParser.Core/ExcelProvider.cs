using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.OleDb;
using System.Linq;
using System.Reflection;
using System.Text;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace ExcelParser.Core
{
    public class ExcelProvider : IDisposable
    {
        private readonly string _fileName;

        private readonly MsExcel.Application _excelApp;

        public ExcelProvider(string fileName)
        {
            _fileName = fileName;
            _excelApp = new MsExcel.Application();
        }

        public MsExcel.Workbook Parse()
        {
            var workbook = _excelApp.Workbooks.Open(_fileName);
            return workbook;
        }

        public bool TryParse(out MsExcel.Workbook workbook, out string errorMessage)
        {
            workbook = null;
            errorMessage = string.Empty;
            try
            {
                workbook = Parse();
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
            return true;
        }

        public IEnumerable<T> ParseExact<T>(Predicate<Row> predicate) where T : class
        {
            var workbook = Parse();
            var parsed = ParseExact<T>(workbook, predicate);
            return parsed;
        }

        public IEnumerable<T> ParseExact<T>(MsExcel.Workbook workbook, Predicate<Row> predicate) where T : class
        {
            if (workbook != null)
            {
                var type = typeof(T);
                var obj = Activator.CreateInstance(type);
                var sheetName = ExcelDataSource.GetSheetName(type);
                var sheet = workbook.Worksheets.Find(x => x.Name.Equals(sheetName));
                var values = ExcelDataSource.GetValues(type, predicate, sheet.UsedRange);
                values.ForEach(value =>
                {
                    var binder = PropertyBinderResolver.Resolve(value.Item1.PropertyType);
                    binder.Bind(obj, value.Item1, value.Item2);
                });
                yield return obj as T;
            }
        }

        public bool TryParseExact<T>(Predicate<Row> predicate, out IEnumerable<T> objects, out Exception exception) where T : class
        {
            objects = null;
            exception = null;
            try
            {
                objects = ParseExact<T>(predicate);
            }
            catch (Exception ex)
            {
                exception = ex;
                return false;
            }
            return true;
        }

        public bool TryParseExact<T>(MsExcel.Workbook workbook, Predicate<Row> predicate, out IEnumerable<T> objects, out Exception exception) where T : class
        {
            objects = null;
            exception = null;
            try
            {
                objects = ParseExact<T>(workbook, predicate);
            }
            catch (Exception ex)
            {
                exception = ex;
                return false;
            }
            return true;
        }

        public void Dispose()
        {
            _excelApp.Quit();
        }
    }

    public class Row
    {
        public int Index { get; set; }

        public IEnumerable<Cell> Cells { get; set; }
    }

    public class Cell
    {
        public string ColumnHeader { get; set; }

        public string Value { get; set; }
    }

    public interface IPropertyBinder
    {
        void Bind(object obj, PropertyInfo propertyInfo, object value);
    }

    public class ValueTypeBinder : IPropertyBinder
    {
        public void Bind(object obj, PropertyInfo propertyInfo, object value)
        {
            propertyInfo.SetValue(obj, value, null);
        }
    }

    public class EnumTypeBinder : IPropertyBinder
    {
        public void Bind(object obj, PropertyInfo propertyInfo, object value)
        {
            throw new NotImplementedException();
        }
    }

    public class ObjectTypeBinder : IPropertyBinder
    {
        public void Bind(object obj, PropertyInfo propertyInfo, object value)
        {
            throw new NotImplementedException();
        }
    }

    public class ListTypeBinder : IPropertyBinder
    {
        public void Bind(object obj, PropertyInfo propertyInfo, object value)
        {
            throw new NotImplementedException();
        }
    }

    public static class ExcelDataSource
    {
        public static List<Tuple<PropertyInfo, object>> GetValues(Type type, Predicate<Row> predicate, MsExcel.Range rows)
        {
            var values = new List<Tuple<PropertyInfo, object>>();
            var properties = type.GetProperties();
            var row = rows.ToRows().FirstOrDefault(r => predicate(r));
            properties.ToList().ForEach(property =>
                        {
                            var columnHeader = GetColumnHeader(property);
                            var cells = from cell in row.Cells
                                        where cell.ColumnHeader.Equals(columnHeader)
                                        select cell.Value;
                            var value = cells.First();
                            values.Add(new Tuple<PropertyInfo, object>(property, value));
                        });
            return values;

        }

        public static string GetSheetName(Type type)
        {
            var customAttributes = type.GetCustomAttributes(typeof(ExcelSheetAttribute), false);
            if (customAttributes.Length > 0)
            {
                return (customAttributes[0] as ExcelSheetAttribute).Name;
            }
            return type.Name;
        }

        public static string GetColumnHeader(PropertyInfo propertyInfo)
        {
            var customAttributes = propertyInfo.GetCustomAttributes(typeof(ExcelColumnAttribute), false);
            if (customAttributes.Length > 0)
            {
                return (customAttributes[0] as ExcelColumnAttribute).Name;
            }
            return propertyInfo.Name;
        }
    }

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

    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    public class ExcelSheetAttribute : Attribute
    {
        public string Name { get; set; }
    }

    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class ExcelColumnAttribute : Attribute
    {
        public string Name { get; set; }
    }

    public static class ExcelExtensions
    {
        public static MsExcel.Worksheet Find(this MsExcel.Sheets worksheets, Predicate<MsExcel.Worksheet> predicate)
        {
            return worksheets.Cast<MsExcel.Worksheet>().FirstOrDefault(worksheet => predicate(worksheet));
        }

        internal static IEnumerable<Row> ToRows(this MsExcel.Range usedRange)
        {
            var columnsCount = usedRange.Columns.Count;
            var rowsCount = usedRange.Rows.Count;
            for (var i = 2; i <= rowsCount; i++)
            {
                yield return new Row {Index = i - 2, Cells = GetCells(usedRange, i, columnsCount)};
            }
        }

        internal static IEnumerable<Cell> GetCells(MsExcel.Range usedRange, int rowIndex, int columnsCount)
        {
            for (var i = 1; i <= columnsCount; i++)
            {
                yield return
                    new Cell
                        {
                            ColumnHeader = usedRange[1, i].Value.ToString(),
                            Value = usedRange.Cells[rowIndex, i].Value.ToString()
                        };
            }
        }
    }
}
