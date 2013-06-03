using System;
using System.Collections.Generic;
using ExcelParser.Core.Extensions;
using ExcelParser.Model;
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

        #region Parse Methods

        public MsExcel.Workbook Parse()
        {
            var workbook = _excelApp.Workbooks.Open(_fileName);
            return workbook;
        }

        #endregion

        #region ParseExact<T> Methods

        public IEnumerable<T> ParseExact<T>() where T : class
        {
            return ParseExact<T>(x => true);
        }

        public IEnumerable<T> ParseExact<T>(MsExcel.Workbook workbook) where T : class
        {
            return ParseExact<T>(workbook, x => true);
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
                var sheetName = ExcelSheetManager.GetSheetName(type);
                var sheet = workbook.Worksheets.Find(x => x.Name.Equals(sheetName));
                var values = ExcelSheetManager.GetValues(type, predicate, sheet.UsedRange);
                values.ForEach(value =>
                {
                    var binder = PropertyBinderResolver.Resolve(value.Item1.PropertyType);
                    binder.Bind(obj, value.Item1, value.Item2);
                });
                yield return obj as T;
            }
        }

        #endregion

        #region TryParse Methods

        public bool TryParse(out MsExcel.Workbook workbook, out Exception exception)
        {
            workbook = null;
            exception = null;
            try
            {
                workbook = Parse();
            }
            catch (Exception ex)
            {
                exception = ex;
                return false;
            }
            return true;
        }

        #endregion

        #region TryParseExact<T> Methods

        public bool TryParseExact<T>(out IEnumerable<T> objects, out Exception exception) where T : class
        {
            objects = null;
            try
            {
                MsExcel.Workbook workbook;
                var isWorkbookParsed = TryParse(out workbook, out exception);
                if (isWorkbookParsed)
                {
                    objects = ParseExact<T>(x => true);
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                exception = ex;
                return false;
            }
            return true;
        }

        public bool TryParseExact<T>(MsExcel.Workbook workbook, out IEnumerable<T> objects, out Exception exception) where T : class
        {
            objects = null;
            exception = null;
            try
            {
                objects = ParseExact<T>(x => true);
            }
            catch (Exception ex)
            {
                exception = ex;
                return false;
            }
            return true;
        }

        public bool TryParseExact<T>(Predicate<Row> predicate, out IEnumerable<T> objects, out Exception exception) where T : class
        {
            objects = null;
            try
            {
                MsExcel.Workbook workbook;
                var isWorkbookParsed = TryParse(out workbook, out exception);
                if (isWorkbookParsed)
                {
                    objects = ParseExact<T>(predicate);
                }
                else
                {
                    return false;
                }
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

        #endregion

        #region IDisposable Methods

        public void Dispose()
        {
            _excelApp.Quit();
        }

        #endregion
    }
}
