using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelParser.Model;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace ExcelParser
{
    public class ExcelProvider
    {
        /// <summary>
        /// Gets or sets the Excel filename
        /// </summary>
        private string FileName { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="fileName">The Excel file to process</param>
        public ExcelProvider(string fileName)
        {
            FileName = fileName;
        }

        /// <summary>
        /// Gets the Excel Workbook data
        /// </summary>
        /// <returns>The parsed workbook</returns>
        public Workbook GetWorkbook()
        {
            var excelApp = new MsExcel.Application();
            try
            {
                var excelWorkbook = excelApp.Workbooks.Open(FileName);
                var workbook = new Workbook { Worksheets = new SheetCollection() };
                foreach (MsExcel.Worksheet excelWorksheet in excelWorkbook.Worksheets)
                {
                    var worksheet = new Worksheet { Rows = new RowCollection(), Columns = new List<string>() };
                    var usedRange = excelWorksheet.UsedRange;
                    for (int i = 2; i <= usedRange.Rows.Count; i++)
                    {
                        var row = new Row { Cells = new CellCollection() };
                        for (int j = 1; j <= usedRange.Columns.Count; j++)
                        {
                            var cell = new Cell
                                           {
                                               ColumnHeader = usedRange.Cells[1, j].Value != null ? usedRange.Cells[1, j].Value.ToString() : string.Empty,
                                               Value = usedRange.Cells[i, j].Value != null ? usedRange.Cells[i, j].Value.ToString() : string.Empty
                                           };
                            if (string.IsNullOrEmpty(cell.ColumnHeader) == false)
                            {
                                row.Cells.Add(cell);
                            }
                        }
                        worksheet.Rows.Add(row);
                    }
                    worksheet.Name = excelWorksheet.Name;
                    if (worksheet.Rows.Count > 0)
                    {
                        worksheet.Columns.AddRange(worksheet.Rows[0].Cells.ColumnHeaders);
                    }
                    workbook.Worksheets.Add(worksheet);
                }
                return workbook;
            }
            catch (Exception ex)
            {
                return null;
            }
            finally
            {
                excelApp.Quit();
            }
        }
    }
}
