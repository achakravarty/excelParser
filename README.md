ExcelParser
===========

ExcelParser is a library that parses an excel file using [Microsoft Interop Services](http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.aspx) into a strongly typed object of type specified by the consumer when the parse method is called. The parser deconstructs the type into its constituents and then parses the excel file accordingly to populate the object. It also exposes certain extension methods on top of the [Worksheet] (http://msdn.microsoft.com/en-us/library/office/microsoft.office.interop.excel.worksheet.aspx) type that parses the rows in a sheet into an IEnumerable<> of Row.

