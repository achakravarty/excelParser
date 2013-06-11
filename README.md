ExcelParser
===========

ExcelParser is a library that parses an excel file using [Microsoft Interop Services](http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.aspx) into a strongly typed object of type specified by the consumer when the parse method is called. The parser deconstructs the type into its constituent properties and then parses the excel file accordingly to populate the object.

## Features ##

- Supports complex types
- Supports multiple excel sheets within a workbook
- Provides attributes to tag the sheet name with the type

## How to use ##

Add a reference to your project for the ExcelParser.dll

Use the ExcelProvider class defined in this class to open a connection to the excel file that you wish to parse.

```csharp
using(var excelProvider = new ExcelProvider(fileName))
{
}
```

Since the class implements IDisposable interface, the connection is automatically closed upon exiting the using block.

Then you can invoke the [ParseExact&lt;T&gt;](../../Core#excelprovider) method with the type into which the excel needs to be parsed into.

```csharp
using(var excelProvider = new ExcelProvider(fileName))
{
	var customers = excelProvider.ParseExact<Customer>();
}
```
The above method will return an IEnumerable of Customer object by parsing the sheet name Customer. If you wish to add a constraint on the rows that are to be parsed, you can do so by providing a Predicate to the ParseExact&lt;T&gt; method, like

```csharp
using(var excelProvider = new ExcelProvider(fileName))
{
	var customers = excelProvider.ParseExact<Customer>(x=>x.Cells["Id"].Value.Equals("1"));
}
```

The Predicate filters on a [Row](ExcelParser/wiki/Model#row) type that is defined with the ExcelParser library and exposes a property [Cells](Wiki/Model#cellindexer) which can index the Cells in a row based on the ColumnHeader. So the above will return all customers whose Column with Header Id has value 1.

Once you have the IEnumerable, you can easily iterate over them and utilize them any which way you want.

```csharp
using(var excelProvider = new ExcelProvider(fileName))
{
	var customers = excelProvider.ParseExact<Customer>(x=>x.Cells["Id"].Value.Equals("1"));
	foreach(var customer in customers)
	{
		//Do Something
	}
}
```

## Some More Stuff ##

The ExcelParser library also exposes certain extension methods on the Worksheet type of Microsoft Interop Services which type casts the rows within a sheet to a custom Row type.

You can invoke this extension method by

```csharp
var rows = Worksheet.UsedRange.ToRows();
```

The reason I have used Worksheet.UsedRange over here is that the UsedRange property gives me the rows in the sheet which are in used i.e have data in them.
