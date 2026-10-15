# ExcelImporter

A flexible, high-performance .NET library for importing Excel (`.xlsx`, `.xls`) and CSV files into DataTables or strongly-typed model lists.

## Features

- Import Excel files (`.xlsx`, `.xls`) and `.csv` files into `DataTable` or `List<T>`.
- Attribute-based column mapping (`[ExcelColumn]`) by name or 1-based index.
- Ignore specific properties during import using `[ExcelColumn(Ignore = true)]`.
- Flexible import options (`ExcelImportOptions`) including custom header row index, data start row, sheet name / index selection, culture settings, and option to skip empty rows.
- Cross-platform .NET library built on top of `ExcelDataReader` and `ClosedXML`.

## Installation

Add the `ExcelImporter` package or reference the project in your .NET solution.

```xml
<ItemGroup>
  <ProjectReference Include="path/to/src/ExcelImporter/ExcelImporter.csproj" />
</ItemGroup>
```

## Quick Start

### 1. Read Excel into DataTable

```csharp
using ExcelImporter;

var importer = new ExcelImporter();
DataTable table = importer.ReadToDataTable("sample.xlsx");
```

### 2. Read Excel into Strongly-Typed Models

Define your model class:

```csharp
using ExcelImporter;

public class Employee
{
    [ExcelColumn("Full Name")]
    public string Name { get; set; } = string.Empty;

    [ExcelColumn("Department")]
    public string Department { get; set; } = string.Empty;

    [ExcelColumn(Index = 3)] // 1-based index
    public decimal Salary { get; set; }

    [ExcelColumn(Ignore = true)]
    public string InternalNotes { get; set; } = string.Empty;
}
```

Import into a list:

```csharp
var importer = new ExcelImporter();
List<Employee> employees = importer.ReadToList<Employee>("employees.xlsx");
```

### 3. Custom Import Options

```csharp
var options = new ExcelImportOptions
{
    SheetName = "Q4 Data",
    HasHeaderRow = true,
    HeaderRowIndex = 1,
    DataStartRowIndex = 2,
    SkipEmptyRows = true
};

List<Employee> employees = importer.ReadToList<Employee>("quarterly.xlsx", options);
```

## License

MIT
