using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using ExcelDataReader;

namespace ExcelImporter
{
    public class ExcelImporter : IExcelImporter
    {
        static ExcelImporter()
        {
            // Register encoding provider for ExcelDataReader (supports older code pages like 1252)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        public DataTable ReadToDataTable(string filePath, ExcelImportOptions? options = null)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"File not found at path: {filePath}", filePath);
            }

            using var stream = File.OpenRead(filePath);
            return ReadToDataTable(stream, options);
        }

        public DataTable ReadToDataTable(Stream stream, ExcelImportOptions? options = null)
        {
            options ??= new ExcelImportOptions();

            // We use ExcelDataReader which supports .xlsx, .xls, .csv, etc.
            using var reader = ExcelReaderFactory.CreateReader(stream);

            var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = false // We handle header row and data rows manually according to options
                }
            });

            DataTable sourceTable;
            if (!string.IsNullOrEmpty(options.SheetName))
            {
                if (!dataSet.Tables.Contains(options.SheetName))
                {
                    throw new ArgumentException($"Worksheet '{options.SheetName}' was not found in the Excel document.");
                }
                sourceTable = dataSet.Tables[options.SheetName]!;
            }
            else if (options.SheetIndex > 0 && options.SheetIndex <= dataSet.Tables.Count)
            {
                sourceTable = dataSet.Tables[options.SheetIndex - 1];
            }
            else
            {
                sourceTable = dataSet.Tables[0];
            }

            var resultTable = new DataTable();

            int rowCount = sourceTable.Rows.Count;
            int colCount = sourceTable.Columns.Count;

            int headerRowIdx = options.HeaderRowIndex - 1; // Convert 1-based to 0-based
            int dataStartRowIdx = options.HasHeaderRow
                ? (options.DataStartRowIndex > options.HeaderRowIndex ? options.DataStartRowIndex - 1 : options.HeaderRowIndex)
                : options.DataStartRowIndex - 1;

            if (options.HasHeaderRow && headerRowIdx >= 0 && headerRowIdx < rowCount)
            {
                var headerRow = sourceTable.Rows[headerRowIdx];
                for (int c = 0; c < colCount; c++)
                {
                    string colName = headerRow[c]?.ToString()?.Trim() ?? string.Empty;
                    if (string.IsNullOrEmpty(colName))
                    {
                        colName = $"Column{c + 1}";
                    }

                    // Ensure unique column names in DataTable
                    string uniqueColName = colName;
                    int dupCount = 1;
                    while (resultTable.Columns.Contains(uniqueColName))
                    {
                        uniqueColName = $"{colName}_{dupCount++}";
                    }

                    resultTable.Columns.Add(uniqueColName, typeof(object));
                }
            }
            else
            {
                for (int c = 0; c < colCount; c++)
                {
                    resultTable.Columns.Add($"Column{c + 1}", typeof(object));
                }
            }

            for (int r = dataStartRowIdx; r < rowCount; r++)
            {
                var row = sourceTable.Rows[r];

                if (options.SkipEmptyRows && IsRowEmpty(row, colCount))
                {
                    continue;
                }

                DataRow newRow = resultTable.NewRow();
                for (int c = 0; c < colCount; c++)
                {
                    var val = row[c];
                    newRow[c] = val ?? DBNull.Value;
                }
                resultTable.Rows.Add(newRow);
            }

            return resultTable;
        }

        public List<T> ReadToList<T>(string filePath, ExcelImportOptions? options = null) where T : new()
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"File not found at path: {filePath}", filePath);
            }

            using var stream = File.OpenRead(filePath);
            return ReadToList<T>(stream, options);
        }

        public List<T> ReadToList<T>(Stream stream, ExcelImportOptions? options = null) where T : new()
        {
            DataTable dataTable = ReadToDataTable(stream, options);
            var result = new List<T>();

            PropertyInfo[] properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanWrite)
                .ToArray();

            // Map properties to column indexes in dataTable
            var propertyMappings = new List<(PropertyInfo Property, int ColumnIndex)>();

            foreach (var prop in properties)
            {
                var attr = prop.GetCustomAttribute<ExcelColumnAttribute>();
                if (attr != null && attr.Ignore)
                {
                    continue;
                }

                int colIndex = -1;

                if (attr != null && attr.Index > 0)
                {
                    // 1-based index provided
                    int targetIdx = attr.Index - 1;
                    if (targetIdx < dataTable.Columns.Count)
                    {
                        colIndex = targetIdx;
                    }
                }

                if (colIndex == -1 && attr != null && !string.IsNullOrEmpty(attr.Name))
                {
                    if (dataTable.Columns.Contains(attr.Name))
                    {
                        colIndex = dataTable.Columns[attr.Name]!.Ordinal;
                    }
                }

                if (colIndex == -1)
                {
                    // Match by property name
                    if (dataTable.Columns.Contains(prop.Name))
                    {
                        colIndex = dataTable.Columns[prop.Name]!.Ordinal;
                    }
                }

                if (colIndex != -1)
                {
                    propertyMappings.Add((prop, colIndex));
                }
            }

            foreach (DataRow row in dataTable.Rows)
            {
                T item = new T();
                foreach (var (prop, colIndex) in propertyMappings)
                {
                    var val = row[colIndex];
                    if (val != null && val != DBNull.Value)
                    {
                        SetPropertyValue(item, prop, val, options?.Culture);
                    }
                }
                result.Add(item);
            }

            return result;
        }

        private static bool IsRowEmpty(DataRow row, int colCount)
        {
            for (int c = 0; c < colCount; c++)
            {
                var val = row[c];
                if (val != null && val != DBNull.Value && !string.IsNullOrWhiteSpace(val.ToString()))
                {
                    return false;
                }
            }
            return true;
        }

        private static void SetPropertyValue(object target, PropertyInfo prop, object value, IFormatProvider? culture)
        {
            Type targetType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;

            try
            {
                object? convertedValue;

                if (targetType.IsEnum)
                {
                    if (value is string strEnum)
                    {
                        convertedValue = Enum.Parse(targetType, strEnum, ignoreCase: true);
                    }
                    else
                    {
                        convertedValue = Enum.ToObject(targetType, Convert.ChangeType(value, Enum.GetUnderlyingType(targetType)));
                    }
                }
                else if (targetType == typeof(Guid))
                {
                    convertedValue = Guid.Parse(value.ToString()!);
                }
                else if (targetType == typeof(DateTime))
                {
                    if (value is DateTime dt)
                    {
                        convertedValue = dt;
                    }
                    else
                    {
                        convertedValue = Convert.ToDateTime(value, culture ?? System.Globalization.CultureInfo.InvariantCulture);
                    }
                }
                else
                {
                    convertedValue = Convert.ChangeType(value, targetType, culture ?? System.Globalization.CultureInfo.InvariantCulture);
                }

                prop.SetValue(target, convertedValue);
            }
            catch
            {
                // Fallback / ignore conversion failures if type conversion is incompatible
            }
        }
    }
}
