using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading.Tasks;

namespace ExcelImporter
{
    public interface IExcelImporter
    {
        /// <summary>
        /// Reads Excel file from stream into a DataTable.
        /// </summary>
        DataTable ReadToDataTable(Stream stream, ExcelImportOptions? options = null);

        /// <summary>
        /// Reads Excel file from file path into a DataTable.
        /// </summary>
        DataTable ReadToDataTable(string filePath, ExcelImportOptions? options = null);

        /// <summary>
        /// Reads Excel file from stream into a list of objects of type T.
        /// </summary>
        List<T> ReadToList<T>(Stream stream, ExcelImportOptions? options = null) where T : new();

        /// <summary>
        /// Reads Excel file from file path into a list of objects of type T.
        /// </summary>
        List<T> ReadToList<T>(string filePath, ExcelImportOptions? options = null) where T : new();
    }
}
