using System;

namespace ExcelImporter
{
    public class ExcelImportOptions
    {
        /// <summary>
        /// Gets or sets whether the first row contains column headers. Default is true.
        /// </summary>
        public bool HasHeaderRow { get; set; } = true;

        /// <summary>
        /// 1-based header row index. Default is 1.
        /// </summary>

        public int HeaderRowIndex { get; set; } = 1;

        /// <summary>
        /// 1-based start row index for reading data. Default is 2 (or 1 if HasHeaderRow is false).
        /// </summary>
        public int DataStartRowIndex { get; set; } = 2;

        /// <summary>
        /// Name of the worksheet to import. If null or empty, reads the first worksheet.
        /// </summary>
        public string? SheetName { get; set; }

        /// <summary>
        /// 1-based index of the worksheet to import. Default is 1.
        /// </summary>
        public int SheetIndex { get; set; } = 1;

        /// <summary>
        /// Gets or sets whether empty rows should be skipped. Default is true.
        /// </summary>
        public bool SkipEmptyRows { get; set; } = true;

        /// <summary>
        /// Custom string format or culture for parsing dates/numbers if needed.
        /// </summary>
        public IFormatProvider? Culture { get; set; }
    }
}
