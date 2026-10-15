using System;

namespace ExcelImporter
{
    /// <summary>
    /// Attribute to map property to an Excel column name or 1-based column index.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelColumnAttribute : Attribute
    {
        public string? Name { get; set; }
        public int Index { get; set; } = -1;
        public bool Ignore { get; set; }

        public ExcelColumnAttribute()
        {
        }

        public ExcelColumnAttribute(string name)
        {
            Name = name;
        }

        public ExcelColumnAttribute(int index)
        {
            Index = index;
        }
    }
}
