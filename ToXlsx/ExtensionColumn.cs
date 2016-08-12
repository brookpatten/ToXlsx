using System;
using OfficeOpenXml;

namespace ToXlsx
{
    internal class ExtensionColumn<T>
    {
        internal Func<T, object> Map { get; set; }
        internal string Header { get; set; }
        internal Action<ExcelColumn> ConfigureColumn { get; set; }
        internal Action<ExcelRange> ConfigureHeader { get; set; }
        internal Action<ExcelRange, T> ConfigureCell { get; set; }
    }
}