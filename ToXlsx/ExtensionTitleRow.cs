using System;
using OfficeOpenXml;

namespace ToXlsx
{
    internal class ExtensionTitleRow
    {
        internal string Title { get; set; }
        internal Action<ExcelRange> ConfigureTitle { get; set; }
    }
}