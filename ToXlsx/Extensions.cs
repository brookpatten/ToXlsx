using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Text.RegularExpressions;

using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Style;

namespace ToXlsx
{
    public static class Extensions
    {
        /// <summary>
        /// creates an epplus worksheet from a list
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rows"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static ExtensionWorksheet<T> ToWorksheet<T>(this IList<T> rows, string name, Action<ExcelColumn> configureColumn = null, Action<ExcelRange> configureHeader = null, Action<ExcelRange> configureHeaderRow = null, Action<ExcelRange, T> configureCell = null)
        {
            var worksheet = new ExtensionWorksheet<T>()
            {
                Name = name,
                Workbook = new ExtensionWorkbook(),
                Rows = rows,
                Columns = new List<ExtensionColumn<T>>(),
                ConfigureHeader = configureHeader,
                ConfigureColumn = configureColumn,
                ConfigureHeaderRow = configureHeaderRow,
                ConfigureCell = configureCell
            };
            return worksheet;
        }
        /// <summary>
        /// starts a new worksheet
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="K"></typeparam>
        /// <param name="previousSheet"></param>
        /// <param name="rows"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static ExtensionWorksheet<T> NextWorksheet<T, K>(this ExtensionWorksheet<K> previousSheet, IList<T> rows, string name, Action<ExcelColumn> configureColumn = null, Action<ExcelRange> configureHeader = null, Action<ExcelRange> configureHeaderRow = null, Action<ExcelRange, T> configureCell = null)
        {
            previousSheet.AppendWorksheet();
            var worksheet = new ExtensionWorksheet<T>()
            {
                Name = name,
                Workbook = previousSheet.Workbook,
                Rows = rows,
                Columns = new List<ExtensionColumn<T>>(),
                ConfigureHeader = configureHeader ?? previousSheet.ConfigureHeader,
                ConfigureColumn = configureColumn ?? previousSheet.ConfigureColumn,
                ConfigureHeaderRow = configureHeaderRow ?? previousSheet.ConfigureHeaderRow,
                ConfigureCell = configureCell
            };
            return worksheet;
        }

        /// <summary>
        /// adds a column mapping.  If no column mappings are specified all public properties will be used
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="map"></param>
        /// <param name="columnHeader"></param>
        /// <param name="configureColumn"></param>
        /// <param name="configureHeader"></param>
        /// <returns></returns>
        public static ExtensionWorksheet<T> WithColumn<T>(this ExtensionWorksheet<T> worksheet, Func<T, object> map,
            string columnHeader, Action<ExcelColumn> configureColumn = null, Action<ExcelRange> configureHeader = null, Action<ExcelRange, T> configureCell = null)
        {
            worksheet.Columns.Add(new ExtensionColumn<T>()
            {
                Map = map,
                ConfigureHeader = configureHeader,
                ConfigureColumn = configureColumn,
                Header = columnHeader,
                ConfigureCell = configureCell
            });
            return worksheet;
        }

        /// <summary>
        /// adds a title row to the top of the sheet
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="title"></param>
        /// <param name="configureTitle"></param>
        /// <returns></returns>
        public static ExtensionWorksheet<T> WithTitle<T>(this ExtensionWorksheet<T> worksheet, string title, Action<ExcelRange> configureTitle = null)
        {
            if (worksheet.Titles == null)
            {
                worksheet.Titles = new List<ExtensionTitleRow>();
            }

            worksheet.Titles.Add(new ExtensionTitleRow()
            {
                Title = title,
                ConfigureTitle = configureTitle
            });

            return worksheet;
        }
        public static ExcelPackage ToPackage<T>(this IList<T> rows)
        {
            return rows.ToWorksheet(typeof(T).Name).ToPackage();
        }
        public static ExcelPackage ToPackage<T>(this ExtensionWorksheet<T> lastWorksheet)
        {
            lastWorksheet.AppendWorksheet();
            return lastWorksheet.Workbook.Package;
        }
        public static byte[] ToXlsx<T>(this IList<T> rows)
        {
            return rows.ToWorksheet(typeof(T).Name).ToXlsx();
        }
        public static byte[] ToXlsx<T>(this ExtensionWorksheet<T> lastWorksheet)
        {
            lastWorksheet.AppendWorksheet();
            var package = lastWorksheet.Workbook.Package;

            using (var stream = new MemoryStream())
            {
                package.SaveAs(stream);
                package.Dispose();
                return stream.ToArray();
            }
        }
        public static string ToSentenceCase(this string str)
        {
            return Regex.Replace(str, "[a-z][A-Z]", m => $"{m.Value[0]} {m.Value[1]}");
        }
    }
}
