using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using OfficeOpenXml;

namespace ToXlsx
{
    public class ExtensionWorksheet<T>
    {
        internal string Name { get; set; }
        internal ExtensionWorkbook Workbook { get; set; }
        internal IList<T> Rows { get; set; }
        internal IList<ExtensionColumn<T>> Columns { get; set; }
        internal IList<ExtensionTitleRow> Titles { get; set; }
        internal Action<ExcelColumn> ConfigureColumn { get; set; }
        internal Action<ExcelRange> ConfigureHeader { get; set; }
        internal Action<ExcelRange> ConfigureHeaderRow { get; set; }
        internal Action<ExcelRange, T> ConfigureCell { get; set; }

        /// <summary>
        /// generates columns for all public properties on the type
        /// </summary>
        /// <returns></returns>
        internal IList<ExtensionColumn<T>> AutoGenerateColumns()
        {
            var columns = new List<ExtensionColumn<T>>();

            var type = typeof(T);
            var properties = type.GetProperties();

            foreach (var property in properties)
            {
                var column = new ExtensionColumn<T>();
                column.Header = property.Name.ToSentenceCase();
                column.Map = GetGetter<T>(property.Name);
                column.ConfigureColumn = c => c.AutoFit();
                columns.Add(column);
            }

            return columns;
        }

        /// <summary>
        /// Generates a Func from a propertyName</T>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        private Func<T, object> GetGetter<T>(string propertyName)
        {
            ParameterExpression arg = Expression.Parameter(typeof(T), "x");
            MemberExpression expression = Expression.Property(arg, propertyName);
            UnaryExpression conversion = Expression.Convert(expression, typeof(object));
            return Expression.Lambda<Func<T, object>>(conversion, arg).Compile();
        }

        /// <summary>
        /// wraps creation of an epplus worksheet
        /// </summary>
        internal void AppendWorksheet()
        {
            if (Workbook.Package == null)
            {
                Workbook.Package = new ExcelPackage();
            }

            var worksheet = Workbook.Package.Workbook.Worksheets.Add(this.Name);

            int rowOffset = 0;

            //if no columns specified auto generate them with reflection
            if (Columns == null || !Columns.Any())
            {
                Columns = AutoGenerateColumns();
            }

            //render title rows
            if (Titles != null)
            {
                for (var i = 0; i < Titles.Count; i++)
                {
                    var range = worksheet.Cells[rowOffset + 1, 1, rowOffset + 1, Columns.Count];
                    range.Merge = true;
                    range.Value = Titles[i].Title;
                    if (Titles[i].ConfigureTitle != null)
                    {
                        Titles[i].ConfigureTitle(range);
                    }
                }
                rowOffset = rowOffset + Titles.Count;
            }

            //render headers
            for (int i = 0; i < Columns.Count; i++)
            {
                worksheet.Cells[rowOffset + 1, i + 1].Value = Columns[i].Header;
                if (ConfigureHeader != null)
                {
                    ConfigureHeader(worksheet.Cells[rowOffset + 1, i + 1]);
                }
                if (Columns[i].ConfigureHeader != null)
                {
                    Columns[i].ConfigureHeader(worksheet.Cells[rowOffset + 1, i + 1]);
                }
            }

            //configure the header row
            if (ConfigureHeaderRow != null)
            {
                ConfigureHeaderRow(worksheet.Cells[rowOffset + 1, 1, rowOffset + 1, Columns.Count]);
            }
            else
            {
                worksheet.Cells[rowOffset + 1, 1, rowOffset + 1, Columns.Count].AutoFilter = true;
            }

            rowOffset++;

            //render data
            if (Rows != null)
            {
                for (var r = 0; r < Rows.Count(); r++)
                {
                    for (var c = 0; c < Columns.Count(); c++)
                    {
                        worksheet.Cells[r + rowOffset + 1, c + 1].Value = Columns[c].Map(Rows[r]);

                        if (this.ConfigureCell != null)
                        {
                            this.ConfigureCell(worksheet.Cells[r + rowOffset + 1, c + 1], Rows[r]);
                        }
                        if (Columns[c].ConfigureCell != null)
                        {
                            Columns[c].ConfigureCell(worksheet.Cells[r + rowOffset + 1, c + 1], Rows[r]);
                        }
                    }
                }
            }

            //configure columns
            for (int i = 0; i < Columns.Count; i++)
            {
                if (ConfigureColumn != null)
                {
                    ConfigureColumn(worksheet.Column(i + 1));
                }
                if (Columns[i].ConfigureColumn != null)
                {
                    Columns[i].ConfigureColumn(worksheet.Column(i + 1));
                }
            }
        }
    }
}