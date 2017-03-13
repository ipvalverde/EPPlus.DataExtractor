using OfficeOpenXml;
using System;
using System.Linq.Expressions;

namespace EPPlus.DataExtractor
{
    internal interface IColumnDataExtractor<TRow>
    {
        void SetPropertyValue(TRow dataInstance, int row, ExcelRange cellRange);
    }

    internal class ColumnDataExtractor<TRow, TValue> : PropertyValueSetter<TRow, TValue>,
        IColumnDataExtractor<TRow>
        where TRow : class, new()
    {
        private readonly string column;

        public ColumnDataExtractor(string column, Expression<Func<TRow, TValue>> propertyExpression,
            Func<object, TValue> cellValueConverter)
            : base(propertyExpression, cellValueConverter)
        {
            this.column = column;
        }

        public void SetPropertyValue(TRow dataInstance, int row, ExcelRange cellRange)
        {
            var cell = cellRange[column + row];
            base.SetPropertyValue(dataInstance, cell);
        }
    }
}
