using EPPlus.DataExtractor.Data;
using OfficeOpenXml;
using System;
using System.Linq.Expressions;

namespace EPPlus.DataExtractor
{
    internal interface IColumnDataExtractor<TRow>
    {
        /// <summary>
        /// Sets the property value for the <paramref name="dataInstance"/>.
        /// This method also checks the validation actions, before and after casting the cell value,
        /// if one of them aborts the execution, this method will return false and it will not set the
        /// value for this property.
        /// </summary>
        /// <param name="dataInstance"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        bool SetPropertyValue(TRow dataInstance, int row, ExcelRange cellRange);
    }

    internal class ColumnDataExtractor<TRow, TValue> : PropertyValueSetter<TRow, TValue>,
        IColumnDataExtractor<TRow>
        where TRow : class, new()
    {
        private readonly string column;

        public ColumnDataExtractor(string column, Expression<Func<TRow, TValue>> propertyExpression,
            Func<object, TValue> cellValueConverter, Action<PropertyExtractionContext, object> validateValue,
            Action<PropertyExtractionContext, TValue> validateCastedValue)
            : base(propertyExpression, cellValueConverter, validateValue, validateCastedValue)
        {
            this.column = column;
        }

        /// <summary>
        /// Sets the property value for the <paramref name="dataInstance"/>.
        /// This method also checks the validation actions, before and after casting the cell value,
        /// if one of them aborts the execution, this method will return false and it will not set the
        /// value for this property.
        /// </summary>
        /// <param name="dataInstance"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public bool SetPropertyValue(TRow dataInstance, int row, ExcelRange cellRange)
        {
            var cell = cellRange[column + row];
            return base.SetPropertyValue(dataInstance, cell);
        }
    }
}
