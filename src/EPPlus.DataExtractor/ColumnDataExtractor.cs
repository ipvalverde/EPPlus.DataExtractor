using OfficeOpenXml;
using System;

namespace EPPlus.DataExtractor
{
    internal abstract class BaseColumnDataExtractor<TRow>
    {
        public abstract void SetPropertyValue(TRow dataInstance, int row, ExcelRange cellRange);
    }

    internal class ColumnDataExtractor<TRow, TValue> : BaseColumnDataExtractor<TRow>
        where TRow : class, new()
    {
        private readonly Action<TRow, TValue> setPropertyValueAction;
        private readonly string column;

        public ColumnDataExtractor(string column, Action<TRow, TValue> setPropertyValueAction)
        {
            this.setPropertyValueAction = setPropertyValueAction;
            this.column = column;
        }

        public override void SetPropertyValue(TRow dataInstance, int row, ExcelRange cellRange)
        {
            var value = cellRange[column + row].GetValue<TValue>();
            setPropertyValueAction(dataInstance, value);
        }
    }
}
