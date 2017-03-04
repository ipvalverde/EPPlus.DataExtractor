using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Text.RegularExpressions;

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

    internal class DataExtractor<TRow> : IDataExtractor<TRow>, IConfiguredDataExtractor<TRow>
        where TRow : class, new()
    {
        // Regex use to check the column string.
        private static readonly Regex columnRegex = new Regex("^[A-Za-z]+$", RegexOptions.Compiled);

        private readonly ExcelWorksheet worksheet;
        private readonly List<BaseColumnDataExtractor<TRow>> propertySetters;

        internal DataExtractor(ExcelWorksheet worksheet)
        {
            this.worksheet = worksheet;
            this.propertySetters = new List<BaseColumnDataExtractor<TRow>>();
        }

        /// <summary>
        /// Maps a property from the type defined as the row model
        /// to the column identifier that has its value.
        /// </summary>
        /// <typeparam name="TValue"></typeparam>
        /// <param name="propertyExpression">Expression for the property to be mapped.</param>
        /// <param name="column">The column that contains the value to be mapped to
        /// the property defined by <paramref name="propertyExpression"/>.</param>
        /// <returns></returns>
        public IConfiguredDataExtractor<TRow> WithProperty<TValue>(Expression<Func<TRow, TValue>> propertyExpression, string column)
        {
            if (propertyExpression == null)
                throw new ArgumentNullException("propertyExpression");
            if(string.IsNullOrWhiteSpace(column))
                throw new ArgumentNullException("column");
            if (!columnRegex.IsMatch(column))
                throw new ArgumentException("The column value must contain only letters.", "column");

            var parameter = Expression.Parameter(typeof(TValue));
            var setPropActionExpression = Expression.Lambda<Action<TRow, TValue>>(
                Expression.Assign(propertyExpression.Body, parameter),
                propertyExpression.Parameters[0], parameter);

            propertySetters.Add(new ColumnDataExtractor<TRow, TValue>(column, setPropActionExpression.Compile()));

            return this;
        }

        /// <summary>
        /// Obtains the entities for the rows previously configured.
        /// </summary>
        /// <param name="fromRow">The initial row to start the data extraction.</param>
        /// <param name="toRow">The last row to read the data.</param>
        /// <returns>Returns an <see cref="IEnumerable{T}"/> with the data of the columns.</returns>
        public IEnumerable<TRow> GetData(int fromRow, int toRow)
        {
            for(int row = fromRow; row <= toRow; row++)
            {
                var dataInstance = new TRow();

                foreach(var propertySetter in this.propertySetters)
                    propertySetter.SetPropertyValue(dataInstance, row, this.worksheet.Cells);

                yield return dataInstance;
            }
        }
    }
}