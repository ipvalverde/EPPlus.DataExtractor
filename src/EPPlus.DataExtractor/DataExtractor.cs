using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Text.RegularExpressions;

namespace EPPlus.DataExtractor
{
    internal class CollectionColumnDataExtractor<TRow, TValue> : BaseColumnDataExtractor<TRow>
    {
        private readonly string initialColumn;
        private readonly string finalColumn;

        public CollectionColumnDataExtractor(string initialColumn, string finalColumn)
        {
            this.initialColumn = initialColumn;
            this.finalColumn = finalColumn;
        }

        public override void SetPropertyValue(TRow dataInstance, int row, ExcelRange cellRange)
        {
            foreach(var cell in cellRange[initialColumn + row + ":" + finalColumn + row])
            {

            }
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
            string initialColumn = "H";
            string finalColumn = "R";

            for (int row = fromRow; row <= toRow; row++)
            {
                //foreach (var column in this.worksheet.Cells[initialColumn+row+":"+finalColumn+row])
                {
                    var dataInstance = new TRow();

                    foreach (var propertySetter in this.propertySetters)
                        propertySetter.SetPropertyValue(dataInstance, row, this.worksheet.Cells);

                    //foreach(var collectionPropertySetter in this.collectionPropertySetters)
                    //    collectionPropertySetter.SetPropertyValue(dataInstance, row, this.worksheet.Cells);

                    yield return dataInstance;
                }
            }
        }
    }
}