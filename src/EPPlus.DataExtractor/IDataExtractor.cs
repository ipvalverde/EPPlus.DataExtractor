using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.DataExtractor
{
    /// <summary>
    /// Instance used to configure the data
    /// extraction.
    /// </summary>
    /// <typeparam name="TRow"></typeparam>
    public interface IDataExtractor<TRow>
        where TRow : class, new()
    {
        /// <summary>
        /// Maps a property from the type defined as the row model
        /// to the column identifier that has its value.
        /// </summary>
        /// <typeparam name="TValue"></typeparam>
        /// <param name="propertyExpression">Expression for the property to be mapped.</param>
        /// <param name="column">The column that contains the value to be mapped to
        /// the property defined by <paramref name="propertyExpression"/>.</param>
        /// <returns></returns>
        IConfiguredDataExtractor<TRow> WithProperty<TValue>(Expression<Func<TRow, TValue>> propertyExpression,
            string column);
    }

    public interface IConfiguredDataExtractor<TRow> : IDataExtractor<TRow>
        where TRow : class, new()
    {
        /// <summary>
        /// Obtains the entities for the rows previously configured.
        /// </summary>
        /// <param name="fromRow">The initial row to start the data extraction.</param>
        /// <param name="toRow">The last row to read the data.</param>
        /// <returns>Returns an <see cref="IEnumerable{T}"/> with the data of the columns.</returns>
        IEnumerable<TRow> GetData(int fromRow, int toRow);
    }
}
