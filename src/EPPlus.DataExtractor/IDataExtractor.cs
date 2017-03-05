using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq.Expressions;

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
        ICollectionPropertyConfiguration<TRow> WithProperty<TValue>(Expression<Func<TRow, TValue>> propertyExpression,
            string column);
    }

    public interface IGetData<TRow>
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

    public interface IConfiguredDataExtractor<TRow> : IDataExtractor<TRow>, IGetData<TRow>
        where TRow : class, new()
    {
    }

    public interface ICollectionPropertyConfiguration<TRow> : IConfiguredDataExtractor<TRow>
        where TRow : class, new()
    {
        ICollectionPropertyConfiguration<TRow> WithCollectionProperty<TCollectionItem, THeaderValue, TRowValue>(
            Expression<Func<TRow, List<TCollectionItem>>> propertyCollection,
            Expression<Func<TCollectionItem, THeaderValue>> headerProperty, int headerRow,
            Expression<Func<TCollectionItem, TRowValue>> rowProperty,
            string startColumn, string endColumn)
            where TCollectionItem : class, new();

        ICollectionPropertyConfiguration<TRow> WithCollectionProperty<TCollectionItem, THeaderValue, TRowValue>(
            Expression<Func<TRow, HashSet<TCollectionItem>>> propertyCollection,
            Expression<Func<TCollectionItem, THeaderValue>> headerProperty, int headerRow,
            Expression<Func<TCollectionItem, TRowValue>> rowProperty,
            string startColumn, string endColumn)
            where TCollectionItem : class, new();

        ICollectionPropertyConfiguration<TRow> WithCollectionProperty<TCollectionItem, THeaderValue, TRowValue>(
            Expression<Func<TRow, Collection<TCollectionItem>>> propertyCollection,
            Expression<Func<TCollectionItem, THeaderValue>> headerProperty, int headerRow,
            Expression<Func<TCollectionItem, TRowValue>> rowProperty,
            string startColumn, string endColumn)
            where TCollectionItem : class, new();
    }
}