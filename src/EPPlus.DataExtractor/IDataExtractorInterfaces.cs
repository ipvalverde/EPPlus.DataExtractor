﻿using EPPlus.DataExtractor.Data;
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
        /// <param name="setPropertyValueCallback">Optional callback that gets executed before retrieving the cell value casted to <typeparamref name="TValue"/>.
        /// The first parameter contains the cell address and a method that can abort the entire execution.
        /// The second one the value of the cell.</param>
        /// <param name="setPropertyCastedValueCallback">Optional callback that gets executed after retrieving the cell value casted to <typeparamref name="TValue"/>.
        /// The first parameter contains the cell address and a method that can abort the entire execution.
        /// The second one the value of the cell.</param>
        /// <returns></returns>
        ICollectionPropertyConfiguration<TRow> WithProperty<TValue>(Expression<Func<TRow, TValue>> propertyExpression,
            string column,
            Action<PropertyExtractionContext, object> setPropertyValueCallback = null,
            Action<PropertyExtractionContext, TValue> setPropertyCastedValueCallback = null);

        /// <summary>
        /// Maps a property from the type defined as the row model
        /// to the column identifier that has its value.
        /// </summary>
        /// <typeparam name="TValue"></typeparam>
        /// <param name="propertyExpression">Expression for the property to be mapped.</param>
        /// <param name="column">The column that contains the value to be mapped to
        /// the property defined by <paramref name="propertyExpression"/>.</param>
        /// <param name="convertDataFunc">Function that can be used to convert the cell value, which is an object
        /// to the desirable <typeparamref name="TValue"/>.</param>
        /// <param name="setPropertyValueCallback">Optional callback that gets executed prior to the <paramref name="convertDataFunc"/>.
        /// The first parameter contains the cell address and a method that can abort the entire execution.
        /// The second one the value of the cell.</param>
        /// <param name="validateCastedCellValue">Optional callback that gets executed after the <paramref name="convertDataFunc"/>.
        /// The first parameter contains the cell address and a method that can abort the entire execution.
        /// The second one the value of the cell.</param>
        /// <returns></returns>
        ICollectionPropertyConfiguration<TRow> WithProperty<TValue>(Expression<Func<TRow, TValue>> propertyExpression,
            string column, Func<object, TValue> convertDataFunc, Action<PropertyExtractionContext, object> setPropertyValueCallback = null,
            Action<PropertyExtractionContext, TValue> setPropertyCastedValueCallback = null);
    }

    public interface IGetData<TRow>
        where TRow : class, new()
    {
        /// <summary>
        /// Obtains the entities for the columns previously configured.
        /// </summary>
        /// <param name="fromRow">The initial row to start the data extraction.</param>
        /// <param name="toRow">The last row to read the data.</param>
        /// <returns>Returns an <see cref="IEnumerable{T}"/> with the data of the columns.</returns>
        IEnumerable<TRow> GetData(int fromRow, int toRow);

        /// <summary>
        /// Obtains the entities for the columns previously configured.
        /// The row that corresponds to the <paramref name="fromRow"/> will be read first,
        /// the following rows will be read until the <param name="while" /> returns false,
        /// the parameter for the predicate are the row index and the <see cref="ExcelRange" />
        /// containing the data of the worksheet.
        /// <para>
        /// The predicate works like the condition of a do-while statement.
        /// </para>
        /// </summary>
        /// <param name="while">The initial row to start the data extraction.</param>
        /// <param name="continueToNextRow">The condition that must.</param>
        /// <returns>Returns an <see cref="IEnumerable{T}"/> with the data of the columns.</returns>
        IEnumerable<TRow> GetData(int fromRow, Predicate<int> @while);
    }

    public interface IConfiguredDataExtractor<TRow> : IDataExtractor<TRow>, IGetData<TRow>
        where TRow : class, new()
    {
    }

    public interface ICollectionPropertyConfiguration<TRow> : IConfiguredDataExtractor<TRow>
        where TRow : class, new()
    {
        /// <summary>
        /// Configure a collection property from <typeparamref name="TRow"/> object
        /// that will be populated by columns data, instead of rows.
        /// </summary>
        /// <typeparam name="TCollectionItem">The type used inside a collection.
        /// This type will usually have two properties, one to hold the column header and another
        /// one for the row value.</typeparam>
        /// <typeparam name="THeaderValue">The type of the header column.</typeparam>
        /// <typeparam name="TRowValue">The type of the row value.</typeparam>
        /// <param name="propertyCollection">An expression that specifies the collection
        /// property that will hold the values.</param>
        /// <param name="headerProperty">The expression property from <typeparamref name="TCollectionItem"/>
        /// indicating the property that will be populated with the header value.</param>
        /// <param name="headerRow">The row number that contains the header data.</param>
        /// <param name="rowProperty">>The expression property from <typeparamref name="TCollectionItem"/>
        /// indicating the property that will be populated with the row value.</param>
        /// <param name="startColumn">The start of the column that will be extract to the collection.</param>
        /// <param name="endColumn">The start of the column that will be extract to the collection.</param>
        /// <returns></returns>
        ICollectionPropertyConfiguration<TRow> WithCollectionProperty<TCollectionItem, THeaderValue, TRowValue>(
            Expression<Func<TRow, List<TCollectionItem>>> propertyCollection,
            Expression<Func<TCollectionItem, THeaderValue>> headerProperty, int headerRow,
            Expression<Func<TCollectionItem, TRowValue>> rowProperty,
            string startColumn, string endColumn)
            where TCollectionItem : class, new();

        /// <summary>
        /// Configure a collection property from <typeparamref name="TRow"/> object
        /// that will be populated by columns data, instead of rows.
        /// </summary>
        /// <typeparam name="TCollectionItem">The type used inside a collection.
        /// This type will usually have two properties, one to hold the column header and another
        /// one for the row value.</typeparam>
        /// <typeparam name="THeaderValue">The type of the header column.</typeparam>
        /// <typeparam name="TRowValue">The type of the row value.</typeparam>
        /// <param name="propertyCollection">An expression that specifies the collection
        /// property that will hold the values.</param>
        /// <param name="headerProperty">The expression property from <typeparamref name="TCollectionItem"/>
        /// indicating the property that will be populated with the header value.</param>
        /// <param name="headerRow">The row number that contains the header data.</param>
        /// <param name="rowProperty">>The expression property from <typeparamref name="TCollectionItem"/>
        /// indicating the property that will be populated with the row value.</param>
        /// <param name="startColumn">The start of the column that will be extract to the collection.</param>
        /// <param name="endColumn">The start of the column that will be extract to the collection.</param>
        /// <returns></returns>
        ICollectionPropertyConfiguration<TRow> WithCollectionProperty<TCollectionItem, THeaderValue, TRowValue>(
            Expression<Func<TRow, HashSet<TCollectionItem>>> propertyCollection,
            Expression<Func<TCollectionItem, THeaderValue>> headerProperty, int headerRow,
            Expression<Func<TCollectionItem, TRowValue>> rowProperty,
            string startColumn, string endColumn)
            where TCollectionItem : class, new();

        /// <summary>
        /// Configure a collection property from <typeparamref name="TRow"/> object
        /// that will be populated by columns data, instead of rows.
        /// </summary>
        /// <typeparam name="TCollectionItem">The type used inside a collection.
        /// This type will usually have two properties, one to hold the column header and another
        /// one for the row value.</typeparam>
        /// <typeparam name="THeaderValue">The type of the header column.</typeparam>
        /// <typeparam name="TRowValue">The type of the row value.</typeparam>
        /// <param name="propertyCollection">An expression that specifies the collection
        /// property that will hold the values.</param>
        /// <param name="headerProperty">The expression property from <typeparamref name="TCollectionItem"/>
        /// indicating the property that will be populated with the header value.</param>
        /// <param name="headerRow">The row number that contains the header data.</param>
        /// <param name="rowProperty">>The expression property from <typeparamref name="TCollectionItem"/>
        /// indicating the property that will be populated with the row value.</param>
        /// <param name="startColumn">The start of the column that will be extract to the collection.</param>
        /// <param name="endColumn">The start of the column that will be extract to the collection.</param>
        /// <returns></returns>
        ICollectionPropertyConfiguration<TRow> WithCollectionProperty<TCollectionItem, THeaderValue, TRowValue>(
            Expression<Func<TRow, Collection<TCollectionItem>>> propertyCollection,
            Expression<Func<TCollectionItem, THeaderValue>> headerProperty, int headerRow,
            Expression<Func<TCollectionItem, TRowValue>> rowProperty,
            string startColumn, string endColumn)
            where TCollectionItem : class, new();

        /// <summary>
        /// Configure a collection property from <typeparamref name="TRow"/> object
        /// that will be populated by columns data, instead of rows.
        /// </summary>
        /// <typeparam name="TCollectionItem">The type used inside a collection.
        /// This type will usually have two properties, one to hold the column header and another
        /// one for the row value.</typeparam>
        /// <param name="propertyCollection">An expression that specifies the collection
        /// property that will hold the values.</param>
        /// <param name="startColumn">The start of the column that will be extract to the collection.</param>
        /// <param name="endColumn">The start of the column that will be extract to the collection.</param>
        /// <returns></returns>
        ICollectionPropertyConfiguration<TRow> WithCollectionProperty<TCollectionItem>(
            Expression<Func<TRow, List<TCollectionItem>>> propertyCollection,
            string startColumn, string endColumn) where TCollectionItem : class;

        /// <summary>
        /// Configure a collection property from <typeparamref name="TRow"/> object
        /// that will be populated by columns data, instead of rows.
        /// </summary>
        /// <typeparam name="TCollectionItem">The type used inside a collection.
        /// This type will usually have two properties, one to hold the column header and another
        /// one for the row value.</typeparam>
        /// <param name="propertyCollection">An expression that specifies the collection
        /// property that will hold the values.</param>
        /// <param name="startColumn">The start of the column that will be extract to the collection.</param>
        /// <param name="endColumn">The start of the column that will be extract to the collection.</param>
        /// <returns></returns>
        ICollectionPropertyConfiguration<TRow> WithCollectionProperty<TCollectionItem>(
            Expression<Func<TRow, HashSet<TCollectionItem>>> propertyCollection,
            string startColumn, string endColumn) where TCollectionItem : class;

        /// <summary>
        /// Configure a collection property from <typeparamref name="TRow"/> object
        /// that will be populated by columns data, instead of rows.
        /// </summary>
        /// <typeparam name="TCollectionItem">The type used inside a collection.
        /// This type will usually have two properties, one to hold the column header and another
        /// one for the row value.</typeparam>
        /// <param name="propertyCollection">An expression that specifies the collection
        /// property that will hold the values.</param>
        /// <param name="startColumn">The start of the column that will be extract to the collection.</param>
        /// <param name="endColumn">The start of the column that will be extract to the collection.</param>
        /// <returns></returns>
        ICollectionPropertyConfiguration<TRow> WithCollectionProperty<TCollectionItem>(
            Expression<Func<TRow, Collection<TCollectionItem>>> propertyCollection,
            string startColumn, string endColumn) where TCollectionItem : class;
    }
}