﻿using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq.Expressions;
using System.Text.RegularExpressions;

namespace EPPlus.DataExtractor
{
    internal class DataExtractor<TRow> : ICollectionPropertyConfiguration<TRow>, IConfiguredDataExtractor<TRow>
        where TRow : class, new()
    {
        // Regex use to check the column string.
        private static readonly Regex columnRegex = new Regex("^[A-Za-z]+$", RegexOptions.Compiled);

        private readonly ExcelWorksheet worksheet;
        private readonly List<IColumnDataExtractor<TRow>> propertySetters;
        private readonly List<ICollectionColumnDataExtractor<TRow>> collectionColumnSetters;

        internal DataExtractor(ExcelWorksheet worksheet)
        {
            this.worksheet = worksheet;
            this.propertySetters = new List<IColumnDataExtractor<TRow>>();
            this.collectionColumnSetters = new List<ICollectionColumnDataExtractor<TRow>>();
        }

        /// <summary>
        /// Maps a property from the type defined as the row model
        /// to the column identifier that has its value.
        /// </summary>
        /// <typeparam name="TValue"></typeparam>
        /// <param name="propertyExpression">Expression for the property to be mapped.</param>
        /// <param name="column">The column that contains the value to be mapped to
        /// the property defined by <paramref name="propertyExpression"/>.</param>
        /// <param name="convertDataFunc">Optional function that can be used to convert the cell value, which is an object
        /// to the desirable <typeparamref name="TValue"/>.</param>
        /// <returns></returns>
        public ICollectionPropertyConfiguration<TRow> WithProperty<TValue>(Expression<Func<TRow, TValue>> propertyExpression, string column,
            Func<object, TValue> convertDataFunc = null)
        {
            if (propertyExpression == null)
                throw new ArgumentNullException("propertyExpression");
            if(string.IsNullOrWhiteSpace(column))
                throw new ArgumentNullException("column");
            if (!columnRegex.IsMatch(column))
                throw new ArgumentException("The column value must contain only letters.", "column");

            propertySetters.Add(new ColumnDataExtractor<TRow, TValue>(column, propertyExpression, convertDataFunc));

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
            for (int row = fromRow; row <= toRow; row++)
            {
                var dataInstance = new TRow();

                foreach (var propertySetter in this.propertySetters)
                    propertySetter.SetPropertyValue(dataInstance, row, this.worksheet.Cells);

                foreach (var collectionPropertySetter in this.collectionColumnSetters)
                    collectionPropertySetter.SetPropertyValue(dataInstance, row, this.worksheet.Cells);

                yield return dataInstance;
            }
        }

        public ICollectionPropertyConfiguration<TRow> WithCollectionProperty<TCollectionItem, THeaderValue, TRowValue>(
            Expression<Func<TRow, List<TCollectionItem>>> propertyCollection,
            Expression<Func<TCollectionItem, THeaderValue>> headerProperty, int headerRow,
            Expression<Func<TCollectionItem, TRowValue>> rowProperty,
            string startColumn, string endColumn) where TCollectionItem : class, new()
        {
            var collectionConfiguration = new CollectionColumnDataExtractor<TRow, List<TCollectionItem>, TCollectionItem, THeaderValue, TRowValue>
                (propertyCollection, headerProperty, headerRow, rowProperty, startColumn, endColumn);

            this.collectionColumnSetters.Add(collectionConfiguration);

            return this;
        }

        public ICollectionPropertyConfiguration<TRow> WithCollectionProperty<TCollectionItem, THeaderValue, TRowValue>(
            Expression<Func<TRow, HashSet<TCollectionItem>>> propertyCollection,
            Expression<Func<TCollectionItem, THeaderValue>> headerProperty, int headerRow,
            Expression<Func<TCollectionItem, TRowValue>> rowProperty,
            string startColumn, string endColumn) where TCollectionItem : class, new()
        {
            var collectionConfiguration = new CollectionColumnDataExtractor<TRow, HashSet<TCollectionItem>, TCollectionItem, THeaderValue, TRowValue>
                (propertyCollection, headerProperty, headerRow, rowProperty, startColumn, endColumn);

            this.collectionColumnSetters.Add(collectionConfiguration);

            return this;
        }

        public ICollectionPropertyConfiguration<TRow> WithCollectionProperty<TCollectionItem, THeaderValue, TRowValue>(
            Expression<Func<TRow, Collection<TCollectionItem>>> propertyCollection,
            Expression<Func<TCollectionItem, THeaderValue>> headerProperty, int headerRow,
            Expression<Func<TCollectionItem, TRowValue>> rowProperty,
            string startColumn, string endColumn) where TCollectionItem : class, new()
        {
            var collectionConfiguration = new CollectionColumnDataExtractor<TRow, Collection<TCollectionItem>, TCollectionItem, THeaderValue, TRowValue>
                (propertyCollection, headerProperty, headerRow, rowProperty, startColumn, endColumn);

            this.collectionColumnSetters.Add(collectionConfiguration);

            return this;
        }
    }
}