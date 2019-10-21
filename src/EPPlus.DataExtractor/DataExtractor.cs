using EPPlus.DataExtractor.Data;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq.Expressions;
using System.Text.RegularExpressions;

namespace EPPlus.DataExtractor
{
    internal static class DataExtractor
    {
        // Regex used to check the column string.
        internal static readonly Regex ColumnRegex = new Regex("^[A-Za-z]+$", RegexOptions.Compiled);
    }

    internal class DataExtractor<TRow> : ICollectionPropertyConfiguration<TRow>
        where TRow : class, new()
    {
        private readonly ExcelWorksheet worksheet;
        private readonly List<IColumnDataExtractor<TRow>> propertySetters;
        private readonly List<ICollectionColumnDataExtractor<TRow>> collectionColumnSetters;
        private readonly List<ISimpleCollectionColumnDataExtractor<TRow>> simpleCollectionColumnSetters;

        internal DataExtractor(ExcelWorksheet worksheet)
        {
            this.worksheet = worksheet;
            this.propertySetters = new List<IColumnDataExtractor<TRow>>();
            this.collectionColumnSetters = new List<ICollectionColumnDataExtractor<TRow>>();
            this.simpleCollectionColumnSetters = new List<ISimpleCollectionColumnDataExtractor<TRow>>();
        }

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
        public ICollectionPropertyConfiguration<TRow> WithProperty<TValue>(Expression<Func<TRow, TValue>> propertyExpression,
            string column,
            Action<PropertyExtractionContext, object> setPropertyValueCallback = null,
            Action<PropertyExtractionContext, TValue> setPropertyCastedValueCallback = null)
        {
            return this.WithProperty(propertyExpression, column, null,
                setPropertyValueCallback, setPropertyCastedValueCallback);
        }

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
        /// <param name="setPropertyCastedValueCallback">Optional callback that gets executed after the <paramref name="convertDataFunc"/>.
        /// The first parameter contains the cell address and a method that can abort the entire execution.
        /// The second one the value of the cell.</param>
        /// <returns></returns>
        public ICollectionPropertyConfiguration<TRow> WithProperty<TValue>(Expression<Func<TRow, TValue>> propertyExpression,
            string column, Func<object, TValue> convertDataFunc, Action<PropertyExtractionContext, object> setPropertyValueCallback = null,
            Action<PropertyExtractionContext, TValue> setPropertyCastedValueCallback = null)
        {
            if (propertyExpression == null)
                throw new ArgumentNullException(nameof(propertyExpression));
            if(string.IsNullOrWhiteSpace(column))
                throw new ArgumentNullException(nameof(column));
            if (!DataExtractor.ColumnRegex.IsMatch(column))
                throw new ArgumentException("The column value must contain only letters.", nameof(column));

            propertySetters.Add(new ColumnDataExtractor<TRow, TValue>(column, propertyExpression, convertDataFunc,
                setPropertyValueCallback, setPropertyCastedValueCallback));

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
            return this.GetData(fromRow, (currentRow) => currentRow <= toRow);
        }

        /// <summary>
        /// Obtains the entities for the columns previously configured.
        /// The <paramref name="fromRow"/> indicates the initial row that will be read,
        /// the data extraction will only occur while the <param name="while" /> predicate returns true.
        /// It'll get executed receiving the row index as parameter before extracting the data of each row.
        /// </summary>
        /// <param name="fromRow">The initial row to start the data extraction.</param>
        /// <param name="@while">The condition that must.</param>
        /// <returns>Returns an <see cref="IEnumerable{T}"/> with the data of the columns.</returns>
        public IEnumerable<TRow> GetData(int fromRow, Predicate<int> @while)
        {
            for (int row = fromRow; @while(row); row++)
            {
                var dataInstance = new TRow();

                bool continueExecution = true;
                for (int index = 0; continueExecution && index < this.propertySetters.Count; index++)
                    continueExecution = this.propertySetters[index].SetPropertyValue(dataInstance, row, this.worksheet.Cells);

                if (!continueExecution)
                {
                    yield return dataInstance;
                    break;
                }

                foreach (var collectionPropertySetter in this.collectionColumnSetters)
                    collectionPropertySetter.SetPropertyValue(dataInstance, row, this.worksheet.Cells);

                foreach (var simpleCollectionColumnSetter in this.simpleCollectionColumnSetters)
                    simpleCollectionColumnSetter.SetPropertyValue(dataInstance, row, this.worksheet.Cells);

                yield return dataInstance;
            }
        }

        public ICollectionPropertyConfiguration<TRow> WithCollectionProperty<TCollectionItem>(
            Expression<Func<TRow, List<TCollectionItem>>> propertyCollection,
            string startColumn, string endColumn) where TCollectionItem : class
        {
            var collectionConfiguration = new SimpleNewableCollectionColumnDataExtractor<TRow, List<TCollectionItem>, TCollectionItem>
                (propertyCollection, startColumn, endColumn);

            this.simpleCollectionColumnSetters.Add(collectionConfiguration);

            return this;
        }

        public ICollectionPropertyConfiguration<TRow> WithCollectionProperty<TCollectionItem>(
            Expression<Func<TRow, HashSet<TCollectionItem>>> propertyCollection,
            string startColumn, string endColumn) where TCollectionItem : class
        {
            var collectionConfiguration = new SimpleNewableCollectionColumnDataExtractor<TRow, HashSet<TCollectionItem>, TCollectionItem>
                (propertyCollection, startColumn, endColumn);

            this.simpleCollectionColumnSetters.Add(collectionConfiguration);

            return this;
        }

        public ICollectionPropertyConfiguration<TRow> WithCollectionProperty<TCollectionItem>(
            Expression<Func<TRow, Collection<TCollectionItem>>> propertyCollection,
            string startColumn, string endColumn) where TCollectionItem : class
        {
            var collectionConfiguration = new SimpleNewableCollectionColumnDataExtractor<TRow, Collection<TCollectionItem>, TCollectionItem>
                (propertyCollection, startColumn, endColumn);

            this.simpleCollectionColumnSetters.Add(collectionConfiguration);

            return this;
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

        /// <summary>
        /// Configures a collection property to unpivot multiple columns to items in the collection property.
        /// Different from the overloads, this method allows for having an undefined amount of columns
        /// to be unpivoted to the collection.
        /// </summary>
        /// <param name="propertyCollection">Expression for the collection that will be populated
        /// with elements from the columns.</param>
        /// <param name="headerRow">The number of the row where the header is defined. This row will be used
        /// to search for the text of the collection columns mapping.</param>
        /// <param name="startingColumn">Indicates the column address (with letters) where this collection
        /// starts.</param>
        /// <param name="configurePropertiesAction">Action to be used to configure the columns
        /// for the collection items. Use the method
        /// <see cref="IColumnToCollectionConfiguration.WithColumn{TRowValue}(Expression{Func{TCollectionItem, TRowValue}}, string)"/>
        /// to define the mappings of the columns.
        /// </param>
        /// <returns></returns>
        public ICollectionPropertyConfigurationWithoutColumnsToCollection<TRow> WithCollectionProperty<TCollectionItem>(
            Expression<Func<TRow, List<TCollectionItem>>> propertyCollection,
            int headerRow,
            string startingColumn,
            Action<IColumnToCollectionConfiguration<TCollectionItem>> configurePropertiesAction)
            where TCollectionItem : class, new()
        {
            if (propertyCollection == null)
                throw new ArgumentNullException(nameof(propertyCollection));
            if (configurePropertiesAction == null)
                throw new ArgumentNullException(nameof(configurePropertiesAction));

            var columnToCollectionConfiguration = new ColumnToCollectionConfiguration<TCollectionItem>();
            configurePropertiesAction(columnToCollectionConfiguration);

            var columnToCollectionDataExtractor = new ColumnToCollectionDataExtractor<TRow, List<TCollectionItem>, TCollectionItem>(
                propertyCollection, headerRow, startingColumn, columnToCollectionConfiguration);
            this.collectionColumnSetters.Add(columnToCollectionDataExtractor);

            return this;
        }

        /// <summary>
        /// Configures a collection property to unpivot multiple columns to items in the collection property.
        /// Different from the overloads, this method allows for having an undefined amount of columns
        /// to be unpivoted to the collection.
        /// </summary>
        /// <param name="propertyCollection">Expression for the collection that will be populated
        /// with elements from the columns.</param>
        /// <param name="headerRow">The number of the row where the header is defined. This row will be used
        /// to search for the text of the collection columns mapping.</param>
        /// <param name="startingColumn">Indicates the column address (with letters) where this collection
        /// starts.</param>
        /// <param name="configurePropertiesAction">Action to be used to configure the columns
        /// for the collection items. Use the method
        /// <see cref="IColumnToCollectionConfiguration.WithColumn{TRowValue}(Expression{Func{TCollectionItem, TRowValue}}, string)"/>
        /// to define the mappings of the columns.
        /// </param>
        /// <returns></returns>
        public ICollectionPropertyConfigurationWithoutColumnsToCollection<TRow> WithCollectionProperty<TCollectionItem>(
            Expression<Func<TRow, HashSet<TCollectionItem>>> propertyCollection,
            int headerRow,
            string startingColumn,
            Action<IColumnToCollectionConfiguration<TCollectionItem>> configurePropertiesAction)
            where TCollectionItem : class, new()
        {
            if (propertyCollection == null)
                throw new ArgumentNullException(nameof(propertyCollection));
            if (configurePropertiesAction == null)
                throw new ArgumentNullException(nameof(configurePropertiesAction));

            var columnToCollectionConfiguration = new ColumnToCollectionConfiguration<TCollectionItem>();
            configurePropertiesAction(columnToCollectionConfiguration);

            var columnToCollectionDataExtractor = new ColumnToCollectionDataExtractor<TRow, HashSet<TCollectionItem>, TCollectionItem>(
                propertyCollection, headerRow, startingColumn, columnToCollectionConfiguration);
            this.collectionColumnSetters.Add(columnToCollectionDataExtractor);

            return this;
        }

        /// <summary>
        /// Configures a collection property to unpivot multiple columns to items in the collection property.
        /// Different from the overloads, this method allows for having an undefined amount of columns
        /// to be unpivoted to the collection.
        /// </summary>
        /// <param name="propertyCollection">Expression for the collection that will be populated
        /// with elements from the columns.</param>
        /// <param name="headerRow">The number of the row where the header is defined. This row will be used
        /// to search for the text of the collection columns mapping.</param>
        /// <param name="startingColumn">Indicates the column address (with letters) where this collection
        /// starts.</param>
        /// <param name="configurePropertiesAction">Action to be used to configure the columns
        /// for the collection items. Use the method
        /// <see cref="IColumnToCollectionConfiguration.WithColumn{TRowValue}(Expression{Func{TCollectionItem, TRowValue}}, string)"/>
        /// to define the mappings of the columns.
        /// </param>
        /// <returns></returns>
        public ICollectionPropertyConfigurationWithoutColumnsToCollection<TRow> WithCollectionProperty<TCollectionItem>(
            Expression<Func<TRow, Collection<TCollectionItem>>> propertyCollection,
            int headerRow,
            string startingColumn,
            Action<IColumnToCollectionConfiguration<TCollectionItem>> configurePropertiesAction)
            where TCollectionItem : class, new()
        {
            if (propertyCollection == null)
                throw new ArgumentNullException(nameof(propertyCollection));
            if (configurePropertiesAction == null)
                throw new ArgumentNullException(nameof(configurePropertiesAction));

            var columnToCollectionConfiguration = new ColumnToCollectionConfiguration<TCollectionItem>();
            configurePropertiesAction(columnToCollectionConfiguration);

            var columnToCollectionDataExtractor = new ColumnToCollectionDataExtractor<TRow, Collection<TCollectionItem>, TCollectionItem>(
                propertyCollection, headerRow, startingColumn, columnToCollectionConfiguration);
            this.collectionColumnSetters.Add(columnToCollectionDataExtractor);

            return this;
        }
    }
}