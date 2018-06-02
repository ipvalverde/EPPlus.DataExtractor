using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq.Expressions;

namespace EPPlus.DataExtractor
{
    public partial interface ICollectionPropertyConfigurationWithoutColumnsToCollection<TRow> : IConfiguredDataExtractor<TRow>
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
    }

    public interface ICollectionPropertyConfiguration<TRow> : ICollectionPropertyConfigurationWithoutColumnsToCollection<TRow>
        where TRow : class, new()
    {
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
        ICollectionPropertyConfigurationWithoutColumnsToCollection<TRow> WithCollectionProperty<TCollectionItem>(
            Expression<Func<TRow, List<TCollectionItem>>> propertyCollection,
            int headerRow,
            string startingColumn,
            Action<IColumnToCollectionConfiguration<TCollectionItem>> configurePropertiesAction)
            where TCollectionItem : class, new();

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
        ICollectionPropertyConfigurationWithoutColumnsToCollection<TRow> WithCollectionProperty<TCollectionItem>(
            Expression<Func<TRow, HashSet<TCollectionItem>>> propertyCollection,
            int headerRow,
            string startingColumn,
            Action<IColumnToCollectionConfiguration<TCollectionItem>> configurePropertiesAction)
            where TCollectionItem : class, new();

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
        ICollectionPropertyConfigurationWithoutColumnsToCollection<TRow> WithCollectionProperty<TCollectionItem>(
            Expression<Func<TRow, Collection<TCollectionItem>>> propertyCollection,
            int headerRow,
            string startingColumn,
            Action<IColumnToCollectionConfiguration<TCollectionItem>> configurePropertiesAction)
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

    public interface IColumnToCollectionConfiguration<TCollectionItem>
        where TCollectionItem : class, new()
    {
        /// <summary>
        /// Configure the mapping of the collection items to the columns
        /// that will be read to extract multiple items to the collection.
        /// </summary>
        /// <typeparam name="TColumnValue">The type of the column.</typeparam>
        /// <param name="columnValueProperty">Expression mapping to the property where the value of the column will be extracted to.</param>
        /// <param name="columnHeader">The header of the column that will be used to identify if the value of the column
        /// should be mapped to this property. This it the value of the column that will be defined in the
        /// header row specified in the <see cref="ICollectionPropertyConfiguration.WithCollectionProperty{TCollectionItem}(Expression{Func{TRow, List{TCollectionItem}}}, int, Action{IColumnToCollectionConfiguration{TCollectionItem}})"/>
        /// </param>
        /// <returns></returns>
        IColumnToCollectionConfiguration<TCollectionItem> WithProperty<TColumnValue>(
            Expression<Func<TCollectionItem, TColumnValue>> columnValueProperty, string columnHeader);
    }
}