namespace EPPlus.DataExtractor
{
    using System;
    using System.Collections.Generic;
    using System.Linq.Expressions;

    public class ColumnToCollectionConfiguration<TCollectionItem> : IColumnToCollectionConfiguration<TCollectionItem>
        where TCollectionItem : class, new()
    {
        private readonly IDictionary<string, IRowDataExtractor<TCollectionItem>> propertiesSettersByHeader;

        public ColumnToCollectionConfiguration()
        {
            this.propertiesSettersByHeader = new Dictionary<string, IRowDataExtractor<TCollectionItem>>();
        }

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
        public IColumnToCollectionConfiguration<TCollectionItem> WithProperty<TColumnValue>(
            Expression<Func<TCollectionItem, TColumnValue>> columnValueProperty, string columnHeader)
        {
            var dataExtractor = new RowDataExtractor<TCollectionItem, TColumnValue>(columnValueProperty);

            this.propertiesSettersByHeader.Add(columnHeader, dataExtractor);

            return this;
        }


        internal IRowDataExtractor<TCollectionItem> GetRowDataExtractorsByColumnHeaderText(string columnText)
        {
            this.propertiesSettersByHeader.TryGetValue(columnText, out var value);
            return value;
        }
    }
}
