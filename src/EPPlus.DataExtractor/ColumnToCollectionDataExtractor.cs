namespace EPPlus.DataExtractor
{
    using OfficeOpenXml;
    using System;
    using System.Collections.Generic;
    using System.Linq.Expressions;

    internal class ColumnToCollectionDataExtractor<TRow, TCollection, TCollectionItem> : ICollectionColumnDataExtractor<TRow>
        where TCollection : class, ICollection<TCollectionItem>, new()
        where TRow : class, new()
        where TCollectionItem : class, new()
    {
        private readonly int headerRow;
        private readonly int startingColumn;
        private readonly Action<TRow, TCollection> setCollectionProperty;
        private readonly ColumnToCollectionConfiguration<TCollectionItem> columnToCollectionConfiguration;

        public ColumnToCollectionDataExtractor(
            Expression<Func<TRow, TCollection>> collectionPropertyExpr,
            int headerRow,
            string startingColumn,
            ColumnToCollectionConfiguration<TCollectionItem> columnToCollectionConfiguration)
        {
            this.headerRow = headerRow;
            this.startingColumn = SpreadsheetHelper.ConvertColumnHeaderToNumber(startingColumn);
            this.setCollectionProperty = collectionPropertyExpr.CreatePropertyValueSetterAction();
            this.columnToCollectionConfiguration = columnToCollectionConfiguration;
        }

        public void SetPropertyValue(TRow dataInstance, int row, ExcelRange cellRange)
        {
            var collection = new TCollection();

            var collectionItem = new TCollectionItem();
            var headersSetForCurrentInstance = new HashSet<string>();

            for (int column = this.startingColumn; ; column++)
            {
                var headerText = cellRange[this.headerRow, column].Text;
                if (string.IsNullOrWhiteSpace(headerText))
                    break;

                var rowDataSetter = this.columnToCollectionConfiguration.GetRowDataExtractorsByColumnHeaderText(headerText);
                if (rowDataSetter == null)
                    break;
                
                // If this header text was already set, creates a new instance
                // adding the existing one to the collection.
                if (headersSetForCurrentInstance.Contains(headerText))
                {
                    headersSetForCurrentInstance.Clear();

                    collection.Add(collectionItem);
                    collectionItem = new TCollectionItem();
                }

                headersSetForCurrentInstance.Add(headerText);

                rowDataSetter.SetPropertyValue(collectionItem, cellRange[row, column]);
            }

            if (headersSetForCurrentInstance.Count > 0)
            {
                collection.Add(collectionItem);
            }

            this.setCollectionProperty(dataInstance, collection);
        }
    }
}
