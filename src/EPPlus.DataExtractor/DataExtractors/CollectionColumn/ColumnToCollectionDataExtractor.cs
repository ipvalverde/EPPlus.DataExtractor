namespace EPPlus.DataExtractor.DataExtractors.CollectionColumn
{
    using OfficeOpenXml;
    using System;
    using System.Collections.Generic;

    internal class ColumnToCollectionDataExtractor<TRow, TCollection, TCollectionItem> : ICollectionColumnDataExtractor<TRow>
        where TCollection : class, ICollection<TCollectionItem>
        where TRow : class, new()
        where TCollectionItem : class, new()
    {
        private readonly int headerRow;
        private readonly int startingColumn;
        private readonly ColumnToCollectionConfiguration<TCollectionItem> columnToCollectionConfiguration;
        private readonly Func<TRow, TCollection> getCollectionProperty;

        public ColumnToCollectionDataExtractor(
            Func<TRow, TCollection> getCollectionProperty,
            int headerRow,
            string startingColumn,
            ColumnToCollectionConfiguration<TCollectionItem> columnToCollectionConfiguration)
        {
            this.headerRow = headerRow;
            this.startingColumn = SpreadsheetHelper.ConvertColumnHeaderToNumber(startingColumn);
            this.getCollectionProperty = getCollectionProperty;
            this.columnToCollectionConfiguration = columnToCollectionConfiguration;
        }

        public void SetPropertyValue(TRow dataInstance, int row, ExcelRange cellRange)
        {
            var collection = this.getCollectionProperty(dataInstance);
            if (collection == null)
            {
                throw new InvalidOperationException(
                    $"An instance of the item {typeof(TRow).Name} returned a null collection property. Ensure the collection property getter returns an initialized instance of ICollection where data can be append to.");
            }

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
        }
    }
}
