namespace EPPlus.DataExtractor.DataExtractors.CollectionColumn
{
    using OfficeOpenXml;
    using System;
    using System.Collections.Generic;
    using System.Linq.Expressions;

    internal class CollectionColumnDataExtractor<TRow, TCollection, TCollectionItem, THeadValue, TRowValue>
        : ICollectionColumnDataExtractor<TRow>
        where TCollection : class, ICollection<TCollectionItem>
        where TRow : class, new()
        where TCollectionItem : class, new()
    {
        private readonly int headerRow;
        private readonly string initialColumn;
        private readonly string finalColumn;
        private readonly Func<TRow, TCollection> getCollectionProperty;
        private readonly IRowDataExtractor<TCollectionItem> collectionItemHeadPropertySetter;
        private readonly IRowDataExtractor<TCollectionItem> collectionItemRowPropertySetter;

        public CollectionColumnDataExtractor(
            Func<TRow, TCollection> getCollectionProperty,
            Expression<Func<TCollectionItem, THeadValue>> collectionItemHeaderProperty,
            int headerRow,
            Expression<Func<TCollectionItem, TRowValue>> collectionItemRowProperty,
            string initialColumn,
            string finalColumn)
        {
            this.headerRow = headerRow;
            this.initialColumn = initialColumn;
            this.finalColumn = finalColumn;
            this.getCollectionProperty = getCollectionProperty;
            this.collectionItemHeadPropertySetter = new RowDataExtractor<TCollectionItem, THeadValue>(collectionItemHeaderProperty);
            this.collectionItemRowPropertySetter = new RowDataExtractor<TCollectionItem, TRowValue>(collectionItemRowProperty);
        }

        public void SetPropertyValue(TRow dataInstance, int row, ExcelRange cellRange)
        {
            var collection = this.getCollectionProperty(dataInstance);
            if (collection == null)
            {
                throw new InvalidOperationException(
                    $"An instance of the item {typeof(TRow).Name} returned a null collection property. Ensure the collection property getter returns an initialized instance of ICollection where data can be append to.");
            }

            foreach (var cell in cellRange[this.initialColumn + row + ":" + this.finalColumn + row])
            {
                var collectionItem = new TCollectionItem();

                // cell here will be a single cell, always.
                // So I get the column from that cell in order to obtain the header.
                int column = cell.Start.Column;
                this.collectionItemHeadPropertySetter.SetPropertyValue(collectionItem, cellRange[this.headerRow, column]);

                this.collectionItemRowPropertySetter.SetPropertyValue(collectionItem, cell);

                collection.Add(collectionItem);
            }
        }
    }
}
