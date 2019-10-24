namespace EPPlus.DataExtractor.DataExtractors.CollectionColumn
{
    using OfficeOpenXml;
    using System;
    using System.Collections.Generic;
    using System.Linq.Expressions;

    internal class NewableCollectionColumnDataExtractor<TRow, TCollection, TCollectionItem, THeadValue, TRowValue>
        : ICollectionColumnDataExtractor<TRow>
        where TCollection : class, ICollection<TCollectionItem>, new()
        where TRow : class, new()
        where TCollectionItem : class, new()
    {
        private readonly int headerRow;
        private readonly string initialColumn;
        private readonly string finalColumn;
        private readonly Action<TRow, TCollection> setCollectionProperty;
        private readonly Func<TRow, TCollection> getCollection;
        private readonly IRowDataExtractor<TCollectionItem> collectionItemHeadPropertySetter;
        private readonly IRowDataExtractor<TCollectionItem> collectionItemRowPropertySetter;

        public NewableCollectionColumnDataExtractor(
            Expression<Func<TRow, TCollection>> collectionPropertyExpr,
            Expression<Func<TCollectionItem, THeadValue>> collectionItemHeaderProperty,
            int headerRow,
            Expression<Func<TCollectionItem, TRowValue>> collectionItemRowProperty,
            string initialColumn,
            string finalColumn)
        {
            this.headerRow = headerRow;
            this.initialColumn = initialColumn;
            this.finalColumn = finalColumn;
            this.setCollectionProperty = collectionPropertyExpr.CreatePropertyValueSetterAction();
            this.getCollection = collectionPropertyExpr.Compile();
            this.collectionItemHeadPropertySetter = new RowDataExtractor<TCollectionItem, THeadValue>(collectionItemHeaderProperty);
            this.collectionItemRowPropertySetter = new RowDataExtractor<TCollectionItem, TRowValue>(collectionItemRowProperty);
        }

        public void SetPropertyValue(TRow dataInstance, int row, ExcelRange cellRange)
        {
            var collection = this.getCollection(dataInstance);
            if (collection == null)
            {
                collection = new TCollection();
                this.setCollectionProperty(dataInstance, collection);
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
