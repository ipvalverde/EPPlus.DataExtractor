namespace EPPlus.DataExtractor
{
    using OfficeOpenXml;
    using System;
    using System.Collections.Generic;
    using System.Linq.Expressions;

    internal class SimpleNewableCollectionColumnDataExtractor<TRow, TCollection, TCollectionItem>
        : ISimpleCollectionColumnDataExtractor<TRow>
        where TCollection : class, ICollection<TCollectionItem>, new()
        where TRow : class, new()
        where TCollectionItem : class
    {
        private readonly string initialColumn;
        private readonly string finalColumn;
        private readonly Action<TRow, TCollection> setCollectionProperty;
        private readonly Func<TRow, TCollection> getCollection;

        public SimpleNewableCollectionColumnDataExtractor(
            Expression<Func<TRow, TCollection>> collectionPropertyExpr,
            string initialColumn,
            string finalColumn)
        {
            this.initialColumn = initialColumn;
            this.finalColumn = finalColumn;
            this.setCollectionProperty = collectionPropertyExpr.CreatePropertyValueSetterAction();
            this.getCollection = collectionPropertyExpr.Compile();
        }

        public void SetPropertyValue(TRow dataInstance, int row, ExcelRange cellRange)
        {
            var collection = this.getCollection(dataInstance) ?? new TCollection();

            foreach (var cell in cellRange[this.initialColumn + row + ":" + this.finalColumn + row])
            {
                if(!string.IsNullOrWhiteSpace(cell.Value?.ToString()))
                    collection.Add((TCollectionItem) cell.Value);
            }

            this.setCollectionProperty(dataInstance, collection);
        }
    }
}
