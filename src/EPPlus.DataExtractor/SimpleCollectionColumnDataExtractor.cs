namespace EPPlus.DataExtractor
{
    using OfficeOpenXml;
    using System;
    using System.Collections.Generic;
    using System.Linq.Expressions;

    internal interface ISimpleCollectionColumnDataExtractor<TRow>
        where TRow : class, new()
    {
        void SetPropertyValue(TRow dataInstance, int row, ExcelRange cellRange);
    }

    internal class SimpleCollectionColumnDataExtractor<TRow, TCollection, TCollectionItem>
        : ISimpleCollectionColumnDataExtractor<TRow>
        where TCollection : class, ICollection<TCollectionItem>, new()
        where TRow : class, new()
        where TCollectionItem : class
    {
        private readonly int headerRow;
        private readonly string initialColumn;
        private readonly string finalColumn;
        private readonly Action<TRow, TCollection> setCollectionProperty;

        public SimpleCollectionColumnDataExtractor(
            Expression<Func<TRow, TCollection>> collectionPropertyExpr,
            string initialColumn,
            string finalColumn)
        {
            this.initialColumn = initialColumn;
            this.finalColumn = finalColumn;
            this.setCollectionProperty = collectionPropertyExpr.CreatePropertyValueSetterAction();
        }

        public void SetPropertyValue(TRow dataInstance, int row, ExcelRange cellRange)
        {
            var collection = new TCollection();

            foreach (var cell in cellRange[this.initialColumn + row + ":" + this.finalColumn + row])
            {
                if(!string.IsNullOrWhiteSpace(cell.Value?.ToString()))
                    collection.Add((TCollectionItem) cell.Value);
            }

            this.setCollectionProperty(dataInstance, collection);
        }
    }
}
