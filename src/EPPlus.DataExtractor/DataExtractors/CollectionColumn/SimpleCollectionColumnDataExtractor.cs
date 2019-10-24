namespace EPPlus.DataExtractor.DataExtractors.CollectionColumn
{
    using OfficeOpenXml;
    using System;
    using System.Collections.Generic;

    internal class SimpleCollectionColumnDataExtractor<TRow, TCollection, TCollectionItem>
        : ISimpleCollectionColumnDataExtractor<TRow>
        where TCollection : ICollection<TCollectionItem>
        where TRow : class, new()
        where TCollectionItem : class
    {
        private readonly string initialColumn;
        private readonly string finalColumn;
        private readonly Func<TRow, TCollection> getCollectionProperty;

        public SimpleCollectionColumnDataExtractor(
            Func<TRow, TCollection> getCollectionProperty,
            string initialColumn,
            string finalColumn)
        {
            this.initialColumn = initialColumn;
            this.finalColumn = finalColumn;
            this.getCollectionProperty = getCollectionProperty;
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
                if(!string.IsNullOrWhiteSpace(cell.Value?.ToString()))
                    collection.Add((TCollectionItem) cell.Value);
            }
        }
    }
}
