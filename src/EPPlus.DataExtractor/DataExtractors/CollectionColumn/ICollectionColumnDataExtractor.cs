namespace EPPlus.DataExtractor.DataExtractors.CollectionColumn
{
    using OfficeOpenXml;

    internal interface ICollectionColumnDataExtractor<in TRow>
        where TRow : class, new()
    {
        void SetPropertyValue(TRow dataInstance, int row, ExcelRange cellRange);
    }
}
