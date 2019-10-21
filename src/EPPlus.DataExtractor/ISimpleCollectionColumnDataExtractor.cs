namespace EPPlus.DataExtractor
{
    using OfficeOpenXml;

    internal interface ISimpleCollectionColumnDataExtractor<TRow>
        where TRow : class, new()
    {
        void SetPropertyValue(TRow dataInstance, int row, ExcelRange cellRange);
    }
}
