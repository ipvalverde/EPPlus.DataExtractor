using OfficeOpenXml;
using System;
using System.Linq.Expressions;

namespace EPPlus.DataExtractor
{
    internal interface IRowDataExtractor<TModel>
    {
        void SetPropertyValue(TModel dataInstance, ExcelRangeBase cellRange);
    }

    internal class RowDataExtractor<TModel, TValue> : PropertyValueSetter<TModel, TValue>,
        IRowDataExtractor<TModel>
        where TModel : class, new()
    {
        public RowDataExtractor(Expression<Func<TModel, TValue>> propertyExpression) : base(propertyExpression)
        {}

        void IRowDataExtractor<TModel>.SetPropertyValue(TModel dataInstance, ExcelRangeBase cellRange)
        {
            base.SetPropertyValue(dataInstance, cellRange);
        }
    }
}
