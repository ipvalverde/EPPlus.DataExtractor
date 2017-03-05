using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.DataExtractor
{
    internal abstract class PropertyValueSetter<TModel, TValue>
        where TModel : class, new()
    {
        private readonly Action<TModel, TValue> setPropertyValueAction;

        internal PropertyValueSetter(Expression<Func<TModel, TValue>> propertyExpression)
        {
            this.setPropertyValueAction = propertyExpression.CreatePropertyValueSetterAction();
        }

        protected void SetPropertyValue(TModel dataInstance, ExcelRangeBase cell)
        {
            var value = cell.GetValue<TValue>();
            setPropertyValueAction(dataInstance, value);
        }
    }
}
