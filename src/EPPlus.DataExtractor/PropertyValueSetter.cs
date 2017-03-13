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
        private readonly Func<object, TValue> cellValueConverter;
        private readonly Action<TModel, TValue> setPropertyValueAction;

        internal PropertyValueSetter(Expression<Func<TModel, TValue>> propertyExpression,
            Func<object, TValue> cellValueConverter)
        {
            this.setPropertyValueAction = propertyExpression.CreatePropertyValueSetterAction();
            this.cellValueConverter = cellValueConverter;
        }

        protected void SetPropertyValue(TModel dataInstance, ExcelRangeBase cell)
        {
            TValue value;
            if (cellValueConverter == null)
                value = cell.GetValue<TValue>();
            else
                value = this.cellValueConverter(cell.Value);

            setPropertyValueAction(dataInstance, value);
        }
    }
}
