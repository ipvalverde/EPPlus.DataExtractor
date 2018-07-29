using EPPlus.DataExtractor.Data;
using OfficeOpenXml;
using System;
using System.Linq.Expressions;

namespace EPPlus.DataExtractor
{
    internal abstract class PropertyValueSetter<TModel, TValue>
        where TModel : class, new()
    {
        private readonly Action<PropertyExtractionContext, object> validateValue;
        private readonly Action<PropertyExtractionContext, TValue> validateCastedValue;

        private readonly Func<object, TValue> cellValueConverter;
        private readonly Action<TModel, TValue> setPropertyValueAction;

        protected PropertyValueSetter(Expression<Func<TModel, TValue>> propertyExpression,
            Func<object, TValue> cellValueConverter, Action<PropertyExtractionContext, object> validateValue,
            Action<PropertyExtractionContext, TValue> validateCastedValue)
        {
            this.setPropertyValueAction = propertyExpression.CreatePropertyValueSetterAction();
            this.cellValueConverter = cellValueConverter;
            this.validateValue = validateValue;
            this.validateCastedValue = validateCastedValue;
        }

        /// <summary>
        /// Sets the property value for the <paramref name="dataInstance"/>.
        /// This method also checks the validation actions, before and after casting the cell value,
        /// if one of them aborts the execution, this method will return false and it will not set the
        /// value for this property.
        /// </summary>
        /// <param name="dataInstance"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        protected bool SetPropertyValue(TModel dataInstance, ExcelRangeBase cell)
        {
            // This instance should be created only if there is at least one callback function defined.
            var context = (this.validateValue != null || this.validateCastedValue != null) ?
                new PropertyExtractionContext(new CellAddress(cell))
                :
                null;

            if (this.validateValue != null)
            {
                this.validateValue(context, cell.Value);
                if (context.Aborted)
                    return false;
            }

            TValue value;
            if (cellValueConverter == null)
                value = cell.GetValue<TValue>();
            else
                value = this.cellValueConverter(cell.Value);

            if (this.validateCastedValue != null)
            {
                this.validateCastedValue(context, value);
#pragma warning disable S2259 // Since "validateCastedValue" is not null, there is no way for "context" to be ull.
                if (context.Aborted)
                    return false;
#pragma warning restore S2259
            }

            setPropertyValueAction(dataInstance, value);
            return true;
        }
    }
}
