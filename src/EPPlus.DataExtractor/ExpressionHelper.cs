using System;
using System.Linq.Expressions;

namespace EPPlus.DataExtractor
{
    internal static class ExpressionExtensions
    {
        internal static Action<TModel, TValue> CreatePropertyValueSetterAction<TModel, TValue>(
            this Expression<Func<TModel, TValue>> propertyExpression)
        {
            var parameter = Expression.Parameter(typeof(TValue));
            var setPropActionExpression = Expression.Lambda<Action<TModel, TValue>>(
                Expression.Assign(propertyExpression.Body, parameter),
                propertyExpression.Parameters[0], parameter);

            return setPropActionExpression.Compile();
        }
    }
}
