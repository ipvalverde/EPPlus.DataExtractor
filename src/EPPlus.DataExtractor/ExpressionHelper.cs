using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;

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

        internal static void ValidatePropertyExpressionType<TModel, TCollectionItem, TCollectionProperty>(Expression<Func<TModel, TCollectionProperty>> collectionPropertyExpr)
            where TCollectionProperty : class, ICollection<TCollectionItem>, new()
        {
            var expectedType = typeof(TCollectionProperty);
            if (!expectedType.IsGenericType)
            {
                throw new ArgumentException($"The collection property must be a generic type. Given type: {expectedType.FullName}");
            }
            expectedType = expectedType.GetGenericTypeDefinition();

            var memberExpression = GetMemberExpression();
            var propertyInfo = memberExpression.Member as PropertyInfo;
            
            var propertyType = propertyInfo.PropertyType;
            if (!propertyType.IsGenericType)
            {
                throw new ArgumentException($"The property for the given collection expression must be of generic type {expectedType}.");
            }
            var genericCollectionType = propertyType.GetGenericTypeDefinition();
            if (expectedType != genericCollectionType)
            {
                throw new ArgumentException($"The property for the given collection expression must be of type {expectedType.FullName}. Given type: {genericCollectionType.FullName}");
            }


            MemberExpression GetMemberExpression()
            {
                switch (collectionPropertyExpr.Body)
                {
                    case UnaryExpression unaryExpression:
                        if (unaryExpression.Operand is MemberExpression)
                            return (MemberExpression)unaryExpression.Operand;
                        break;
                    case MemberExpression mExpression:
                        return mExpression;
                }

                throw new ArgumentException();
            }
        }
    }
}
