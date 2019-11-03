using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq.Expressions;
using System.Text;
using Xunit;

namespace EPPlus.DataExtractor.Tests
{
    public class ExpressionHelperTests
    {
        class ModelWithObservableCollection
        {
            public ObservableCollection<string> CollectionProp { get; set; }
        }

        [Fact]
        public void ValidatePropertyExpressionType_WithIncompatiblePropertyType_ShouldFail()
        {
            var collectionPropExpression = GetExpression<ModelWithObservableCollection, string>(m => m.CollectionProp);

            Assert.Throws<ArgumentException>(() =>
                ExpressionExtensions.ValidatePropertyExpressionType<ModelWithObservableCollection, string, Collection<string>>(collectionPropExpression));
        }

        class ModelWithCollection
        {
            public Collection<string> CollectionProp { get; set; }
        }

        [Fact]
        public void ValidatePropertyExpressionType_WithSameCollectionPropertyType()
        {
            var collectionPropExpression = GetExpression<ModelWithCollection, string>(m => m.CollectionProp);

            ExpressionExtensions.ValidatePropertyExpressionType<ModelWithCollection, string, Collection<string>>(collectionPropExpression);
        }

        class ModelWithList
        {
            public List<string> CollectionProp { get; set; }
        }

        [Fact]
        public void ValidatePropertyExpressionType_WithSameListPropertyType()
        {
            var collectionPropExpression = GetExpression<ModelWithList, string>(m => m.CollectionProp);

            ExpressionExtensions.ValidatePropertyExpressionType<ModelWithList, string, List<string>>(collectionPropExpression);
        }

        private static Expression<Func<TModel, Collection<TItem>>> GetExpression<TModel, TItem>(Expression<Func<TModel, Collection<TItem>>> property) => property;
        private static Expression<Func<TModel, List<TItem>>> GetExpression<TModel, TItem>(Expression<Func<TModel, List<TItem>>> property) => property;
    }
}
