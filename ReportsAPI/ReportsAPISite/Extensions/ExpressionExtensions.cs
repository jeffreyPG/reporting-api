using System;
using System.Linq.Expressions;

namespace ReportsAPISite.Extensions
{
    public static class ExpressionExtensions
    {
        public static CandidateProperties<TProp> GetCandidateProperties<T, TProp>(this Expression<Func<T, TProp>> expression, T obj)
        {
            var memberExpr = (MemberExpression)expression.Body;
            var name = memberExpr.Member.Name;
            var compiledDelegate = expression.Compile();
            var value = compiledDelegate(obj);
            var candidate = new CandidateProperties<TProp>
            {
                Name = name,
                Value = value
            };

            return candidate;
        }
    }
}