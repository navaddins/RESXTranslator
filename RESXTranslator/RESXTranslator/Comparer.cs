using System;
using System.Collections.Generic;

namespace RESXTranslator
{
    public static class Comparer
    {
        public static Comparer<T> Create<T>(Comparison<T> comparison)
        {
            if (comparison == null) throw new ArgumentNullException("comparison");
            return new ComparisonComparer<T>(comparison);
        }

        private sealed class ComparisonComparer<T> : Comparer<T>
        {
            private readonly Comparison<T> comparison;

            public ComparisonComparer(Comparison<T> comparison)
            {
                this.comparison = comparison;
            }

            public override int Compare(T x, T y)
            {
                return comparison(x, y);
            }
        }
    }
}