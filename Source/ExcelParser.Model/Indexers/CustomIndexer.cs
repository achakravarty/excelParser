using System;

namespace ExcelParser.Model.Indexers
{
    public abstract class CustomIndexer<TIn, TOut>
    {
        protected abstract Func<TIn, TOut> Getter { get; }
        protected abstract Action<TIn, TOut> Setter { get; }

        public TOut this[TIn index]
        {
            get { return Getter(index); }
            set { Setter(index, value); }
        }
    }
}