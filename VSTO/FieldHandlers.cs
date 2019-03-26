using System;
using System.Reflection;
using Google.Apis.Calendar.v3.Data;

namespace R.GoogleOutlookSync
{
    //internal delegate bool ComparerDelegate(Event googleItem, object outlookItem);
    //internal delegate void SetterDelegate(Event googleItem, object outlookItem, Target target);

    internal class FieldHandlers
    {
        internal Func<Event, object, bool> Comparer { get; private set; }
        internal Action<Event, object, Target> Setter { get; private set; }

        internal FieldHandlers(Func<Event, object, bool> comparer, Action<Event, object, Target> setter)
        {
            this.Comparer = comparer;
            this.Setter = setter;
        }
    }
}
